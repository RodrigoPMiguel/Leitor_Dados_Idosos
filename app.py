import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import docx
import os
import re
import io
import time
import fitz  # PyMuPDF para converter PDF em imagem

# Configuração da página
st.set_page_config(page_title="Painel de Gestão - Defesa do Idoso SP", layout="wide")

# --- CONEXÃO COM O GOOGLE SHEETS E DRIVE ---
@st.cache_resource
def conectar_google_services():
    try:
        scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        creds_raw = st.secrets["connections"]["gsheets"]
        creds_dict = {k: v for k, v in creds_raw.items()}
        
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scope)
        
        client_sheets = gspread.authorize(credentials)
        spreadsheet_url = creds_dict["spreadsheet"]
        sh = client_sheets.open_by_url(spreadsheet_url)
        
        drive_service = build('drive', 'v3', credentials=credentials)
        
        return sh, drive_service
    except Exception as e:
        st.error(f"Erro ao conectar aos serviços do Google: {e}")
        return None, None

sh, drive_service = conectar_google_services()

# --- FUNÇÃO PARA UPLOAD DE ARQUIVOS NO GOOGLE DRIVE (TRANSFERINDO POSSE) ---
def upload_para_drive(file_bytes, filename, mime_type, folder_name="imagAppIdoso"):
    try:
        if drive_service is None:
            return None, None
            
        # Busca a pasta imagAppIdoso
        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        results = drive_service.files().list(q=query, fields="files(id, name, owners)").execute()
        folders = results.get('files', [])
        
        if not folders:
            st.error(f"Pasta '{folder_name}' não encontrada no Drive. Verifique se ela foi compartilhada com a Conta de Serviço.")
            return None, None
            
        folder_id = folders[0]['id']
            
        # Metadados do arquivo (Salvando dentro da sua pasta)
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
        
        # Envia o arquivo para a pasta
        file_uploaded = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
        
        file_id = file_uploaded.get('id')
        
        # Torna o arquivo acessível publicamente para exibir miniaturas no Streamlit
        permission = {
            'type': 'anyone',
            'role': 'reader'
        }
        drive_service.permissions().create(fileId=file_id, body=permission).execute()
        
        link_view = f"https://lh3.googleusercontent.com/d/{file_id}"
        link_drive = file_uploaded.get('webViewLink')
        
        return link_view, link_drive
    except Exception as e:
        st.error(f"Erro ao salvar arquivo no Drive: {e}")
        return None, None

# --- FUNÇÕES AUXILIARES DE BANCO DE DADOS ---
@st.cache_data(ttl=60)
def ler_aba(nome_aba):
    try:
        if sh is None: return pd.DataFrame()
        try:
            worksheet = sh.worksheet(nome_aba)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=nome_aba, rows="1000", cols="40")
            return pd.DataFrame()
            
        data = worksheet.get_all_records()
        if not data:
            all_values = worksheet.get_all_values()
            if len(all_values) > 1:
                df = pd.DataFrame(all_values[1:], columns=all_values[0])
                return df.dropna(how="all")
            return pd.DataFrame()
            
        df = pd.DataFrame(data)
        return df.dropna(how="all")
    except Exception:
        return pd.DataFrame()

def salvar_aba(df, nome_aba):
    try:
        if sh is None: return False
        try:
            worksheet = sh.worksheet(nome_aba)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=nome_aba, rows="1000", cols="40")
            
        worksheet.clear()
        
        df_clean = df.copy()
        for col in df_clean.columns:
            if pd.api.types.is_datetime64_any_dtype(df_clean[col]):
                df_clean[col] = df_clean[col].dt.strftime('%d/%m/%Y %H:%M:%S')
            else:
                df_clean[col] = df_clean[col].astype(str)
        
        df_clean = df_clean.fillna("")
        dados_lista = [df_clean.columns.values.tolist()] + df_clean.values.tolist()
        worksheet.update(dados_lista)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar na aba {nome_aba}: {e}")
        return False

# --- LISTAS AUXILIARES ---
MESES = ["Todos", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
ANOS = ["Todos", "2026", "2025", "2024", "2023"]

SENHAS_VALIDAS = ["kico21688", "res1aaa", "res2aaa"]

if "modo_edicao" not in st.session_state:
    st.session_state["modo_edicao"] = False

def df_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
    return output.getvalue()

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("### 🔒 Controle de Acesso")
    if not st.session_state["modo_edicao"]:
        st.info("👁️ **Modo Consulta (Leitura)**")
        senha_input = st.text_input("Digite a senha para editar:", type="password", key="input_senha")
        if st.button("🔓 Ativar Modo Edição"):
            if senha_input in SENHAS_VALIDAS:
                st.session_state["modo_edicao"] = True
                st.success("Modo Edição Ativado!")
                st.rerun()
            else:
                st.error("Senha incorreta!")
    else:
        st.success("✏️ **Modo Edição Ativo**")
        if st.button("🔒 Sair do Modo Edição"):
            st.session_state["modo_edicao"] = False
            st.rerun()

# --- CABEÇALHO ---
if os.path.exists("logo.png"):
    st.image("logo.png", width="stretch")

st.subheader("Painel Geral de Gestão: Políticas e Atenção ao Idoso - SP")
st.markdown("---")

# --- ABAS ---
aba_crono, aba_subpref, aba_conselheiros, aba_anotacoes, aba_registros, aba_mapa, aba_noticias, aba_sobre = st.tabs([
    "📋 Cronograma (Distritos)",
    "🏛️ Subprefeituras",
    "👥 Conselheiros Municipais",
    "📝 Anotações Importantes",
    "📌 Registros / Casas de Repouso",
    "🗺️ Fotos, Mapas e Legendas",
    "📰 Notícias e Publicações",
    "ℹ️ Sobre"
])

# --- ABA 1: CRONOGRAMA ---
with aba_crono:
    st.markdown("### Cronograma por Distrito")
    df_crono = ler_aba("cronograma_dados")
    if not df_crono.empty:
        distritos = ["Todos os Distritos"] + sorted(list(df_crono["distrito"].astype(str).unique()))
        distrito_sel = st.selectbox("🔍 Selecione o Distrito para Filtrar:", distritos)
        df_exibir = df_crono[df_crono["distrito"] == distrito_sel] if distrito_sel != "Todos os Distritos" else df_crono.copy()
            
        if st.session_state["modo_edicao"]:
            df_editado = st.data_editor(df_exibir, num_rows="dynamic", key="editor_crono", width="stretch")
            if st.button("💾 Salvar Alterações (Cronograma)"):
                if distrito_sel == "Todos os Distritos":
                    df_crono = df_editado
                else:
                    df_crono.update(df_editado)
                if salvar_aba(df_crono, "cronograma_dados"):
                    st.success("Dados salvos com sucesso!")
                    st.rerun()
        else:
            st.dataframe(df_exibir, width="stretch")
        st.download_button("📥 Baixar Tabela em Excel", data=df_para_excel(df_exibir), file_name="Cronograma_Distritos.xlsx")

# --- ABA 2: SUBPREFEITURAS ---
with aba_subpref:
    st.markdown("### Totais por Subprefeitura")
    df_sub = ler_aba("subprefeituras_dados")
    if not df_sub.empty:
        if st.session_state["modo_edicao"]:
            df_edit = st.data_editor(df_sub, num_rows="dynamic", key="editor_sub", width="stretch")
            if st.button("💾 Salvar Alterações (Subprefeituras)"):
                if salvar_aba(df_edit, "subprefeituras_dados"):
                    st.success("Subprefeituras atualizadas!")
                    st.rerun()
        else:
            st.dataframe(df_sub, width="stretch")
        st.download_button("📥 Baixar Subprefeituras em Excel", data=df_para_excel(df_sub), file_name="Subprefeituras.xlsx")

# --- ABA 3: CONSELHEIROS MUNICIPAIS (COM FORMULÁRIO DE CADASTRO + FOTO) ---
with aba_conselheiros:
    st.markdown("### 👥 Conselheiros Municipais")
    df_cons = ler_aba("conselheiros")
    
    if st.session_state["modo_edicao"]:
        with st.expander("➕ Cadastrar Novo Conselheiro", expanded=False):
            with st.form("form_novo_conselheiro", clear_on_submit=True):
                col_c1, col_c2 = st.columns(2)
                with col_c1:
                    nome_c = st.text_input("Nome Completo:*")
                    cargo_c = st.text_input("Cargo / Função:")
                    telefone_c = st.text_input("Telefone:")
                with col_c2:
                    email_c = st.text_input("E-mail:")
                    regiao_c = st.text_input("Região / Subprefeitura:")
                    foto_c = st.file_uploader("Foto do Conselheiro (JPG/PNG):", type=["jpg", "png", "jpeg"])
                obs_c = st.text_area("Observações:")
                
                btn_cadastrar_cons = st.form_submit_button("💾 Cadastrar Conselheiro")
                
                if btn_cadastrar_cons and nome_c:
                    foto_link = ""
                    if foto_c:
                        link_view, _ = upload_para_drive(foto_c.read(), foto_c.name, foto_c.type)
                        foto_link = link_view if link_view else ""
                    
                    novo_id = len(df_cons) + 1
                    novo_cons = pd.DataFrame([{
                        "id": str(novo_id),
                        "nome": nome_c,
                        "cargo": cargo_c,
                        "telefone": telefone_c,
                        "email": email_c,
                        "regiao": regiao_c,
                        "foto": foto_link,
                        "observacoes": obs_c
                    }])
                    df_cons = pd.concat([df_cons, novo_cons], ignore_index=True)
                    if salvar_aba(df_cons, "conselheiros"):
                        st.success("Conselheiro cadastrado com sucesso!")
                        st.rerun()

        st.markdown("---")
        df_edit_c = st.data_editor(df_cons, num_rows="dynamic", key="editor_cons", width="stretch")
        if st.button("💾 Salvar Alterações Tabela (Conselheiros)"):
            if salvar_aba(df_edit_c, "conselheiros"):
                st.success("Conselheiros atualizados!")
                st.rerun()
    else:
        if not df_cons.empty:
            for _, c_row in df_cons.iterrows():
                with st.container():
                    col_img, col_info = st.columns([1, 4])
                    with col_img:
                        if str(c_row.get("foto", "")).strip():
                            st.image(str(c_row["foto"]), width=120)
                        else:
                            st.markdown("👤 *Sem foto*")
                    with col_info:
                        st.markdown(f"**{c_row.get('nome', '')}** — *{c_row.get('cargo', '')}*")
                        st.write(f"📞 Telefone: {c_row.get('telefone', '')} | ✉️ E-mail: {c_row.get('email', '')}")
                        st.write(f"📍 Região: {c_row.get('regiao', '')}")
                        if str(c_row.get("observacoes", "")).strip():
                            st.caption(f"Obs: {c_row.get('observacoes', '')}")
                    st.markdown("---")

# --- ABA 4: ANOTAÇÕES IMPORTANTES ---
with aba_anotacoes:
    st.markdown("### Anotações Importantes")
    df_anot = ler_aba("anotacoes")
    if st.session_state["modo_edicao"]:
        df_edit_a = st.data_editor(df_anot, num_rows="dynamic", key="editor_anot", width="stretch")
        if st.button("💾 Salvar Anotações"):
            if salvar_aba(df_edit_a, "anotacoes"):
                st.success("Anotações salvas com sucesso!")
                st.rerun()
    else:
        st.dataframe(df_anot, width="stretch")

# --- ABA 5: REGISTROS / CASAS DE REPOUSO ---
with aba_registros:
    st.markdown("### Registros Gerais e Casas de Repouso")
    df_reg = ler_aba("registros_base")
    if st.session_state["modo_edicao"]:
        df_edit_r = st.data_editor(df_reg, num_rows="dynamic", key="editor_reg", width="stretch")
        if st.button("💾 Salvar Registros"):
            if salvar_aba(df_edit_r, "registros_base"):
                st.success("Base de registros salva!")
                st.rerun()
    else:
        st.dataframe(df_reg, width="stretch")

# --- ABA 6: FOTOS, MAPAS E LEGENDAS (COM FILTRO MÊS/ANO E ORDENAÇÃO) ---
with aba_mapa:
    st.markdown("### 🗺️ Galeria de Fotos, Mapas e Legendas")
    df_fotos = ler_aba("fotos_mapas")
    
    if st.session_state["modo_edicao"]:
        with st.expander("➕ Enviar Nova Foto ou Mapa", expanded=True):
            with st.form("form_upload_foto", clear_on_submit=True):
                tit_f = st.text_input("Título / Descrição da Imagem:*")
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    mes_f = st.selectbox("Mês de Referência:", MESES[1:])
                with col_f2:
                    ano_f = st.selectbox("Ano de Referência:", ANOS[1:])
                file_f = st.file_uploader("Selecione a Imagem (JPG, PNG):", type=["jpg", "png", "jpeg"])
                
                btn_env_f = st.form_submit_button("🚀 Enviar para o Google Drive")
                
                if btn_env_f and tit_f and file_f:
                    bytes_f = file_f.read()
                    link_v, link_d = upload_para_drive(bytes_f, file_f.name, file_f.type)
                    if link_d:
                        nova_f = pd.DataFrame([{
                            "titulo": tit_f,
                            "mes": mes_f,
                            "ano": ano_f,
                            "link_imagem": link_v,
                            "link_drive": link_d,
                            "created_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                        }])
                        df_fotos = pd.concat([df_fotos, nova_f], ignore_index=True)
                        if salvar_aba(df_fotos, "fotos_mapas"):
                            st.success("Foto/Mapa cadastrado com sucesso!")
                            st.rerun()

    # Filtros e exibição
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        filtro_mes = st.selectbox("Filtrar por Mês:", MESES, key="f_mes_foto")
    with col_m2:
        filtro_ano = st.selectbox("Filtrar por Ano:", ANOS, key="f_ano_foto")
        
    df_exibir_f = df_fotos.copy()
    if not df_exibir_f.empty:
        if filtro_mes != "Todos" and "mes" in df_exibir_f.columns:
            df_exibir_f = df_exibir_f[df_exibir_f["mes"] == filtro_mes]
        if filtro_ano != "Todos" and "ano" in df_exibir_f.columns:
            df_exibir_f = df_exibir_f[df_exibir_f["ano"] == filtro_ano]
            
        # Ordenar da mais nova para a mais antiga
        if "created_at" in df_exibir_f.columns:
            df_exibir_f = df_exibir_f.sort_values(by="created_at", ascending=False)
            
        cols_grid = st.columns(3)
        for i, (_, row_f) in enumerate(df_exibir_f.iterrows()):
            with cols_grid[i % 3]:
                if str(row_f.get("link_imagem", "")).strip():
                    st.image(str(row_f["link_imagem"]), width="stretch")
                st.markdown(f"**{row_f.get('titulo', '')}**")
                st.caption(f"📅 {row_f.get('mes', '')}/{row_f.get('ano', '')}")
                if str(row_f.get("link_drive", "")).strip():
                    st.markdown(f"[🔗 Abrir no Google Drive]({row_f['link_drive']})")
                st.markdown("---")

# --- ABA 7: NOTÍCIAS E PUBLICAÇÕES (COM PROCESSAMENTO DE PDF E ORDENAÇÃO) ---
with aba_noticias:
    st.markdown("### 📰 Notícias e Publicações (PDF / Imagem)")
    df_not = ler_aba("noticias_pdf")
    
    if st.session_state["modo_edicao"]:
        with st.expander("➕ Enviar Nova Notícia ou Boletim (PDF)", expanded=True):
            with st.form("form_noticia_pdf", clear_on_submit=True):
                tit_n = st.text_input("Título da Notícia / Boletim:*")
                col_n1, col_n2 = st.columns(2)
                with col_n1:
                    mes_n = st.selectbox("Mês de Referência:", MESES[1:])
                with col_n2:
                    ano_n = st.selectbox("Ano de Referência:", ANOS[1:])
                file_n = st.file_uploader("Selecione o arquivo (PDF, JPG, PNG):", type=["pdf", "jpg", "png", "jpeg"])
                
                btn_env_n = st.form_submit_button("🚀 Processar e Enviar Notícia")
                
                if btn_env_n and tit_n and file_n:
                    bytes_n = file_n.read()
                    capa_link, link_drive = None, None
                    
                    if file_n.name.lower().endswith('.pdf'):
                        try:
                            doc = fitz.open(stream=bytes_n, filetype="pdf")
                            page = doc[0]
                            pix = page.get_pixmap()
                            capa_bytes = pix.tobytes("png")
                            capa_link, _ = upload_para_drive(capa_bytes, f"capa_{file_n.name}.png", "image/png")
                            _, link_drive = upload_para_drive(bytes_n, file_n.name, "application/pdf")
                        except Exception as e:
                            st.error(f"Erro ao converter PDF: {e}")
                    else:
                        capa_link, link_drive = upload_para_drive(bytes_n, file_n.name, file_n.type)
                        
                    if link_drive:
                        nova_n = pd.DataFrame([{
                            "titulo": tit_n,
                            "mes": mes_n,
                            "ano": ano_n,
                            "link_capa": capa_link if capa_link else "",
                            "link_drive": link_drive,
                            "created_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                        }])
                        df_not = pd.concat([df_not, nova_n], ignore_index=True)
                        if salvar_aba(df_not, "noticias_pdf"):
                            st.success("Notícia cadastrada com sucesso!")
                            st.rerun()

    # Filtros e exibição ordenada
    col_fn1, col_fn2 = st.columns(2)
    with col_fn1:
        filtro_mes_n = st.selectbox("Filtrar por Mês:", MESES, key="f_mes_noticia")
    with col_fn2:
        filtro_ano_n = st.selectbox("Filtrar por Ano:", ANOS, key="f_ano_noticia")
        
    df_exibir_n = df_not.copy()
    if not df_exibir_n.empty:
        if filtro_mes_n != "Todos" and "mes" in df_exibir_n.columns:
            df_exibir_n = df_exibir_n[df_exibir_n["mes"] == filtro_mes_n]
        if filtro_ano_n != "Todos" and "ano" in df_exibir_n.columns:
            df_exibir_n = df_exibir_n[df_exibir_n["ano"] == filtro_ano_n]
            
        if "created_at" in df_exibir_n.columns:
            df_exibir_n = df_exibir_n.sort_values(by="created_at", ascending=False)
            
        cols_n_grid = st.columns(2)
        for idx_n, (_, row_n) in enumerate(df_exibir_n.iterrows()):
            with cols_n_grid[idx_n % 2]:
                st.subheader(str(row_n.get("titulo", "")))
                st.caption(f"📅 Referência: {row_n.get('mes', '')}/{row_n.get('ano', '')}")
                if str(row_n.get("link_capa", "")).strip():
                    st.image(str(row_n["link_capa"]), width="stretch")
                if str(row_n.get("link_drive", "")).strip():
                    st.markdown(f"📄 [Abrir Documento Completo / PDF]({row_n['link_drive']})")
                st.markdown("---")

# --- ABA 8: SOBRE ---
with aba_sobre:
    st.markdown("### Sobre o Sistema")
    st.markdown("""
    **Painel Geral de Gestão - Políticas e Atenção ao Idoso (SP)**
    
    - **Banco de Dados:** Google Sheets
    - **Mídia e Documentos:** Google Drive (`imagAppIdoso`)
    - **Modo de Edição:** Protegido por Senha
    """)
