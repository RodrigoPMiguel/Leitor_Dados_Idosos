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
from PIL import Image

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

# --- FUNÇÃO PARA SALVAR ARQUIVOS NO GOOGLE DRIVE ---
def upload_para_drive(file_bytes, filename, mime_type, folder_name="imagAppIdoso"):
    try:
        if drive_service is None:
            return None, None
            
        # Busca a pasta imagAppIdoso
        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        results = drive_service.files().list(q=query, fields="files(id, name)").execute()
        folders = results.get('files', [])
        
        if not folders:
            # Se não achar a pasta, cria automaticamente
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
            folder_id = folder.get('id')
        else:
            folder_id = folders[0]['id']
            
        # Metadados do arquivo
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
        file_uploaded = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink, webContentLink'
        ).execute()
        
        file_id = file_uploaded.get('id')
        
        # Torna o arquivo legível publicamente para exibir no Streamlit
        permission = {'type': 'anyone', 'role': 'reader'}
        drive_service.permissions().create(fileId=file_id, body=permission).execute()
        
        link_view = f"https://lh3.googleusercontent.com/d/{file_id}"
        link_download = file_uploaded.get('webViewLink')
        
        return link_view, link_download
    except Exception as e:
        st.error(f"Erro no envio para o Google Drive: {e}")
        return None, None

# --- FUNÇÕES AUXILIARES COM CACHE DE LEITURA ---
@st.cache_data(ttl=300)
def ler_aba(nome_aba):
    try:
        if sh is None:
            return pd.DataFrame()
        
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
        if sh is None:
            return False
            
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

# --- CARGA INICIAL DE DADOS ---
def inicializar_planilha_se_vazia():
    df_cons = ler_aba("conselheiros")
    if df_cons.empty or (len(df_cons) == 1 and str(df_cons.iloc[0, 0]).strip() in ["1", "id"]):
        dados_cons = pd.DataFrame([
            {"id": "1", "nome": "Francisco Miguel Filho", "cargo": "Conselheiro", "telefone": "", "email": "", "regiao": "São Paulo", "observacoes": ""},
            {"id": "2", "nome": "Vanessa Nassif", "cargo": "Conselheira", "telefone": "", "email": "", "regiao": "São Paulo", "observacoes": ""}
        ])
        salvar_aba(dados_cons, "conselheiros")
        time.sleep(1)

    df_crono = ler_aba("cronograma_dados")
    if (df_crono.empty or (len(df_crono) == 1 and str(df_crono.iloc[0, 0]).strip() in ["1", "id"])) and os.path.exists("cronograma.docx"):
        try:
            doc = docx.Document("cronograma.docx")
            if len(doc.tables) > 0:
                rows_list = []
                table = doc.tables[0]
                for row in table.rows[3:]:
                    txts = [cell.text.replace('\n', ' ').strip() for cell in row.cells]
                    if not txts or len(txts) < 18: continue
                    distrito = txts[0]
                    if not distrito or distrito.upper().startswith("ZONA") or distrito.upper().startswith("DISTRITOS"): continue
                    
                    def sep(val, q=2):
                        if not val or val == '-' or val == '0-': return ['0'] * q
                        p = [x.strip() for x in re.split(r'[-–—]', str(val)) if x.strip()]
                        while len(p) < q: p.append('0')
                        return p[:q]
                    
                    cras, creas = sep(txts[7], 2)
                    ubs, ama = sep(txts[13], 2)
                    ursi, upa = sep(txts[16], 2)
                    cdi, pai = sep(txts[17], 2)
                    ceu, cdc, ccint = sep(txts[18], 3)
                    col18 = txts[19] if len(txts) > 19 else ""
                    ilpi_2, cdia, nci_2 = sep(col18, 3)
                    col19 = txts[20] if len(txts) > 20 else ""

                    rows_list.append({
                        "distrito": distrito, "ano_criacao": txts[1], "no_mapa": txts[2], "pop_total": txts[3],
                        "pop_masc": txts[4], "pop_fem": txts[5], "subpref_sn": txts[6], "cras": cras, "creas": creas,
                        "caei": txts[8].replace('-',''), "nci": txts[9].replace('-',''), "ilpi": txts[10].replace('-',''),
                        "bpc": txts[11], "rank_vun": txts[12], "ubs": ubs, "ama": ama, "idrpg": txts[14], "emad": txts[15],
                        "ursi": ursi, "upa": upa, "cdi": cdi, "pai": pai, "ceu": ceu, "cdc": cdc, "ccint": ccint,
                        "ilpi_2setor": ilpi_2, "cdia": cdia, "nci_2setor": nci_2, "outros_projetos": col19
                    })
                if rows_list:
                    salvar_aba(pd.DataFrame(rows_list), "cronograma_dados")
                    time.sleep(1)
        except Exception:
            pass

    df_sub = ler_aba("subprefeituras_dados")
    if (df_sub.empty or (len(df_sub) == 1 and str(df_sub.iloc[0, 0]).strip() in ["1", "id"])) and os.path.exists("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx"):
        try:
            df_t = pd.read_excel("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx", sheet_name="TOTAIS").dropna(how="all")
            if 'Unnamed: 0' in df_t.columns: df_t = df_t.drop(columns=['Unnamed: 0'])
            salvar_aba(df_t, "subprefeituras_dados")
            time.sleep(1)
        except Exception:
            pass

    df_reg = ler_aba("registros_base")
    if (df_reg.empty or (len(df_reg) == 1 and str(df_reg.iloc[0, 0]).strip() in ["1", "id"])) and os.path.exists("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx"):
        try:
            df_b = pd.read_excel("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx", sheet_name="Base de Dados", header=1).dropna(how="all")
            salvar_aba(df_b, "registros_base")
            time.sleep(1)
        except Exception:
            pass

inicializar_planilha_se_vazia()

# --- SENHAS DE ACESSO ---
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

# --- ABAS DA APLICAÇÃO ---
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

# --- ABA 1: CRONOGRAMA POR DISTRITO ---
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
                df_crono = df_editado if distrito_sel == "Todos os Distritos" else df_crono.update(df_editado)
                if salvar_aba(df_crono, "cronograma_dados"):
                    st.success("Dados salvos com sucesso!")
                    st.rerun()
        else:
            st.info("ℹ️ Tabela em modo de leitura.")
            st.dataframe(df_exibir, width="stretch")
        st.download_button("📥 Baixar Tabela em Excel", data=df_para_excel(df_exibir), file_name="Cronograma_Distritos.xlsx")
    else:
        st.warning("Nenhum dado encontrado.")

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

# --- ABA 3: CONSELHEIROS MUNICIPAIS ---
with aba_conselheiros:
    st.markdown("### Conselheiros Municipais")
    df_cons = ler_aba("conselheiros")
    if not df_cons.empty:
        if st.session_state["modo_edicao"]:
            df_edit_c = st.data_editor(df_cons, num_rows="dynamic", key="editor_cons", width="stretch")
            if st.button("💾 Salvar Alterações (Conselheiros)"):
                if salvar_aba(df_edit_c, "conselheiros"):
                    st.success("Conselheiros atualizados!")
                    st.rerun()
        else:
            st.dataframe(df_cons, width="stretch")
        st.download_button("📥 Baixar Conselheiros em Excel", data=df_para_excel(df_cons), file_name="Conselheiros.xlsx")

# --- ABA 4: ANOTAÇÕES IMPORTANTES ---
with aba_anotacoes:
    st.markdown("### Anotações Importantes")
    df_anot = ler_aba("anotacoes")
    if df_anot.empty:
        df_anot = pd.DataFrame([{"id": "1", "titulo": "Anotação Inicial", "conteudo": "Digite aqui observações relevantes sobre a gestão.", "data": "2026-08-03"}])
    
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
    if not df_reg.empty:
        if st.session_state["modo_edicao"]:
            df_edit_r = st.data_editor(df_reg, num_rows="dynamic", key="editor_reg", width="stretch")
            if st.button("💾 Salvar Registros"):
                if salvar_aba(df_edit_r, "registros_base"):
                    st.success("Base de registros salva!")
                    st.rerun()
        else:
            st.dataframe(df_reg, width="stretch")
        st.download_button("📥 Baixar Registros em Excel", data=df_para_excel(df_reg), file_name="Registros_Base.xlsx")

# --- ABA 6: FOTOS, MAPAS E LEGENDAS ---
with aba_mapa:
    st.markdown("### 🗺️ Galeria de Fotos, Mapas e Legendas")
    df_fotos = ler_aba("fotos_mapas")
    
    if st.session_state["modo_edicao"]:
        st.markdown("#### 📤 Upload de Nova Foto ou Mapa")
        titulo_foto = st.text_input("Título / Descrição da Imagem:")
        arquivo_foto = st.file_uploader("Selecione uma Imagem (JPG, PNG)", type=["jpg", "jpeg", "png"], key="up_foto")
        
        if st.button("🚀 Enviar Foto para o Google Drive") and arquivo_foto and titulo_foto:
            bytes_foto = arquivo_foto.read()
            link_img, link_drive = upload_para_drive(bytes_foto, arquivo_foto.name, arquivo_foto.type)
            
            if link_img:
                novo_rec = pd.DataFrame([{
                    "titulo": titulo_foto,
                    "link_imagem": link_img,
                    "link_drive": link_drive,
                    "data": pd.Timestamp.now().strftime("%d/%m/%Y")
                }])
                df_fotos = pd.concat([df_fotos, novo_rec], ignore_index=True)
                salvar_aba(df_fotos, "fotos_mapas")
                st.success("Foto enviada e vinculada com sucesso!")
                st.rerun()

    if not df_fotos.empty:
        cols = st.columns(3)
        for idx, row in df_fotos.iterrows():
            with cols[idx % 3]:
                if str(row.get("link_imagem", "")).strip():
                    st.image(str(row["link_imagem"]), caption=str(row.get("titulo", "Sem título")), width="stretch")
                st.markdown(f"**{row.get('titulo', '')}**")
                if str(row.get("link_drive", "")).strip():
                    st.markdown(f"[🔗 Ver no Google Drive]({row['link_drive']})")
    else:
        st.info("Nenhuma foto ou mapa cadastrado na galeria.")

# --- ABA 7: NOTÍCIAS E PUBLICAÇÕES (COM PROCESSAMENTO DE PDF) ---
with aba_noticias:
    st.markdown("### 📰 Notícias e Boletins (PDF / Imagem)")
    df_noticias = ler_aba("noticias_pdf")
    
    if st.session_state["modo_edicao"]:
        st.markdown("#### 📤 Enviar Nova Notícia ou Boletim em PDF")
        titulo_noticia = st.text_input("Título da Notícia / Boletim:")
        arquivo_pdf = st.file_uploader("Selecione o arquivo PDF ou Imagem", type=["pdf", "png", "jpg", "jpeg"], key="up_noticia")
        
        if st.button("🚀 Processar e Enviar Notícia") and arquivo_pdf and titulo_noticia:
            bytes_file = arquivo_pdf.read()
            
            # Se for PDF, extrai a 1ª página como imagem de capa (fitz)
            if arquivo_pdf.name.lower().endswith('.pdf'):
                try:
                    doc = fitz.open(stream=bytes_file, filetype="pdf")
                    page = doc[0]
                    pix = page.get_pixmap()
                    img_bytes = pix.tobytes("png")
                    
                    # Upload da imagem da capa
                    capa_link, _ = upload_para_drive(img_bytes, f"capa_{arquivo_pdf.name}.png", "image/png")
                    # Upload do PDF completo
                    pdf_link, link_drive = upload_para_drive(bytes_file, arquivo_pdf.name, "application/pdf")
                except Exception as e:
                    st.error(f"Erro ao extrair capa do PDF: {e}")
                    capa_link, link_drive = None, None
            else:
                capa_link, link_drive = upload_para_drive(bytes_file, arquivo_pdf.name, arquivo_pdf.type)
            
            if link_drive:
                nova_noticia = pd.DataFrame([{
                    "titulo": titulo_noticia,
                    "link_capa": capa_link if capa_link else "",
                    "link_drive": link_drive,
                    "data": pd.Timestamp.now().strftime("%d/%m/%Y")
                }])
                df_noticias = pd.concat([df_noticias, nova_noticia], ignore_index=True)
                salvar_aba(df_noticias, "noticias_pdf")
                st.success("Notícia/Boletim cadastrado com sucesso!")
                st.rerun()

    if not df_noticias.empty:
        cols_n = st.columns(2)
        for idx_n, row_n in df_noticias.iterrows():
            with cols_n[idx_n % 2]:
                st.subheader(str(row_n.get("titulo", "Sem Título")))
                if str(row_n.get("link_capa", "")).strip():
                    st.image(str(row_n["link_capa"]), width="stretch")
                st.caption(f"Data: {row_n.get('data', '')}")
                if str(row_n.get("link_drive", "")).strip():
                    st.markdown(f"📄 [Abrir Documento Completo / PDF]({row_n['link_drive']})")
                st.markdown("---")
    else:
        st.info("Nenhuma notícia cadastrada.")

# --- ABA 8: SOBRE ---
with aba_sobre:
    st.markdown("### Sobre o Sistema")
    st.markdown("""
    **Painel Geral de Gestão - Políticas e Atenção ao Idoso (SP)**
    
    - **Banco de Dados:** Google Sheets
    - **Armazenamento de Mídia:** Google Drive (Pasta `imagAppIdoso`)
    - **Modo de Acesso:** Leitura Aberta / Edição Protegida por Senha
    """)
