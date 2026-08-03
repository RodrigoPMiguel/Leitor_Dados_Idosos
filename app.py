import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import docx
import os
import re
import io
import time
import requests
import fitz  # PyMuPDF
from PIL import Image

# Configuração da página
st.set_page_config(page_title="Painel de Gestão - Defesa do Idoso SP", layout="wide")

# --- CONEXÃO COM O GOOGLE SHEETS ---
@st.cache_resource
def conectar_gsheets():
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
        client = gspread.authorize(credentials)
        
        spreadsheet_url = creds_dict["spreadsheet"]
        sh = client.open_by_url(spreadsheet_url)
        return sh
    except Exception as e:
        st.error(f"Erro ao conectar ao Google Sheets: {e}")
        return None

sh = conectar_gsheets()

# --- FUNÇÃO DE UPLOAD DE IMAGEM VIA IMGBB (GERA URL CURTA COMPATÍVEL COM GOOGLE SHEETS) ---
def upload_imagem_imgbb(file_bytes):
    try:
        # Recupera a chave da API do Secrets
        api_key = st.secrets.get("IMGBB_API_KEY", "")
        if not api_key:
            st.error("Chave 'IMGBB_API_KEY' não encontrada no Secrets do Streamlit!")
            return None

        # Otimiza a imagem com PIL antes do upload
        img = Image.open(io.BytesIO(file_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        img.thumbnail((1200, 1200), Image.LANCZOS)
        
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=80, optimize=True)
        img_bytes_compressed = buffer.getvalue()

        # Envia para a API do ImgBB
        url = "https://api.imgbb.com/1/upload"
        payload = {"key": api_key}
        files = {"image": img_bytes_compressed}
        
        response = requests.post(url, data=payload, files=files)
        data = response.json()
        
        if response.status_code == 200 and data.get("success"):
            return data["data"]["url"]  # Retorna a URL curta da imagem
        else:
            st.error(f"Erro no serviço de hospedagem ImgBB: {data.get('error', {}).get('message', 'Falha no envio')}")
            return None
    except Exception as e:
        st.error(f"Erro ao fazer upload da imagem: {e}")
        return None

# --- FUNÇÕES AUXILIARES COM CACHE DE LEITURA ---
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

# --- CARGA E CRUZAMENTO INICIAL DOS DADOS ---
def inicializar_planilha_se_vazia():
    df_cons = ler_aba("conselheiros")
    if df_cons.empty or (len(df_cons) == 1 and str(df_cons.iloc[0, 0]).strip() in ["1", "id"]):
        dados_cons = pd.DataFrame([
            {"id": "1", "nome": "Francisco Miguel Filho", "cargo": "Conselheiro", "telefone": "", "email": "", "regiao": "São Paulo", "foto": "", "observacoes": ""},
            {"id": "2", "nome": "Vanessa Nassif", "cargo": "Conselheira", "telefone": "", "email": "", "regiao": "São Paulo", "foto": "", "observacoes": ""}
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

# --- ABA 3: CONSELHEIROS MUNICIPAIS ---
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
                    url_foto = ""
                    if foto_c:
                        url_foto = upload_imagem_imgbb(foto_c.read())
                    
                    novo_id = len(df_cons) + 1
                    novo_cons = pd.DataFrame([{
                        "id": str(novo_id),
                        "nome": nome_c,
                        "cargo": cargo_c,
                        "telefone": telefone_c,
                        "email": email_c,
                        "regiao": regiao_c,
                        "foto": url_foto if url_foto else "",
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

# --- ABA 6: FOTOS, MAPAS E LEGENDAS ---
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
                
                btn_env_f = st.form_submit_button("🚀 Salvar Foto na Galeria")
                
                if btn_env_f and tit_f and file_f:
                    bytes_f = file_f.read()
                    url_img = upload_imagem_imgbb(bytes_f)
                    
                    if url_img:
                        nova_f = pd.DataFrame([{
                            "titulo": tit_f,
                            "mes": mes_f,
                            "ano": ano_f,
                            "link_imagem": url_img,
                            "created_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                        }])
                        df_fotos = pd.concat([df_fotos, nova_f], ignore_index=True)
                        if salvar_aba(df_fotos, "fotos_mapas"):
                            st.success("Foto/Mapa cadastrado com sucesso!")
                            st.rerun()

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
            
        if "created_at" in df_exibir_f.columns:
            df_exibir_f = df_exibir_f.sort_values(by="created_at", ascending=False)
            
        cols_grid = st.columns(3)
        for i, (_, row_f) in enumerate(df_exibir_f.iterrows()):
            with cols_grid[i % 3]:
                if str(row_f.get("link_imagem", "")).strip():
                    st.image(str(row_f["link_imagem"]), width="stretch")
                st.markdown(f"**{row_f.get('titulo', '')}**")
                st.caption(f"📅 Referência: {row_f.get('mes', '')}/{row_f.get('ano', '')}")
                st.markdown("---")

# --- ABA 7: NOTÍCIAS E PUBLICAÇÕES (COM CONVERSÃO DE PDF EM CAPA) ---
with aba_noticias:
    st.markdown("### 📰 Notícias e Publicações (PDF / Imagem)")
    df_not = ler_aba("noticias_pdf")
    
    if st.session_state["modo_edicao"]:
        with st.expander("➕ Enviar Nova Notícia ou Boletim", expanded=True):
            with st.form("form_noticia_pdf", clear_on_submit=True):
                tit_n = st.text_input("Título da Notícia / Boletim:*")
                col_n1, col_n2 = st.columns(2)
                with col_n1:
                    mes_n = st.selectbox("Mês de Referência:", MESES[1:])
                with col_n2:
                    ano_n = st.selectbox("Ano de Referência:", ANOS[1:])
                file_n = st.file_uploader("Selecione o arquivo (PDF, JPG, PNG):", type=["pdf", "jpg", "png", "jpeg"])
                
                btn_env_n = st.form_submit_button("🚀 Salvar Publicação")
                
                if btn_env_n and tit_n and file_n:
                    bytes_n = file_n.read()
                    url_capa = None
                    
                    if file_n.name.lower().endswith('.pdf'):
                        try:
                            doc = fitz.open(stream=bytes_n, filetype="pdf")
                            page = doc[0]
                            pix = page.get_pixmap(dpi=120)
                            capa_bytes = pix.tobytes("png")
                            url_capa = upload_imagem_imgbb(capa_bytes)
                        except Exception as e:
                            st.error(f"Erro ao converter a 1ª página do PDF: {e}")
                    else:
                        url_capa = upload_imagem_imgbb(bytes_n)
                        
                    if url_capa:
                        nova_n = pd.DataFrame([{
                            "titulo": tit_n,
                            "mes": mes_n,
                            "ano": ano_n,
                            "link_capa": url_capa,
                            "created_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                        }])
                        df_not = pd.concat([df_not, nova_n], ignore_index=True)
                        if salvar_aba(df_not, "noticias_pdf"):
                            st.success("Notícia cadastrada com sucesso!")
                            st.rerun()

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
                st.markdown("---")

# --- ABA 8: SOBRE ---
with aba_sobre:
    st.markdown("### Sobre o Sistema")
    st.markdown("""
    **Painel Geral de Gestão - Políticas e Atenção ao Idoso (SP)**
    
    - **Banco de Dados:** Google Sheets
    - **Armazenamento Mídia:** ImgBB Cloud API
    - **Modo de Edição:** Protegido por Senha
    """)
