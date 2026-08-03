import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import docx
import os
import re
import io
import time
import fitz  # PyMuPDF para converter PDF em imagem

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

# --- FUNÇÕES AUXILIARES COM CACHE DE LEITURA (EVITA EXCEDER COTA) ---
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
        
        # Converte datas e outros tipos não serializáveis em string
        df_clean = df.copy()
        for col in df_clean.columns:
            if pd.api.types.is_datetime64_any_dtype(df_clean[col]):
                df_clean[col] = df_clean[col].dt.strftime('%d/%m/%Y %H:%M:%S')
            else:
                df_clean[col] = df_clean[col].astype(str)
        
        df_clean = df_clean.fillna("")
        dados_lista = [df_clean.columns.values.tolist()] + df_clean.values.tolist()
        worksheet.update(dados_lista)
        
        # Limpa o cache de leitura para recarregar os dados atualizados
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar na aba {nome_aba}: {e}")
        return False

# --- CARGA E CRUZAMENTO INICIAL DOS DADOS ---
def inicializar_planilha_se_vazia():
    # 1. Conselheiros Iniciais
    df_cons = ler_aba("conselheiros")
    if df_cons.empty or (len(df_cons) == 1 and str(df_cons.iloc[0, 0]).strip() in ["1", "id"]):
        dados_cons = pd.DataFrame([
            {"id": "1", "nome": "Francisco Miguel Filho", "cargo": "Conselheiro", "telefone": "", "email": "", "regiao": "São Paulo", "observacoes": ""},
            {"id": "2", "nome": "Vanessa Nassif", "cargo": "Conselheira", "telefone": "", "email": "", "regiao": "São Paulo", "observacoes": ""}
        ])
        salvar_aba(dados_cons, "conselheiros")
        time.sleep(1)

    # 2. Cronograma Desmembrado (Word)
    df_crono = ler_aba("cronograma_dados")
    if (df_crono.empty or (len(df_crono) == 1 and str(df_crono.iloc[0, 0]).strip() in ["1", "id"])) and os.path.exists("cronograma.docx"):
        try:
            doc = docx.Document("cronograma.docx")
            if len(doc.tables) > 0:
                rows_list = []
                table = doc.tables[0]
                for row in table.rows[3:]:
                    txts = [cell.text.replace('\n', ' ').strip() for cell in row.cells]
                    if not txts or len(txts) < 18:
                        continue
                    distrito = txts[0]
                    if not distrito or distrito.upper().startswith("ZONA") or distrito.upper().startswith("DISTRITOS"):
                        continue
                    
                    def sep(val, q=2):
                        if not val or val == '-' or val == '0-':
                            return ['0'] * q
                        p = [x.strip() for x in re.split(r'[-–—]', str(val)) if x.strip()]
                        while len(p) < q:
                            p.append('0')
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

    # 3. Subprefeituras Totais (Excel)
    df_sub = ler_aba("subprefeituras_dados")
    if (df_sub.empty or (len(df_sub) == 1 and str(df_sub.iloc[0, 0]).strip() in ["1", "id"])) and os.path.exists("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx"):
        try:
            df_t = pd.read_excel("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx", sheet_name="TOTAIS").dropna(how="all")
            if 'Unnamed: 0' in df_t.columns:
                df_t = df_t.drop(columns=['Unnamed: 0'])
            salvar_aba(df_t, "subprefeituras_dados")
            time.sleep(1)
        except Exception:
            pass

    # 4. Registros Base (Excel)
    df_reg = ler_aba("registros_base")
    if (df_reg.empty or (len(df_reg) == 1 and str(df_reg.iloc[0, 0]).strip() in ["1", "id"])) and os.path.exists("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx"):
        try:
            df_b = pd.read_excel("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx", sheet_name="Base de Dados", header=1).dropna(how="all")
            salvar_aba(df_b, "registros_base")
            time.sleep(1)
        except Exception:
            pass

inicializar_planilha_se_vazia()

# --- SENHAS DE ACESSO AO MODO EDIÇÃO ---
SENHAS_VALIDAS = ["kico21688", "res1aaa", "res2aaa"]

if "modo_edicao" not in st.session_state:
    st.session_state["modo_edicao"] = False

# --- FUNÇÃO AUXILIAR PARA GERAR RELATÓRIOS EM EXCEL ---
def df_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
    return output.getvalue()

MESES_MAPA = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
    "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
    "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
}

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
        
        if distrito_sel != "Todos os Distritos":
            df_exibir = df_crono[df_crono["distrito"] == distrito_sel]
        else:
            df_exibir = df_crono.copy()
            
        if st.session_state["modo_edicao"]:
            df_editado = st.data_editor(df_exibir, num_rows="dynamic", key="editor_crono", width="stretch")
            if st.button("💾 Salvar Alterações (Cronograma)"):
                if distrito_sel != "Todos os Distritos":
                    df_crono.update(df_editado)
                else:
                    df_crono = df_editado
                if salvar_aba(df_crono, "cronograma_dados"):
                    st.success("Dados salvos com sucesso!")
                    st.rerun()
        else:
            st.info("ℹ️ Tabela em modo de leitura. Para editar valores, insira a senha na barra lateral.")
            st.dataframe(df_exibir, width="stretch")
            
        st.download_button("📥 Baixar Tabela em Excel", data=df_para_excel(df_exibir), file_name="Cronograma_Distritos.xlsx")
    else:
        st.warning("Nenhum dado encontrado para o Cronograma.")

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
    else:
        st.warning("Nenhum dado de Subprefeituras encontrado.")

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
    else:
        st.warning("Nenhum registro encontrado na Base de Dados.")

# --- ABA 6: FOTOS, MAPAS E LEGENDAS ---
with aba_mapa:
    st.markdown("### Visualização de Mapas e Documentos")
    if os.path.exists("cronograma.docx"):
        st.success("📄 Arquivo original `cronograma.docx` encontrado na raiz do projeto.")
    else:
        st.info("Nenhum documento de mapa anexado.")

# --- ABA 7: NOTÍCIAS E PUBLICAÇÕES ---
with aba_noticias:
    st.markdown("### Publicações e PDF")
    df_not = ler_aba("noticias_pdf")
    if df_not.empty:
        df_not = pd.DataFrame([{"titulo": "Inauguração do Centro Dia do Idoso", "link": "https://www.prefeitura.sp.gov.br", "data": "2026-08-03"}])
    
    if st.session_state["modo_edicao"]:
        df_edit_n = st.data_editor(df_not, num_rows="dynamic", key="editor_not", width="stretch")
        if st.button("💾 Salvar Notícias"):
            if salvar_aba(df_edit_n, "noticias_pdf"):
                st.success("Notícias atualizadas!")
                st.rerun()
    else:
        st.dataframe(df_not, width="stretch")

# --- ABA 8: SOBRE ---
with aba_sobre:
    st.markdown("### Sobre o Sistema")
    st.markdown("""
    **Painel Geral de Gestão - Políticas e Atenção ao Idoso (SP)**
    
    Este sistema é uma ferramenta para monitoramento de equipamentos sociais, conselheiros municipais e registros de atenção à pessoa idosa na cidade de São Paulo.
    
    - **Sincronização:** Google Sheets API
    - **Modo de Acesso:** Leitura Aberta / Edição Protegida por Senha
    """)
