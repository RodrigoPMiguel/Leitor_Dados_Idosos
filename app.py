import streamlit as st
import pandas as pd
import sqlite3
import docx
import os
import re
import io
import fitz  # PyMuPDF para converter PDF em imagem

# Configuração da página
st.set_page_config(page_title="Painel de Gestão - Defesa do Idoso SP", layout="wide")

DB_FILE = "dados_idosos.db"

# --- SENHAS DE ACESSO AO MODO EDIÇÃO ---
SENHAS_VALIDAS = ["idoso2026", "conselho2026", "gestao2026"]

# Controle de estado do login
if "modo_edicao" not in st.session_state:
    st.session_state["modo_edicao"] = False

# --- FUNÇÃO AUXILIAR PARA GERAR RELATÓRIOS EM EXCEL (.XLSX) ---
def df_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
    return output.getvalue()

# Mapeamento de meses para ordenação numérica
MESES_MAPA = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
    "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
    "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
}

# --- BANCO DE DADOS SQLITE ---
def get_connection():
    return sqlite3.connect(DB_FILE)

def inicializar_banco():
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Tabela de Conselheiros
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS conselheiros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT, cargo TEXT, telefone TEXT, email TEXT,
            regiao TEXT, observacoes TEXT, foto BLOB
        )
    """)
    
    cursor.execute("SELECT COUNT(*) FROM conselheiros")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO conselheiros (nome, cargo, regiao) VALUES ('Francisco Miguel Filho', 'Conselheiro', 'São Paulo')")
        cursor.execute("INSERT INTO conselheiros (nome, cargo, regiao) VALUES ('Vanessa Nassif', 'Conselheira', 'São Paulo')")
    
    # 2. Tabela de Anotações Importantes
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS anotacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            titulo TEXT, categoria TEXT, conteudo TEXT,
            data_criacao DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # 3. Tabela do Cronograma Desmembrado
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cronograma_dados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            distrito TEXT, ano_criacao TEXT, no_mapa TEXT, pop_total TEXT,
            pop_masc TEXT, pop_fem TEXT, subpref_sn TEXT, cras TEXT, creas TEXT,
            caei TEXT, nci TEXT, ilpi TEXT, bpc TEXT, rank_vun TEXT,
            ubs TEXT, ama TEXT, idrpg TEXT, emad TEXT, ursi TEXT, upa TEXT,
            cdi TEXT, pai TEXT, ceu TEXT, cdc TEXT, ccint TEXT,
            ilpi_2setor TEXT, cdia TEXT, nci_2setor TEXT, outros_projetos TEXT
        )
    """)

    cursor.execute("SELECT COUNT(*) FROM cronograma_dados")
    if cursor.fetchone()[0] == 0 and os.path.exists("cronograma.docx"):
        try:
            doc = docx.Document("cronograma.docx")
            if len(doc.tables) > 0:
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
                    
                    cursor.execute("""
                        INSERT INTO cronograma_dados (
                            distrito, ano_criacao, no_mapa, pop_total, pop_masc, pop_fem,
                            subpref_sn, cras, creas, caei, nci, ilpi, bpc, rank_vun,
                            ubs, ama, idrpg, emad, ursi, upa, cdi, pai, ceu, cdc, ccint,
                            ilpi_2setor, cdia, nci_2setor, outros_projetos
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (
                        distrito, txts[1], txts[2], txts[3], txts[4], txts[5], txts[6],
                        cras, creas, txts[8].replace('-',''), txts[9].replace('-',''), txts[10].replace('-',''),
                        txts[11], txts[12], ubs, ama, txts[14], txts[15], ursi, upa, cdi, pai,
                        ceu, cdc, ccint, ilpi_2, cdia, nci_2, col19
                    ))
        except Exception:
            pass

    # 4. Tabela de Subprefeituras
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS subprefeituras_dados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subprefeitura TEXT, regiao TEXT, ilpi_particular INT, ilpi_conveniada INT,
            ilpi_nao_conveniada INT, cdi_particular INT, cdi_conveniado INT,
            nci_conveniado INT, nci_nao_conveniado INT, caei INT, ccinter INT, total INT
        )
    """)
    cursor.execute("SELECT COUNT(*) FROM subprefeituras_dados")
    if cursor.fetchone()[0] == 0 and os.path.exists("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx"):
        try:
            df_t = pd.read_excel("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx", sheet_name="TOTAIS").dropna(how="all")
            if 'Unnamed: 0' in df_t.columns: df_t = df_t.drop(columns=['Unnamed: 0'])
            for _, r in df_t.iterrows():
                cursor.execute("""
                    INSERT INTO subprefeituras_dados (
                        regiao, subprefeitura, ilpi_particular, ilpi_conveniada, ilpi_nao_conveniada,
                        cdi_particular, cdi_conveniado, nci_conveniado, nci_nao_conveniado, caei, ccinter, total
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    str(r.get('REGIÃO', '')), str(r.get('SUBPREFEITURA', '')),
                    int(r.get('ILPI particular', 0) or 0), int(r.get('ILPI conveniada', 0) or 0),
                    int(r.get('ILPI não conveniada', 0) or 0), int(r.get('CDI particular', 0) or 0),
                    int(r.get('CDI conveniado', 0) or 0), int(r.get('NCI conveniado', 0) or 0),
                    int(r.get('NCI não conveniado', 0) or 0), int(r.get('CAEI', 0) or 0),
                    int(r.get('CCINTER', 0) or 0), int(r.get('TOTAL POR SUBPREF', 0) or 0)
                ))
        except Exception:
            pass

    # 5. Tabela de Registros Base
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS registros_base (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            no_registro TEXT, vigencia TEXT, data TEXT, modalidade TEXT,
            nome_institucional TEXT, nome_fantasia TEXT, cnpj_matriz TEXT, cnpj_filial TEXT,
            programa TEXT, res TEXT, convenio TEXT, endereco TEXT, cep TEXT, bairro TEXT,
            subprefeitura TEXT, regiao TEXT, telefone TEXT, email TEXT, site TEXT,
            capacidade TEXT, atendidos TEXT, sei TEXT, id_original TEXT
        )
    """)
    cursor.execute("SELECT COUNT(*) FROM registros_base")
    if cursor.fetchone()[0] == 0 and os.path.exists("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx"):
        try:
            df_b = pd.read_excel("Tabela Geral de Registros 2024 RECUPERADO (1).xlsx", sheet_name="Base de Dados", header=1).dropna(how="all")
            for _, r in df_b.iterrows():
                cursor.execute("""
                    INSERT INTO registros_base (
                        no_registro, vigencia, data, modalidade, nome_institucional, nome_fantasia,
                        cnpj_matriz, cnpj_filial, programa, res, convenio, endereco, cep, bairro,
                        subprefeitura, regiao, telefone, email, site, capacidade, atendidos, sei, id_original
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    str(r.get('Nº REGISTRO', '')), str(r.get('VIGÊNCIA', '')), str(r.get('DATA', '')),
                    str(r.get('MODALIDADE', '')), str(r.get('NOME INSTITUCIONAL', '')), str(r.get('NOME FANTASIA', '')),
                    str(r.get('CNPJ Matriz', '')), str(r.get('CNPJ Filial', '')), str(r.get('PROGRAMA', '')),
                    str(r.get('RES', '')), str(r.get('CONVÊNIO', '')), str(r.get('ENDEREÇO DO PROGRAMA', '')),
                    str(r.get('CEP', '')), str(r.get('BAIRRO', '')), str(r.get('SUBPREFEITURA', '')),
                    str(r.get('REGIÃO', '')), str(r.get('TELEFONE', '')), str(r.get('E-MAIL', '')),
                    str(r.get('SITE', '')), str(r.get('CAPACIDADE', '')), str(r.get('ATENDIDOS', '')),
                    str(r.get('SEI', '')), str(r.get('ID', ''))
                ))
        except Exception:
            pass

    # 6. Tabela de Galeria de Imagens/Fotos
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS galeria_fotos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            titulo TEXT, foto BLOB
        )
    """)

    # 7. Tabela de Notícias e Informativos (PDF / Doc)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS noticias_pdf (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            titulo TEXT, mes_ref TEXT, ano_ref INT, mes_num INT,
            nome_arquivo TEXT, arquivo BLOB,
            data_upload DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.commit()
    conn.close()

inicializar_banco()

# --- BARRA LATERAL (MODO EDIÇÃO COM SENHA) ---
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
    st.image("logo.png", use_container_width=True)

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

# -------------------------------------------------------------
# ABA 1: CRONOGRAMA
# -------------------------------------------------------------
with aba_crono:
    st.subheader("Cronograma por Distrito")
    
    conn = get_connection()
    df_crono_db = pd.read_sql_query("SELECT * FROM cronograma_dados", conn)
    conn.close()

    if st.session_state["modo_edicao"]:
        with st.expander("🛠️ Gerenciar Colunas (Adicionar / Apagar)"):
            c_add, c_del = st.columns(2)
            with c_add:
                st.markdown("**Adicionar Nova Coluna:**")
                nova_col = st.text_input("Nome da nova coluna:", key="crono_add_col")
                if st.button("➕ Adicionar Coluna no Cronograma"):
                    if nova_col and nova_col not in df_crono_db.columns:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(f'ALTER TABLE cronograma_dados ADD COLUMN "{nova_col}" TEXT')
                        conn.commit()
                        conn.close()
                        st.success(f"Coluna '{nova_col}' adicionada com sucesso!")
                        st.rerun()
                    elif nova_col in df_crono_db.columns:
                        st.warning("Esta coluna já existe.")
            
            with c_del:
                st.markdown("**Apagar Coluna Existente:**")
                col_para_apagar = st.selectbox("Selecione a coluna para remover:", [c for c in df_crono_db.columns if c != 'id'], key="crono_del_col")
                if st.button("🗑️ Apagar Coluna Selecionada", key="btn_del_crono"):
                    if col_para_apagar:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(f'ALTER TABLE cronograma_dados DROP COLUMN "{col_para_apagar}"')
                        conn.commit()
                        conn.close()
                        st.success(f"Coluna '{col_para_apagar}' removida com sucesso!")
                        st.rerun()

    col_filtro, _ = st.columns([1, 2])
    with col_filtro:
        distritos_lista = ["Todos os Distritos"] + sorted(list(df_crono_db['distrito'].dropna().unique()))
        distrito_selecionado = st.selectbox("🔍 Selecione o Distrito para Filtrar:", distritos_lista)

    if not st.session_state["modo_edicao"]:
        st.info("ℹ️ Tabela em modo de leitura. Para editar valores, insira a senha na barra lateral.")

    df_exibicao = df_crono_db[df_crono_db['distrito'] == distrito_selecionado] if distrito_selecionado != "Todos os Distritos" else df_crono_db
    
    if st.session_state["modo_edicao"]:
        edited_crono = st.data_editor(df_exibicao, num_rows="dynamic", use_container_width=True, key="editor_crono")
        col_btn1, col_btn2 = st.columns([1, 2])
        with col_btn1:
            if st.button("💾 Salvar Alterações no Cronograma"):
                conn = get_connection()
                edited_crono.to_sql("cronograma_dados", conn, if_exists="replace", index=False)
                conn.close()
                st.success("Alterações salvas com sucesso no banco de dados!")
                st.rerun()
        with col_btn2:
            excel_crono = df_para_excel(edited_crono)
            st.download_button("🖨️ Baixar Relatório Filtrado (Excel)", excel_crono, f"Cronograma_{distrito_selecionado}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.dataframe(df_exibicao, use_container_width=True)
        excel_crono = df_para_excel(df_exibicao)
        st.download_button("🖨️ Baixar Relatório Filtrado (Excel)", excel_crono, f"Cronograma_{distrito_selecionado}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -------------------------------------------------------------
# ABA 2: SUBPREFEITURAS
# -------------------------------------------------------------
with aba_subpref:
    st.subheader("Totais e Indicadores das Subprefeituras")
    
    conn = get_connection()
    df_sub_db = pd.read_sql_query("SELECT * FROM subprefeituras_dados", conn)
    conn.close()

    if st.session_state["modo_edicao"]:
        with st.expander("🛠️ Gerenciar Colunas (Adicionar / Apagar)"):
            s_add, s_del = st.columns(2)
            with s_add:
                st.markdown("**Adicionar Nova Coluna:**")
                nova_col_sub = st.text_input("Nome da nova coluna:", key="sub_add_col")
                if st.button("➕ Adicionar Coluna em Subprefeituras"):
                    if nova_col_sub and nova_col_sub not in df_sub_db.columns:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(f'ALTER TABLE subprefeituras_dados ADD COLUMN "{nova_col_sub}" TEXT')
                        conn.commit()
                        conn.close()
                        st.success(f"Coluna '{nova_col_sub}' adicionada com sucesso!")
                        st.rerun()
                    elif nova_col_sub in df_sub_db.columns:
                        st.warning("Esta coluna já existe.")
            
            with s_del:
                st.markdown("**Apagar Coluna Existente:**")
                col_del_sub = st.selectbox("Selecione a coluna para remover:", [c for c in df_sub_db.columns if c != 'id'], key="sub_del_col")
                if st.button("🗑️ Apagar Coluna Selecionada", key="btn_del_sub"):
                    if col_del_sub:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(f'ALTER TABLE subprefeituras_dados DROP COLUMN "{col_del_sub}"')
                        conn.commit()
                        conn.close()
                        st.success(f"Coluna '{col_del_sub}' removida com sucesso!")
                        st.rerun()

    if not st.session_state["modo_edicao"]:
        st.info("ℹ️ Tabela em modo de leitura. Para editar valores, insira a senha na barra lateral.")
        st.dataframe(df_sub_db, use_container_width=True)
        excel_sub = df_para_excel(df_sub_db)
        st.download_button("🖨️ Baixar Relatório de Subprefeituras (Excel)", excel_sub, "Subprefeituras_Totais.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        edited_sub = st.data_editor(df_sub_db, num_rows="dynamic", use_container_width=True, key="editor_sub")
        col_s1, col_s2 = st.columns([1, 2])
        with col_s1:
            if st.button("💾 Salvar Alterações das Subprefeituras"):
                conn = get_connection()
                edited_sub.to_sql("subprefeituras_dados", conn, if_exists="replace", index=False)
                conn.close()
                st.success("Dados das Subprefeituras atualizados com sucesso!")
                st.rerun()
        with col_s2:
            excel_sub = df_para_excel(edited_sub)
            st.download_button("🖨️ Baixar Relatório de Subprefeituras (Excel)", excel_sub, "Subprefeituras_Totais.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -------------------------------------------------------------
# ABA 3: CONSELHEIROS MUNICIPAIS
# -------------------------------------------------------------
with aba_conselheiros:
    st.subheader("Conselheiros Municipais do Conselho do Idoso")
    
    if st.session_state["modo_edicao"]:
        col_cad, col_list = st.columns([1, 2])
        with col_cad:
            st.markdown("### Cadastrar Novo Conselheiro")
            with st.form("form_conselheiro", clear_on_submit=True):
                nome = st.text_input("Nome Completo:")
                cargo = st.text_input("Cargo / Função:", value="Conselheiro(a)")
                telefone = st.text_input("Telefone / WhatsApp:")
                email = st.text_input("E-mail de Contato:")
                regiao = st.text_input("Região / Subprefeitura:")
                obs = st.text_area("Observações:")
                foto = st.file_uploader("Foto do Conselheiro (Opcional):", type=["jpg", "png", "jpeg"])
                
                submitted = st.form_submit_button("➕ Cadastrar Conselheiro")
                if submitted and nome:
                    foto_bytes = foto.read() if foto else None
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO conselheiros (nome, cargo, telefone, email, regiao, observacoes, foto)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (nome, cargo, telefone, email, regiao, obs, foto_bytes))
                    conn.commit()
                    conn.close()
                    st.success(f"Conselheiro {nome} cadastrado com sucesso!")
                    st.rerun()
    else:
        st.info("ℹ️ Para cadastrar ou excluir conselheiros, ative o Modo Edição na barra lateral.")
        col_list = st.container()

    with col_list:
        st.markdown("### Lista de Conselheiros Cadastrados")
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, cargo, telefone, email, regiao, observacoes, foto FROM conselheiros")
        rows = cursor.fetchall()
        conn.close()
        
        for row in rows:
            c_id, c_nome, c_cargo, c_tel, c_email, c_regiao, c_obs, c_foto = row
            with st.expander(f"👤 {c_nome} - {c_cargo} ({c_regiao or 'SP'})"):
                col_img, col_info = st.columns([1, 3])
                with col_img:
                    if c_foto:
                        st.image(c_foto, width=120)
                    else:
                        st.markdown("🖼️ *Sem Foto*")
                with col_info:
                    st.write(f"**Telefone:** {c_tel or 'Não informado'}")
                    st.write(f"**E-mail:** {c_email or 'Não informado'}")
                    st.write(f"**Observações:** {c_obs or '-'}")
                    
                    if st.session_state["modo_edicao"]:
                        if st.button("❌ Excluir Conselheiro", key=f"del_cons_{c_id}"):
                            conn = get_connection()
                            c = conn.cursor()
                            c.execute("DELETE FROM conselheiros WHERE id = ?", (c_id,))
                            conn.commit()
                            conn.close()
                            st.rerun()

# -------------------------------------------------------------
# ABA 4: ANOTAÇÕES IMPORTANTES
# -------------------------------------------------------------
with aba_anotacoes:
    st.subheader("Bloco de Anotações e Lembretes Importantes")
    
    if st.session_state["modo_edicao"]:
        col_anot_form, col_anot_view = st.columns([1, 2])
        with col_anot_form:
            st.markdown("### Nova Anotação")
            with st.form("form_anotacao", clear_on_submit=True):
                titulo = st.text_input("Título da Anotação:")
                categoria = st.selectbox("Categoria:", ["Geral", "Reunião", "Vistoria ILPI", "Atendimento", "Outros"])
                conteudo = st.text_area("Conteúdo da Anotação / Lembrete:", height=150)
                
                salvar_anot = st.form_submit_button("📌 Salvar Anotação")
                if salvar_anot and titulo:
                    conn = get_connection()
                    c = conn.cursor()
                    c.execute("INSERT INTO anotacoes (titulo, categoria, conteudo) VALUES (?, ?, ?)", (titulo, categoria, conteudo))
                    conn.commit()
                    conn.close()
                    st.success("Anotação salva com sucesso!")
                    st.rerun()
    else:
        st.info("ℹ️ Para incluir novas anotações ou apagar existentes, ative o Modo Edição na barra lateral.")
        col_anot_view = st.container()

    with col_anot_view:
        st.markdown("### Anotações Salvas")
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, titulo, categoria, conteudo, data_criacao FROM anotacoes ORDER BY data_criacao DESC")
        anotacoes = cursor.fetchall()
        conn.close()
        
        for a_id, a_tit, a_cat, a_cont, a_data in anotacoes:
            with st.expander(f"📝 [{a_cat}] {a_tit} - {a_data[:16]}"):
                st.write(a_cont)
                if st.session_state["modo_edicao"]:
                    if st.button("🗑️ Excluir Anotação", key=f"del_anot_{a_id}"):
                        conn = get_connection()
                        c = conn.cursor()
                        c.execute("DELETE FROM anotacoes WHERE id = ?", (a_id,))
                        conn.commit()
                        conn.close()
                        st.rerun()

# -------------------------------------------------------------
# ABA 5: REGISTROS / CASAS DE REPOUSO
# -------------------------------------------------------------
with aba_registros:
    st.subheader("Base de Registros e Equipamentos")
    
    conn = get_connection()
    df_reg_db = pd.read_sql_query("SELECT * FROM registros_base", conn)
    conn.close()

    if st.session_state["modo_edicao"]:
        with st.expander("🛠️ Gerenciar Colunas (Adicionar / Apagar)"):
            r_add, r_del = st.columns(2)
            with r_add:
                st.markdown("**Adicionar Nova Coluna:**")
                nova_col_reg = st.text_input("Nome da nova coluna:", key="reg_add_col")
                if st.button("➕ Adicionar Coluna nos Registros"):
                    if nova_col_reg and nova_col_reg not in df_reg_db.columns:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(f'ALTER TABLE registros_base ADD COLUMN "{nova_col_reg}" TEXT')
                        conn.commit()
                        conn.close()
                        st.success(f"Coluna '{nova_col_reg}' adicionada com sucesso!")
                        st.rerun()
                    elif nova_col_reg in df_reg_db.columns:
                        st.warning("Esta coluna já existe.")
            
            with r_del:
                st.markdown("**Apagar Coluna Existente:**")
                col_del_reg = st.selectbox("Selecione a coluna para remover:", [c for c in df_reg_db.columns if c != 'id'], key="reg_del_col")
                if st.button("🗑️ Apagar Coluna Selecionada", key="btn_del_reg"):
                    if col_del_reg:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(f'ALTER TABLE registros_base DROP COLUMN "{col_del_reg}"')
                        conn.commit()
                        conn.close()
                        st.success(f"Coluna '{col_del_reg}' removida com sucesso!")
                        st.rerun()

    c1, c2 = st.columns(2)
    with c1:
        subpref_sel = st.selectbox("Subprefeitura:", ["Todas"] + sorted([str(x) for x in df_reg_db['subprefeitura'].dropna().unique() if str(x) != 'nan']))
    with c2:
        bairros_disp = df_reg_db['bairro'].dropna().unique() if subpref_sel == "Todas" else df_reg_db[df_reg_db['subprefeitura'] == subpref_sel]['bairro'].dropna().unique()
        bairro_sel = st.selectbox("Bairro:", ["Todos"] + sorted([str(x) for x in bairros_disp if str(x) != 'nan']))
        
    df_reg_filt = df_reg_db.copy()
    if subpref_sel != "Todas":
        df_reg_filt = df_reg_filt[df_reg_filt['subprefeitura'] == subpref_sel]
    if bairro_sel != "Todos":
        df_reg_filt = df_reg_filt[df_reg_filt['bairro'] == bairro_sel]
        
    if not st.session_state["modo_edicao"]:
        st.info("ℹ️ Tabela em modo de leitura. Para editar valores, insira a senha na barra lateral.")
        st.dataframe(df_reg_filt, use_container_width=True)
        excel_reg = df_para_excel(df_reg_filt)
        st.download_button("🖨️ Baixar Registros Filtrados (Excel)", excel_reg, f"Registros_{subpref_sel}_{bairro_sel}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        edited_reg = st.data_editor(df_reg_filt, num_rows="dynamic", use_container_width=True, key="editor_reg")
        col_r1, col_r2 = st.columns([1, 2])
        with col_r1:
            if st.button("💾 Salvar Alterações nos Registros"):
                conn = get_connection()
                edited_reg.to_sql("registros_base", conn, if_exists="replace", index=False)
                conn.close()
                st.success("Registros atualizados com sucesso!")
                st.rerun()
        with col_r2:
            excel_reg = df_para_excel(edited_reg)
            st.download_button("🖨️ Baixar Registros Filtrados (Excel)", excel_reg, f"Registros_{subpref_sel}_{bairro_sel}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -------------------------------------------------------------
# ABA 6: MAPA E LEGENDAS
# -------------------------------------------------------------
with aba_mapa:
    st.subheader("🗺️ Fotos, Mapas e Legendas Oficiais")
    
    @st.dialog("🔍 Visualização Ampliada da Imagem", width="large")
    def popup_imagem(caminho_ou_bytes, titulo_img="Imagem"):
        st.markdown(f"### {titulo_img}")
        zoom_px = st.slider("🔍 Ajuste o Tamanho / Zoom da Imagem (Pixels):", min_value=600, max_value=2200, value=1100, step=100, key="modal_zoom")
        st.image(caminho_ou_bytes, width=zoom_px)
        st.markdown("---")
        if st.button("❌ Fechar Imagem", use_container_width=True):
            st.rerun()

    col_mapa_main, col_upload = st.columns([1, 1])
    
    with col_mapa_main:
        st.markdown("### Mapa de São Paulo e Legenda Padrão")
        nome_imagem = "mapa_legendas.png"
        if os.path.exists(nome_imagem):
            st.image(nome_imagem, width=280, caption="Clique no botão abaixo para expandir em tela cheia")
            if st.button("🔎 Ampliar Mapa em Tela Cheia", key="btn_open_mapa"):
                popup_imagem(nome_imagem, "Mapa de SP e Legendas")
        else:
            st.warning(f"O arquivo '{nome_imagem}' não foi localizado na pasta 'Leitor_Dados_Idosos'.")

    with col_upload:
        if st.session_state["modo_edicao"]:
            st.markdown("### 📥 Enviar Nova Foto / Imagem")
            with st.form("form_upload_foto", clear_on_submit=True):
                titulo_foto = st.text_input("Título ou Identificação da Foto:")
                foto_up = st.file_uploader("Escolher arquivo de imagem do computador:", type=["jpg", "jpeg", "png", "bmp"])
                env = st.form_submit_button("📤 Enviar e Fixar Foto")
                
                if env and foto_up:
                    foto_bytes = foto_up.read()
                    tit = titulo_foto if titulo_foto else foto_up.name
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO galeria_fotos (titulo, foto) VALUES (?, ?)", (tit, foto_bytes))
                    conn.commit()
                    conn.close()
                    st.success("Foto enviada e fixada com sucesso na galeria!")
                    st.rerun()
        else:
            st.info("ℹ️ Para enviar novas fotos/mapas para a galeria, ative o Modo Edição na barra lateral.")

    st.markdown("---")
    st.subheader("🖼️ Galeria de Fotos e Documentos Fixados")
    
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, titulo, foto FROM galeria_fotos ORDER BY id DESC")
    fotos_db = cursor.fetchall()
    conn.close()

    if fotos_db:
        cols_gal = st.columns(4)
        for idx, (f_id, f_tit, f_bytes) in enumerate(fotos_db):
            with cols_gal[idx % 4]:
                st.markdown(f"**{f_tit}**")
                st.image(f_bytes, width=200)
                
                b1, b2 = st.columns(2)
                with b1:
                    if st.button("🔍 Ampliar", key=f"amp_img_{f_id}"):
                        popup_imagem(f_bytes, f_tit)
                with b2:
                    if st.session_state["modo_edicao"]:
                        if st.button("🗑️ Apagar", key=f"del_img_{f_id}"):
                            conn = get_connection()
                            c = conn.cursor()
                            c.execute("DELETE FROM galeria_fotos WHERE id = ?", (f_id,))
                            conn.commit()
                            conn.close()
                            st.rerun()
    else:
        st.info("Nenhuma foto extra foi adicionada ainda à galeria.")

# -------------------------------------------------------------
# ABA 7: NOTÍCIAS E PUBLICAÇÕES (PDF / DOC)
# -------------------------------------------------------------
with aba_noticias:
    st.subheader("📰 Notícias, Informativos e Publicações")

    @st.dialog("🔍 Leitor e Visualizador de Documento", width="large")
    def popup_pdf(bytes_arquivo, nome_arq, titulo_doc):
        st.markdown(f"### {titulo_doc}")
        
        # LÓGICA DE CONVERSÃO DE PDF PARA IMAGENS
        if nome_arq.lower().endswith('.pdf'):
            st.info("Abaixo estão as páginas do documento, convertidas em imagem para facilitar a sua leitura, livre de bloqueios do navegador.")
            try:
                # Tenta abrir o PDF e converter as páginas
                doc = fitz.open(stream=bytes_arquivo, filetype="pdf")
                for i in range(len(doc)):
                    page = doc.load_page(i)
                    pix = page.get_pixmap(dpi=150) # Qualidade da imagem gerada (150 DPI)
                    img_bytes = pix.tobytes("png")
                    
                    st.image(img_bytes, caption=f"Página {i + 1} de {len(doc)}", use_container_width=True)
                    st.markdown("---")
            except Exception as e:
                st.error("Não foi possível processar a imagem deste documento PDF.")
                
            # Mantém o botão de download sempre disponível no final
            st.download_button("📥 Baixar o PDF Original", bytes_arquivo, nome_arq, use_container_width=True)
            
        else:
            st.info("Visualização em tela cheia não suportada para arquivos do Word. Utilize o botão abaixo para baixá-lo em seu computador.")
            st.download_button("📥 Baixar Documento (Word)", bytes_arquivo, nome_arq, use_container_width=True)
            
        if st.button("❌ Fechar Visualizador", use_container_width=True):
            st.rerun()

    col_not_view, col_not_up = st.columns([2, 1])

    with col_not_up:
        if st.session_state["modo_edicao"]:
            st.markdown("### 📥 Publicar Nova Notícia / Documento")
            with st.form("form_noticia", clear_on_submit=True):
                tit_noticia = st.text_input("Título da Publicação / Informativo:")
                
                col_m, col_a = st.columns(2)
                with col_m:
                    mes_sel = st.selectbox("Mês de Referência:", list(MESES_MAPA.keys()))
                with col_a:
                    ano_sel = st.number_input("Ano de Referência:", min_value=2020, max_value=2035, value=2026, step=1)
                
                arq_up = st.file_uploader("Selecione o arquivo (PDF ou Word):", type=["pdf", "docx", "doc"])
                pub_btn = st.form_submit_button("📤 Publicar Notícia")

                if pub_btn and arq_up and tit_noticia:
                    b_arq = arq_up.read()
                    m_num = MESES_MAPA[mes_sel]
                    
                    conn = get_connection()
                    c = conn.cursor()
                    c.execute("""
                        INSERT INTO noticias_pdf (titulo, mes_ref, ano_ref, mes_num, nome_arquivo, arquivo)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (tit_noticia, mes_sel, ano_sel, m_num, arq_up.name, b_arq))
                    conn.commit()
                    conn.close()
                    
                    st.success("Publicação cadastrada com sucesso!")
                    st.rerun()
        else:
            st.info("ℹ️ Para cadastrar novas publicações ou notícias, ative o Modo Edição na barra lateral.")

    with col_not_view:
        st.markdown("### 📚 Publicações Recentes (Ordenadas por Mês/Ano)")
        
        conn = get_connection()
        c = conn.cursor()
        c.execute("""
            SELECT id, titulo, mes_ref, ano_ref, nome_arquivo, arquivo 
            FROM noticias_pdf 
            ORDER BY ano_ref DESC, mes_num DESC, id DESC
        """)
        noticias_db = c.fetchall()
        conn.close()

        if noticias_db:
            for n_id, n_tit, n_mes, n_ano, n_nome_arq, n_bytes in noticias_db:
                with st.expander(f"📄 [{n_mes} / {n_ano}] - {n_tit}"):
                    st.write(f"**Arquivo:** `{n_nome_arq}`")
                    
                    b_ler, b_down, b_del = st.columns([1, 1, 1])
                    
                    with b_ler:
                        if st.button("🔍 Abrir / Ler Janela", key=f"read_pdf_{n_id}"):
                            popup_pdf(n_bytes, n_nome_arq, n_tit)
                    
                    with b_down:
                        mime_t = "application/pdf" if n_nome_arq.lower().endswith('.pdf') else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        st.download_button("📥 Baixar File", n_bytes, file_name=n_nome_arq, mime=mime_t, key=f"down_pdf_{n_id}")
                    
                    with b_del:
                        if st.session_state["modo_edicao"]:
                            if st.button("🗑️ Apagar Publicação", key=f"del_pdf_{n_id}"):
                                conn = get_connection()
                                cur = conn.cursor()
                                cur.execute("DELETE FROM noticias_pdf WHERE id = ?", (n_id,))
                                conn.commit()
                                conn.close()
                                st.rerun()
        else:
            st.info("Nenhuma publicação ou notícia enviada ainda. Ative o Modo Edição para subir novos PDFs.")

# -------------------------------------------------------------
# ABA 8: SOBRE (ÚLTIMA ABA)
# -------------------------------------------------------------
with aba_sobre:
    st.subheader("ℹ️ Sobre esta Aplicação")
    st.markdown("""
    Este painel foi desenvolvido especialmente para apoio à consulta, gestão e acompanhamento de políticas públicas voltadas às pessoas idosas na cidade de São Paulo.
    
    * **👨‍💻 Criadores:** Rodrigo Prado Miguel & Francisco Miguel Filho
    * **⚙️ Desenvolvimento e Organização:** Aplicação local rápida e intuitiva.
    * **📊 Bases Utilizadas:** Tabela Geral de Registros 2024 e Cronograma de Indicadores por Distrito.
    * **🏷️ Versão:** 1.0.0
    """)
