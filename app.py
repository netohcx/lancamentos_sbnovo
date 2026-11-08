import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import date
import io

st.set_page_config(page_title="Gestão Suco Bagaço", layout="wide")

# ==========================================
# 1. CONEXÃO COM O BANCO (NUVEM OU LOCAL)
# ==========================================
def get_engine():
    try:
        if "DB_URL" in st.secrets:
            return create_engine(st.secrets["DB_URL"])
    except Exception:
        pass
    return create_engine("sqlite:///suco_bagaco.db")

engine = get_engine()

# ==========================================
# 2. CRIAÇÃO DAS TABELAS E OPÇÕES FIXAS
# ==========================================
def inicializar_banco():
    is_pg = "postgresql" in str(engine.url)
    id_col = "SERIAL PRIMARY KEY" if is_pg else "INTEGER PRIMARY KEY AUTOINCREMENT"
    
    with engine.begin() as conn:
        conn.execute(text(f"""
            CREATE TABLE IF NOT EXISTS compras (
                id {id_col},
                data_pedido TEXT,
                fornecedor TEXT,
                tipo_produto TEXT,
                valor REAL,
                forma_pagamento TEXT,
                data_vencimento TEXT,
                obs TEXT
            )
        """))

        conn.execute(text(f"""
            CREATE TABLE IF NOT EXISTS vendas (
                id {id_col},
                data TEXT,
                valor_sistema REAL,
                valor_caixa REAL,
                diferenca REAL,
                justificativa TEXT,
                caixa TEXT
            )
        """))

        conn.execute(text(f"""
            CREATE TABLE IF NOT EXISTS outras_despesas (
                id {id_col},
                data_pedido TEXT,
                fornecedor TEXT,
                tipo_produto TEXT,
                valor REAL,
                forma_pagamento TEXT,
                data_vencimento TEXT,
                obs TEXT
            )
        """))

        conn.execute(text(f"""
            CREATE TABLE IF NOT EXISTS opcoes_fixas (
                id {id_col},
                categoria TEXT,
                nome TEXT
            )
        """))

        qtd = conn.execute(text("SELECT COUNT(*) FROM opcoes_fixas")).scalar()
        if qtd == 0:
            iniciais = [
                ("fornecedor_compra", "Fornecedor Padrão"),
                ("produto_compra", "Polpas / Frutas"),
                ("produto_compra", "Hortifruti"),
                ("produto_compra", "Embalagens"),
                ("caixa_venda", "Caixa 1 - Manhã"),
                ("caixa_venda", "Caixa 2 - Noite"),
                ("fornecedor_despesa", "Novo Shopping"),
                ("produto_despesa", "Aluguel / Condomínio"),
                ("produto_despesa", "Energia / Água / Internet"),
                ("produto_despesa", "Manutenção")
            ]
            for cat, nome in iniciais:
                conn.execute(
                    text("INSERT INTO opcoes_fixas (categoria, nome) VALUES (:c, :n)"),
                    {"c": cat, "n": nome}
                )

inicializar_banco()

def listar_opcoes(categoria):
    df_fixas = pd.read_sql(
        text("SELECT DISTINCT nome FROM opcoes_fixas WHERE categoria = :c"),
        engine, params={"c": categoria}
    )
    lista = df_fixas["nome"].dropna().tolist()
    return sorted(lista) if lista else ["(Nenhum cadastrado)"]

def adicionar_opcao(categoria, novo_nome):
    if novo_nome is None:
        return False
    novo_nome = str(novo_nome).strip()
    if novo_nome:
        with engine.begin() as conn:
            existe = conn.execute(
                text("SELECT COUNT(*) FROM opcoes_fixas WHERE categoria = :c AND LOWER(nome) = LOWER(:n)"),
                {"c": categoria, "n": novo_nome}
            ).scalar()
            if existe == 0:
                conn.execute(
                    text("INSERT INTO opcoes_fixas (categoria, nome) VALUES (:c, :n)"),
                    {"c": categoria, "n": novo_nome}
                )
                return True
    return False

def excluir_opcao(categoria, nome_excluir):
    if nome_excluir and nome_excluir != "(Nenhum cadastrado)":
        with engine.begin() as conn:
            conn.execute(
                text("DELETE FROM opcoes_fixas WHERE categoria = :c AND nome = :n"),
                {"c": categoria, "n": nome_excluir}
            )
        return True
    return False

# Função para formatar datas sempre em DD/MM/AAAA
def formatar_data_br(val):
    if pd.isna(val) or not val:
        return ""
    try:
        dt = pd.to_datetime(val, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return str(val)

# ==========================================
# 3. SENHA ÚNICA BÁSICA
# ==========================================
if "logado" not in st.session_state:
    st.session_state.logado = False

if not st.session_state.logado:
    st.title("🍹 Lançamentos Suco Bagaço")
    col1, _ = st.columns(2)
    with col1:
        senha_digitada = st.text_input("Digite a senha de acesso:", type="password")
        if st.button("Entrar"):
            try:
                senha_correta = st.secrets["SENHA_ACESSO"]
            except Exception:
                senha_correta = "bagaco123"
                
            if senha_digitada == senha_correta:
                st.session_state.logado = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
    st.stop()

# ==========================================
# 4. BARRA LATERAL: MENU DE ABAS E EXCEL GERAL
# ==========================================
st.sidebar.title("🍹 Suco Bagaço")
menu = st.sidebar.radio("Selecione a Aba:", ["Compras", "Vendas", "Outras Despesas", "📊 Relatórios e Análises"])

st.sidebar.markdown("---")
st.sidebar.subheader("📥 Exportar Base Completa")

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    df_c_exp = pd.read_sql('SELECT data_pedido AS "DATA PEDIDO", fornecedor AS "FORNECEDOR", tipo_produto AS "TIPO PRODUTO", valor AS "VALOR", forma_pagamento AS "FORMA PAGAMENTO", data_vencimento AS "DATA VENCIMENTO", obs AS "OBS" FROM compras ORDER BY id DESC', engine)
    if not df_c_exp.empty:
        df_c_exp["DATA PEDIDO"] = pd.to_datetime(df_c_exp["DATA PEDIDO"], errors="coerce").dt.strftime("%d/%m/%Y")
        df_c_exp["DATA VENCIMENTO"] = pd.to_datetime(df_c_exp["DATA VENCIMENTO"], errors="coerce").dt.strftime("%d/%m/%Y")
    df_c_exp.to_excel(writer, sheet_name="Compras", index=False)

    df_v_exp = pd.read_sql('SELECT data AS "DATA", valor_sistema AS "VALOR SISTEMA", valor_caixa AS "VALOR CAIXA", diferenca AS "DIFERENÇA", justificativa AS "JUSTIFICATIVA", caixa AS "CAIXA" FROM vendas ORDER BY id DESC', engine)
    if not df_v_exp.empty:
        df_v_exp["DATA"] = pd.to_datetime(df_v_exp["DATA"], errors="coerce").dt.strftime("%d/%m/%Y")
    df_v_exp.to_excel(writer, sheet_name="Vendas", index=False)

    df_d_exp = pd.read_sql('SELECT data_pedido AS "DATA PEDIDO", fornecedor AS "FORNECEDOR", tipo_produto AS "TIPO PRODUTO", valor AS "VALOR", forma_pagamento AS "FORMA PAGAMENTO", data_vencimento AS "DATA VENCIMENTO", obs AS "OBS" FROM outras_despesas ORDER BY id DESC', engine)
    if not df_d_exp.empty:
        df_d_exp["DATA PEDIDO"] = pd.to_datetime(df_d_exp["DATA PEDIDO"], errors="coerce").dt.strftime("%d/%m/%Y")
        df_d_exp["DATA VENCIMENTO"] = pd.to_datetime(df_d_exp["DATA VENCIMENTO"], errors="coerce").dt.strftime("%d/%m/%Y")
    df_d_exp.to_excel(writer, sheet_name="Outras Despesas", index=False)

st.sidebar.download_button(
    label="Baixar Excel Completo (.xlsx)",
    data=buffer.getvalue(),
    file_name="Base_Completa_Suco_Bagaco.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.sidebar.markdown("---")
st.sidebar.subheader("📤 Importar Planilha")
arquivo_excel = st.sidebar.file_uploader("Subir arquivo Excel (.xlsx)", type=["xlsx"])

if arquivo_excel is not None:
    if st.sidebar.button("Confirmar Importação"):
        try:
            xls = pd.ExcelFile(arquivo_excel)
            if "Compras" in xls.sheet_names:
                df_c = pd.read_excel(arquivo_excel, sheet_name="Compras").dropna(how="all")
                if not df_c.empty:
                    df_c = df_c.rename(columns={"DATA PEDIDO": "data_pedido", "FORNECEDOR": "fornecedor", "TIPO PRODUTO": "tipo_produto", "VALOR": "valor", "FORMA PAGAMENTO": "forma_pagamento", "DATA VENCIMENTO": "data_vencimento", "OBS": "obs"})
                    df_c = df_c[[c for c in ["data_pedido", "fornecedor", "tipo_produto", "valor", "forma_pagamento", "data_vencimento", "obs"] if c in df_c.columns]]
                    df_c["data_pedido"] = pd.to_datetime(df_c["data_pedido"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
                    df_c["data_vencimento"] = pd.to_datetime(df_c["data_vencimento"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
                    df_c.to_sql("compras", engine, if_exists="append", index=False)
                    for f in df_c["fornecedor"].dropna().unique(): adicionar_opcao("fornecedor_compra", f)
                    for p in df_c["tipo_produto"].dropna().unique(): adicionar_opcao("produto_compra", p)

            if "Vendas" in xls.sheet_names:
                df_v = pd.read_excel(arquivo_excel, sheet_name="Vendas").dropna(how="all")
                if not df_v.empty:
                    df_v = df_v.rename(columns={"DATA": "data", "VALOR SISTEMA": "valor_sistema", "VALOR CAIXA": "valor_caixa", "DIFERENÇA": "diferenca", "JUSTIFICATIVA": "justificativa", "CAIXA": "caixa"})
                    df_v = df_v[[c for c in ["data", "valor_sistema", "valor_caixa", "diferenca", "justificativa", "caixa"] if c in df_v.columns]]
                    df_v["data"] = pd.to_datetime(df_v["data"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
                    df_v["diferenca"] = df_v["valor_caixa"] - df_v["valor_sistema"]
                    df_v.to_sql("vendas", engine, if_exists="append", index=False)
                    for cx in df_v["caixa"].dropna().unique(): adicionar_opcao("caixa_venda", cx)

            if "Outras Despesas" in xls.sheet_names:
                df_d = pd.read_excel(arquivo_excel, sheet_name="Outras Despesas").dropna(how="all")
                if not df_d.empty:
                    df_d = df_d.rename(columns={"DATA PEDIDO": "data_pedido", "FORNECEDOR": "fornecedor", "TIPO PRODUTO": "tipo_produto", "VALOR": "valor", "FORMA PAGAMENTO": "forma_pagamento", "DATA VENCIMENTO": "data_vencimento", "OBS": "obs"})
                    df_d = df_d[[c for c in ["data_pedido", "fornecedor", "tipo_produto", "valor", "forma_pagamento", "data_vencimento", "obs"] if c in df_d.columns]]
                    df_d["data_pedido"] = pd.to_datetime(df_d["data_pedido"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
                    df_d["data_vencimento"] = pd.to_datetime(df_d["data_vencimento"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
                    df_d.to_sql("outras_despesas", engine, if_exists="append", index=False)
                    for f in df_d["fornecedor"].dropna().unique(): adicionar_opcao("fornecedor_despesa", f)
                    for p in df_d["tipo_produto"].dropna().unique(): adicionar_opcao("produto_despesa", p)

            st.sidebar.success("Planilha importada com sucesso!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Erro ao importar: {e}")

st.sidebar.markdown("---")
if st.sidebar.button("Sair"):
    st.session_state.logado = False
    st.rerun()

# ==========================================
# ABA 1: COMPRAS
# ==========================================
if menu == "Compras":
    st.header("📦 Compras")
    
    with st.expander("⚙️ Gerenciar Lista de Fornecedores e Produtos (Adicionar / Excluir)"):
        col_add1, col_add2 = st.columns(2)
        with col_add1:
            st.markdown("**Fornecedores (Compras)**")
            novo_forn = st.text_input("Cadastrar novo Fornecedor:", key="nf_comp")
            if st.button("➕ Adicionar Fornecedor", key="btn_f_comp"):
                if adicionar_opcao("fornecedor_compra", novo_forn):
                    st.success(f"'{novo_forn}' adicionado!")
                    st.rerun()
            st.markdown("---")
            forn_del = st.selectbox("Remover Fornecedor:", listar_opcoes("fornecedor_compra"), key="del_f_comp_sel")
            if st.button("🗑️ Excluir Fornecedor", key="btn_del_f_comp"):
                if excluir_opcao("fornecedor_compra", forn_del):
                    st.warning(f"Removido!")
                    st.rerun()

        with col_add2:
            st.markdown("**Tipos de Produto (Compras)**")
            novo_prod = st.text_input("Cadastrar novo Tipo de Produto:", key="np_comp")
            if st.button("➕ Adicionar Tipo de Produto", key="btn_p_comp"):
                if adicionar_opcao("produto_compra", novo_prod):
                    st.success(f"'{novo_prod}' adicionado!")
                    st.rerun()
            st.markdown("---")
            prod_del = st.selectbox("Remover Produto:", listar_opcoes("produto_compra"), key="del_p_comp_sel")
            if st.button("🗑️ Excluir Produto", key="btn_del_p_comp"):
                if excluir_opcao("produto_compra", prod_del):
                    st.warning(f"Removido!")
                    st.rerun()

    with st.form("form_compras", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        data_pedido = c1.date_input("DATA PEDIDO", date.today(), format="DD/MM/YYYY")
        fornecedor = c2.selectbox("FORNECEDOR", listar_opcoes("fornecedor_compra"))
        tipo_produto = c3.selectbox("TIPO PRODUTO", listar_opcoes("produto_compra"))
        
        c4, c5, c6 = st.columns(3)
        valor = c4.number_input("VALOR (R$)", min_value=0.0, step=10.0, format="%.2f")
        forma_pagamento = c5.selectbox("FORMA PAGAMENTO", ["Boleto", "PIX", "Cartão", "Dinheiro", "Transferência"])
        data_vencimento = c6.date_input("DATA VENCIMENTO", date.today(), format="DD/MM/YYYY")
        
        obs = st.text_input("OBS")
        
        if st.form_submit_button("Salvar Compra"):
            with engine.begin() as conn:
                conn.execute(text("""
                    INSERT INTO compras (data_pedido, fornecedor, tipo_produto, valor, forma_pagamento, data_vencimento, obs)
                    VALUES (:dp, :f, :tp, :v, :fp, :dv, :obs)
                """), {
                    "dp": data_pedido.strftime("%Y-%m-%d"), "f": fornecedor, "tp": tipo_produto, 
                    "v": valor, "fp": forma_pagamento, "dv": data_vencimento.strftime("%Y-%m-%d"), "obs": obs
                })
            st.success("Compra registrada!")
            st.rerun()

    df_compras = pd.read_sql('SELECT id AS "ID", data_pedido AS "DATA PEDIDO", fornecedor AS "FORNECEDOR", tipo_produto AS "TIPO PRODUTO", valor AS "VALOR", forma_pagamento AS "FORMA PAGAMENTO", data_vencimento AS "DATA VENCIMENTO", obs AS "OBS" FROM compras ORDER BY id DESC', engine)
    if not df_compras.empty:
        df_compras["DATA PEDIDO"] = df_compras["DATA PEDIDO"].apply(formatar_data_br)
        df_compras["DATA VENCIMENTO"] = df_compras["DATA VENCIMENTO"].apply(formatar_data_br)
    st.dataframe(df_compras, use_container_width=True, hide_index=True)

    with st.expander("🗑️ Excluir um lançamento da tabela de Compras"):
        id_del = st.number_input("Digite o ID da linha para excluir:", min_value=1, step=1, key="del_comp")
        if st.button("Excluir Compra"):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM compras WHERE id = :id"), {"id": id_del})
            st.rerun()

# ==========================================
# ABA 2: VENDAS
# ==========================================
elif menu == "Vendas":
    st.header("💰 Vendas")
    
    with st.expander("⚙️ Gerenciar Lista de Caixas (Adicionar / Excluir)"):
        col_cx1, col_cx2 = st.columns(2)
        with col_cx1:
            novo_cx = st.text_input("Cadastrar novo Caixa:", key="nc_venda")
            if st.button("➕ Adicionar Caixa", key="btn_c_venda"):
                if adicionar_opcao("caixa_venda", novo_cx):
                    st.success(f"Adicionado!")
                    st.rerun()
        with col_cx2:
            cx_del = st.selectbox("Remover Caixa:", listar_opcoes("caixa_venda"), key="del_cx_sel")
            if st.button("🗑️ Excluir Caixa", key="btn_del_cx"):
                if excluir_opcao("caixa_venda", cx_del):
                    st.warning(f"Removido!")
                    st.rerun()

    with st.form("form_vendas", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        data_v = c1.date_input("DATA", date.today(), format="DD/MM/YYYY")
        valor_sistema = c2.number_input("VALOR SISTEMA (R$)", min_value=0.0, step=10.0, format="%.2f")
        valor_caixa = c3.number_input("VALOR CAIXA (R$)", min_value=0.0, step=10.0, format="%.2f")
        caixa_resp = c4.selectbox("CAIXA", listar_opcoes("caixa_venda"))
        
        justificativa = st.text_input("JUSTIFICATIVA")
        
        if st.form_submit_button("Salvar Venda"):
            diferenca = round(valor_caixa - valor_sistema, 2)
            with engine.begin() as conn:
                conn.execute(text("""
                    INSERT INTO vendas (data, valor_sistema, valor_caixa, diferenca, justificativa, caixa)
                    VALUES (:dt, :vs, :vc, :dif, :just, :cx)
                """), {
                    "dt": data_v.strftime("%Y-%m-%d"), "vs": valor_sistema, "vc": valor_caixa, 
                    "dif": diferenca, "just": justificativa, "cx": caixa_resp
                })
            st.success(f"Venda registrada! Diferença: R$ {diferenca:.2f}")
            st.rerun()

    df_vendas = pd.read_sql('SELECT id AS "ID", data AS "DATA", valor_sistema AS "VALOR SISTEMA", valor_caixa AS "VALOR CAIXA", diferenca AS "DIFERENÇA", justificativa AS "JUSTIFICATIVA", caixa AS "CAIXA" FROM vendas ORDER BY id DESC', engine)
    if not df_vendas.empty:
        df_vendas["DATA"] = df_vendas["DATA"].apply(formatar_data_br)
    st.dataframe(df_vendas, use_container_width=True, hide_index=True)

    with st.expander("🗑️ Excluir um lançamento da tabela de Vendas"):
        id_del_v = st.number_input("Digite o ID da linha para excluir:", min_value=1, step=1, key="del_vend")
        if st.button("Excluir Venda"):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM vendas WHERE id = :id"), {"id": id_del_v})
            st.rerun()

# ==========================================
# ABA 3: OUTRAS DESPESAS
# ==========================================
elif menu == "Outras Despesas":
    st.header("🧾 Outras Despesas")
    
    with st.expander("⚙️ Gerenciar Fornecedores e Tipos de Despesa (Adicionar / Excluir)"):
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.markdown("**Fornecedores (Despesas)**")
            novo_forn_d = st.text_input("Cadastrar novo Fornecedor:", key="nf_desp")
            if st.button("➕ Adicionar Fornecedor", key="btn_f_desp"):
                if adicionar_opcao("fornecedor_despesa", novo_forn_d):
                    st.success(f"Adicionado!")
                    st.rerun()
            st.markdown("---")
            forn_d_del = st.selectbox("Remover Fornecedor:", listar_opcoes("fornecedor_despesa"), key="del_fd_sel")
            if st.button("🗑️ Excluir Fornecedor", key="btn_del_fd"):
                if excluir_opcao("fornecedor_despesa", forn_d_del):
                    st.warning(f"Removido!")
                    st.rerun()

        with col_d2:
            st.markdown("**Tipos de Despesa**")
            novo_prod_d = st.text_input("Cadastrar novo Tipo de Despesa:", key="np_desp")
            if st.button("➕ Adicionar Tipo de Despesa", key="btn_p_desp"):
                if adicionar_opcao("produto_despesa", novo_prod_d):
                    st.success(f"Adicionado!")
                    st.rerun()
            st.markdown("---")
            prod_d_del = st.selectbox("Remover Tipo de Despesa:", listar_opcoes("produto_despesa"), key="del_pd_sel")
            if st.button("🗑️ Excluir Tipo de Despesa", key="btn_del_pd"):
                if excluir_opcao("produto_despesa", prod_d_del):
                    st.warning(f"Removido!")
                    st.rerun()

    with st.form("form_despesas", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        data_pedido = c1.date_input("DATA PEDIDO", date.today(), format="DD/MM/YYYY")
        fornecedor = c2.selectbox("FORNECEDOR", listar_opcoes("fornecedor_despesa"))
        tipo_produto = c3.selectbox("TIPO PRODUTO", listar_opcoes("produto_despesa"))
        
        c4, c5, c6 = st.columns(3)
        valor = c4.number_input("VALOR (R$)", min_value=0.0, step=10.0, format="%.2f")
        forma_pagamento = c5.selectbox("FORMA PAGAMENTO", ["Boleto", "PIX", "Cartão", "Dinheiro", "Débito Automático", "Transferência"])
        data_vencimento = c6.date_input("DATA VENCIMENTO", date.today(), format="DD/MM/YYYY")
        
        obs = st.text_input("OBS")
        
        if st.form_submit_button("Salvar Despesa"):
            with engine.begin() as conn:
                conn.execute(text("""
                    INSERT INTO outras_despesas (data_pedido, fornecedor, tipo_produto, valor, forma_pagamento, data_vencimento, obs)
                    VALUES (:dp, :f, :tp, :v, :fp, :dv, :obs)
                """), {
                    "dp": data_pedido.strftime("%Y-%m-%d"), "f": fornecedor, "tp": tipo_produto, 
                    "v": valor, "fp": forma_pagamento, "dv": data_vencimento.strftime("%Y-%m-%d"), "obs": obs
                })
            st.success("Despesa registrada!")
            st.rerun()

    df_desp = pd.read_sql('SELECT id AS "ID", data_pedido AS "DATA PEDIDO", fornecedor AS "FORNECEDOR", tipo_produto AS "TIPO PRODUTO", valor AS "VALOR", forma_pagamento AS "FORMA PAGAMENTO", data_vencimento AS "DATA VENCIMENTO", obs AS "OBS" FROM outras_despesas ORDER BY id DESC', engine)
    if not df_desp.empty:
        df_desp["DATA PEDIDO"] = df_desp["DATA PEDIDO"].apply(formatar_data_br)
        df_desp["DATA VENCIMENTO"] = df_desp["DATA_VENCIMENTO"].apply(formatar_data_br) if "DATA_VENCIMENTO" in df_desp.columns else df_desp["DATA VENCIMENTO"].apply(formatar_data_br)
    st.dataframe(df_desp, use_container_width=True, hide_index=True)

    with st.expander("🗑️ Excluir um lançamento da tabela de Outras Despesas"):
        id_del_d = st.number_input("Digite o ID da linha para excluir:", min_value=1, step=1, key="del_desp")
        if st.button("Excluir Despesa"):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM outras_despesas WHERE id = :id"), {"id": id_del_d})
            st.rerun()

# ==========================================
# ABA 4: RELATÓRIOS E ANÁLISES
# ==========================================
elif menu == "📊 Relatórios e Análises":
    st.header("📊 Painel de Relatórios e Gestão")
    
    st.subheader("🔍 Filtros de Período")
    fc1, fc2 = st.columns(2)
    data_inicio = fc1.date_input("Data Inicial", date(date.today().year, date.today().month, 1), format="DD/MM/YYYY")
    data_fim = fc2.date_input("Data Final", date.today(), format="DD/MM/YYYY")
    
    df_compras_db = pd.read_sql("SELECT * FROM compras", engine)
    df_vendas_db = pd.read_sql("SELECT * FROM vendas", engine)
    df_despesas_db = pd.read_sql("SELECT * FROM outras_despesas", engine)
    
    # Converter datas do banco para datetime para filtragem correta
    if not df_compras_db.empty: df_compras_db["dt_obj"] = pd.to_datetime(df_compras_db["data_pedido"], errors="coerce").dt.date
    if not df_vendas_db.empty: df_vendas_db["dt_obj"] = pd.to_datetime(df_vendas_db["data"], errors="coerce").dt.date
    if not df_despesas_db.empty: df_despesas_db["dt_obj"] = pd.to_datetime(df_despesas_db["data_pedido"], errors="coerce").dt.date
    
    c_filt = df_compras_db[(df_compras_db["dt_obj"] >= data_inicio) & (df_compras_db["dt_obj"] <= data_fim)] if not df_compras_db.empty else pd.DataFrame()
    v_filt = df_vendas_db[(df_vendas_db["dt_obj"] >= data_inicio) & (df_vendas_db["dt_obj"] <= data_fim)] if not df_vendas_db.empty else pd.DataFrame()
    d_filt = df_despesas_db[(df_despesas_db["dt_obj"] >= data_inicio) & (df_despesas_db["dt_obj"] <= data_fim)] if not df_despesas_db.empty else pd.DataFrame()
    
    total_compras = c_filt["valor"].sum() if not c_filt.empty else 0.0
    total_vendas = v_filt["valor_caixa"].sum() if not v_filt.empty else 0.0
    total_despesas = d_filt["valor"].sum() if not d_filt.empty else 0.0
    total_diferenca_caixa = v_filt["diferenca"].sum() if not v_filt.empty else 0.0
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("📦 Total Compras", f"R$ {total_compras:,.2f}")
    kpi2.metric("💰 Total Vendas (Caixa)", f"R$ {total_vendas:,.2f}")
    kpi3.metric("🧾 Outras Despesas", f"R$ {total_despesas:,.2f}")
    kpi4.metric("⚖️ Quebra / Diferença Caixa", f"R$ {total_diferenca_caixa:,.2f}")
    
    st.markdown("---")
    
    tipo_relatorio = st.selectbox("Selecione o Relatório Desejado:", [
        "1. Vendas por Período e por Caixa",
        "2. Vendas e Divergências de Caixa",
        "3. Compras por Fornecedor e por Produto",
        "4. Outras Despesas (Fornecedor / Tipo)"
    ])
    
    buffer_rel = io.BytesIO()
    
    if tipo_relatorio.startswith("1"):
        st.subheader("📊 Relatório de Vendas por Período e Caixa")
        if not v_filt.empty:
            caixas_disp = ["Todos"] + list(v_filt["caixa"].dropna().unique())
            cx_escolhido = st.selectbox("Filtrar por Caixa específico:", caixas_disp)
            df_rep1 = v_filt if cx_escolhido == "Todos" else v_filt[v_filt["caixa"] == cx_escolhido]
            
            # Formatar exibição para DD/MM/AAAA
            df_view = df_rep1.copy()
            if "data" in df_view.columns:
                df_view["data"] = df_view["data"].apply(formatar_data_br)
            st.dataframe(df_view.drop(columns=["dt_obj"], errors="ignore"), use_container_width=True, hide_index=True)
            
            st.markdown("**Resumo Consolidado por Caixa:**")
            resumo_cx = df_rep1.groupby("caixa")[["valor_sistema", "valor_caixa", "diferenca"]].sum().reset_index()
            st.dataframe(resumo_cx, use_container_width=True, hide_index=True)
            
            with pd.ExcelWriter(buffer_rel, engine="openpyxl") as w:
                df_ex1 = df_rep1.drop(columns=["dt_obj"], errors="ignore").copy()
                if "data" in df_ex1.columns: df_ex1["data"] = pd.to_datetime(df_ex1["data"]).dt.strftime("%d/%m/%Y")
                df_ex1.to_excel(w, sheet_name="Vendas_Periodo", index=False)
                resumo_cx.to_excel(w, sheet_name="Resumo_Por_Caixa", index=False)
        else:
            st.info("Nenhuma venda registrada no período selecionado.")

    elif tipo_relatorio.startswith("2"):
        st.subheader("📊 Relatório de Vendas e Divergências de Caixa")
        if not v_filt.empty:
            df_view = v_filt.copy()
            if "data" in df_view.columns:
                df_view["data"] = df_view["data"].apply(formatar_data_br)
            st.dataframe(df_view.drop(columns=["dt_obj"], errors="ignore"), use_container_width=True, hide_index=True)
            
            with pd.ExcelWriter(buffer_rel, engine="openpyxl") as w:
                df_ex2 = v_filt.drop(columns=["dt_obj"], errors="ignore").copy()
                if "data" in df_ex2.columns: df_ex2["data"] = pd.to_datetime(df_ex2["data"]).dt.strftime("%d/%m/%Y")
                df_ex2.to_excel(w, sheet_name="Relatorio_Vendas", index=False)
        else:
            st.info("Nenhum dado encontrado no período.")

    elif tipo_relatorio.startswith("3"):
        st.subheader("📊 Relatório de Compras (Por Fornecedor e Produto)")
        if not c_filt.empty:
            forn_disp = ["Todos"] + list(c_filt["fornecedor"].dropna().unique())
            f_escolhido = st.selectbox("Filtrar por Fornecedor:", forn_disp)
            df_rep3 = c_filt if f_escolhido == "Todos" else c_filt[c_filt["fornecedor"] == f_escolhido]
            
            df_view = df_rep3.copy()
            if "data_pedido" in df_view.columns: df_view["data_pedido"] = df_view["data_pedido"].apply(formatar_data_br)
            if "data_vencimento" in df_view.columns: df_view["data_vencimento"] = df_view["data_vencimento"].apply(formatar_data_br)
            st.dataframe(df_view.drop(columns=["dt_obj"], errors="ignore"), use_container_width=True, hide_index=True)
            
            st.markdown("**Total Comprado por Fornecedor:**")
            res_forn = c_filt.groupby("fornecedor")["valor"].sum().reset_index()
            st.dataframe(res_forn, use_container_width=True, hide_index=True)
            
            with pd.ExcelWriter(buffer_rel, engine="openpyxl") as w:
                df_ex3 = df_rep3.drop(columns=["dt_obj"], errors="ignore").copy()
                if "data_pedido" in df_ex3.columns: df_ex3["data_pedido"] = pd.to_datetime(df_ex3["data_pedido"]).dt.strftime("%d/%m/%Y")
                if "data_vencimento" in df_ex3.columns: df_ex3["data_vencimento"] = pd.to_datetime(df_ex3["data_vencimento"]).dt.strftime("%d/%m/%Y")
                df_ex3.to_excel(w, sheet_name="Compras_Filtradas", index=False)
                res_forn.to_excel(w, sheet_name="Total_Por_Fornecedor", index=False)
        else:
            st.info("Nenhuma compra registrada no período.")

    elif tipo_relatorio.startswith("4"):
        st.subheader("📊 Relatório de Outras Despesas")
        if not d_filt.empty:
            df_view = d_filt.copy()
            if "data_pedido" in df_view.columns: df_view["data_pedido"] = df_view["data_pedido"].apply(formatar_data_br)
            if "data_vencimento" in df_view.columns: df_view["data_vencimento"] = df_view["data_vencimento"].apply(formatar_data_br)
            st.dataframe(df_view.drop(columns=["dt_obj"], errors="ignore"), use_container_width=True, hide_index=True)
            
            st.markdown("**Total por Tipo de Despesa:**")
            res_desp = d_filt.groupby("tipo_produto")["valor"].sum().reset_index()
            st.dataframe(res_desp, use_container_width=True, hide_index=True)
            
            with pd.ExcelWriter(buffer_rel, engine="openpyxl") as w:
                df_ex4 = d_filt.drop(columns=["dt_obj"], errors="ignore").copy()
                if "data_ex4" in df_ex4.columns: pass
                if "data_pedido" in df_ex4.columns: df_ex4["data_pedido"] = pd.to_datetime(df_ex4["data_pedido"]).dt.strftime("%d/%m/%Y")
                if "data_vencimento" in df_ex4.columns: df_ex4["data_vencimento"] = pd.to_datetime(df_ex4["data_vencimento"]).dt.strftime("%d/%m/%Y")
                df_ex4.to_excel(w, sheet_name="Outras_Despesas", index=False)
                res_desp.to_excel(w, sheet_name="Resumo_Por_Tipo", index=False)
        else:
            st.info("Nenhuma despesa registrada no período.")

    st.markdown("---")
    st.download_button(
        label="📥 Baixar Relatório Filtrado em Excel (.xlsx)",
        data=buffer_rel.getvalue(),
        file_name=f"Relatorio_Suco_Bagaco_{data_inicio.strftime('%d-%m-%Y')}_a_{data_fim.strftime('%d-%m-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )