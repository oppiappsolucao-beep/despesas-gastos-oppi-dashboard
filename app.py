import re
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="Despesas & Gastos OPPI",
    page_icon="💸",
    layout="wide"
)

TZ = ZoneInfo("America/Sao_Paulo")

# >>> ID da sua planilha (pela imagem enviada)
SHEET_ID = "1cQU5tNwSoiepTPHx_Qc7ZF1PcaER2gstW_dZQ0eCrB4"
WORKSHEET_NAME = "Página1"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# =========================================================
# ESTILO
# =========================================================
st.markdown("""
<style>
    .stApp {
        background: #f6f7fb;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1450px;
    }

    .main-title {
        text-align: center;
        font-size: 2.6rem;
        font-weight: 800;
        color: #1d2340;
        margin-bottom: 0.2rem;
    }

    .main-subtitle {
        text-align: center;
        font-size: 1.1rem;
        color: #6b7280;
        margin-bottom: 1.6rem;
    }

    .top-divider, .section-divider {
        width: 100%;
        height: 18px;
        background: #ffffff;
        border: 1px solid #ececf3;
        border-radius: 999px;
        margin: 0.8rem 0 1.4rem 0;
    }

    .filter-label {
        font-size: 0.94rem;
        color: #2f3552;
        font-weight: 600;
        margin-bottom: 0.3rem;
    }

    .kpi-card {
        background: #ffffff;
        border: 1px solid #ececf3;
        border-left: 6px solid #e91e63;
        border-radius: 22px;
        padding: 1.1rem 1.2rem 0.9rem 1.2rem;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        min-height: 158px;
    }

    .kpi-card.roxo {
        border-left-color: #7c3aed;
    }

    .kpi-card.rosa {
        border-left-color: #e91e63;
    }

    .kpi-card.verde {
        border-left-color: #10b981;
    }

    .kpi-card.laranja {
        border-left-color: #f59e0b;
    }

    .kpi-card.azul {
        border-left-color: #3b82f6;
    }

    .kpi-title {
        font-size: 1rem;
        font-weight: 700;
        color: #28314f;
        margin-bottom: 0.8rem;
    }

    .kpi-value {
        font-size: 2.05rem;
        font-weight: 800;
        color: #081b4b;
        line-height: 1.05;
        margin-bottom: 0.75rem;
    }

    .kpi-caption {
        font-size: 0.92rem;
        color: #667085;
    }

    .section-title {
        font-size: 1.45rem;
        font-weight: 800;
        color: #14213d;
        margin-bottom: 0.25rem;
    }

    .section-text {
        color: #677185;
        font-size: 0.96rem;
        margin-bottom: 1rem;
    }

    .update-card {
        background: #ffffff;
        border: 1px solid #ececf3;
        border-radius: 24px;
        padding: 1.2rem 1.2rem 1rem 1.2rem;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        margin-bottom: 1rem;
    }

    .item-title {
        font-size: 1.55rem;
        font-weight: 800;
        color: #0b1d4d;
        margin-bottom: 0.4rem;
    }

    .item-meta {
        color: #64748b;
        font-size: 0.96rem;
        line-height: 1.6;
    }

    .item-meta b {
        color: #344054;
    }

    .item-value-label {
        color: #64748b;
        font-size: 0.95rem;
        font-weight: 600;
    }

    .item-value {
        font-size: 1.2rem;
        font-weight: 800;
        color: #081b4b;
    }

    .status-pill {
        display: inline-block;
        padding: 0.35rem 0.85rem;
        border-radius: 999px;
        font-size: 0.86rem;
        font-weight: 700;
    }

    .status-pago {
        background: #dff7e8;
        color: #118a43;
    }

    .status-apagar {
        background: #fde6e6;
        color: #c62828;
    }

    .status-areceber {
        background: #efe3ff;
        color: #6d28d9;
    }

    .small-note {
        font-size: 0.87rem;
        color: #6b7280;
        margin-top: 0.4rem;
    }

    div[data-testid="stDataFrame"] {
        background: #ffffff;
        border-radius: 18px;
        border: 1px solid #ececf3;
        padding: 0.35rem;
    }

    .stButton > button {
        border-radius: 14px;
        border: 1px solid #d6d9e5;
        font-weight: 600;
        min-height: 44px;
        background: white;
    }

    .stButton > button:hover {
        border-color: #7c3aed;
        color: #7c3aed;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# HELPERS
# =========================================================
def limpar_texto(s):
    return str(s or "").strip()

def normalizar_coluna(col):
    col = str(col or "")
    col = col.replace("\ufeff", "").replace("\xa0", " ").strip()
    return col

def parse_brl(valor):
    if pd.isna(valor):
        return 0.0
    s = str(valor).strip()
    if not s:
        return 0.0

    s = s.replace("R$", "").replace("r$", "").strip()
    s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)

    try:
        return float(s)
    except Exception:
        return 0.0

def formatar_brl(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def parse_data_br(valor):
    if pd.isna(valor):
        return pd.NaT

    if isinstance(valor, datetime):
        return pd.Timestamp(valor)

    s = str(valor).strip()
    if not s:
        return pd.NaT

    # tenta parse comum BR
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except Exception:
            pass

    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT

def extrair_mes_label(data):
    if pd.isna(data):
        return "Sem data"
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    return f"{meses[data.month - 1]}/{data.year}"

def status_class(status):
    s = limpar_texto(status).lower()
    if s == "pago":
        return "status-pill status-pago"
    if s == "a pagar":
        return "status-pill status-apagar"
    if s == "a receber":
        return "status-pill status-areceber"
    return "status-pill"

# =========================================================
# GOOGLE SHEETS
# =========================================================
@st.cache_resource(show_spinner=False)
def conectar_gsheet():
    creds = Credentials.from_service_account_info(
        st.secrets["google"],
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SHEET_ID)
    worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    return worksheet

@st.cache_data(ttl=30, show_spinner=False)
def carregar_dados():
    ws = conectar_gsheet()
    values = ws.get_all_values()

    if not values:
        return pd.DataFrame(), ws

    headers = [normalizar_coluna(h) for h in values[0]]
    rows = values[1:]

    # garante tamanho igual ao header
    rows_pad = []
    for row in rows:
        if len(row) < len(headers):
            row = row + [""] * (len(headers) - len(row))
        elif len(row) > len(headers):
            row = row[:len(headers)]
        rows_pad.append(row)

    df = pd.DataFrame(rows_pad, columns=headers)

    # normalização básica
    for c in df.columns:
        df[c] = df[c].astype(str).apply(lambda x: x.strip())

    # nomes esperados
    col_mes = "Mês" if "Mês" in df.columns else ("Mes" if "Mes" in df.columns else None)
    col_estabelecimento = "Estabelecimento" if "Estabelecimento" in df.columns else None
    col_valor = "Valor" if "Valor" in df.columns else None
    col_entrada = "Entrada" if "Entrada" in df.columns else None
    col_categoria = "Categoria" if "Categoria" in df.columns else None
    col_status = "Status" if "Status" in df.columns else None
    col_detalhes = "Detalhes" if "Detalhes" in df.columns else None
    col_whatsapp = "Whatsapp" if "Whatsapp" in df.columns else None

    # tipos auxiliares
    if col_mes:
        df["_data_mes"] = df[col_mes].apply(parse_data_br)
        df["_mes_label"] = df["_data_mes"].apply(extrair_mes_label)
    else:
        df["_data_mes"] = pd.NaT
        df["_mes_label"] = "Sem data"

    if col_valor:
        df["_valor_num"] = df[col_valor].apply(parse_brl)
    else:
        df["_valor_num"] = 0.0

    if col_status:
        df["_status_norm"] = df[col_status].astype(str).str.strip()
    else:
        df["_status_norm"] = ""

    if col_entrada:
        df["_entrada_norm"] = df[col_entrada].astype(str).str.strip()
    else:
        df["_entrada_norm"] = ""

    if col_estabelecimento:
        df["_estabelecimento_norm"] = df[col_estabelecimento].astype(str).str.strip()
    else:
        df["_estabelecimento_norm"] = ""

    if col_categoria:
        df["_categoria_norm"] = df[col_categoria].astype(str).str.strip()
    else:
        df["_categoria_norm"] = ""

    if col_detalhes:
        df["_detalhes_norm"] = df[col_detalhes].astype(str).str.strip()
    else:
        df["_detalhes_norm"] = ""

    # linha real da planilha (header = linha 1)
    df["_sheet_row"] = range(2, len(df) + 2)

    return df, ws

def atualizar_status_na_planilha(sheet_row, novo_status):
    ws = conectar_gsheet()
    header = ws.row_values(1)
    header = [normalizar_coluna(h) for h in header]

    if "Status" in header:
        status_col_idx = header.index("Status") + 1
    else:
        raise ValueError("Coluna 'Status' não encontrada na planilha.")

    ws.update_cell(sheet_row, status_col_idx, novo_status)
    st.cache_data.clear()

# =========================================================
# HEADER
# =========================================================
st.markdown('<div class="main-title">Despesas & Gastos OPPI</div>', unsafe_allow_html=True)
st.markdown('<div class="main-subtitle">Gestão financeira de receitas, despesas e status de pagamento</div>', unsafe_allow_html=True)
st.markdown('<div class="top-divider"></div>', unsafe_allow_html=True)

# =========================================================
# LOAD
# =========================================================
try:
    df, ws = carregar_dados()
except Exception as e:
    st.error("Erro ao conectar com a planilha do Google Sheets.")
    st.exception(e)
    st.stop()

if df.empty:
    st.warning("A planilha está vazia.")
    st.stop()

# =========================================================
# FILTROS
# =========================================================
meses_opcoes = ["Todos"] + sorted(
    [m for m in df["_mes_label"].dropna().unique().tolist() if m != "Sem data"],
    key=lambda x: (
        int(x.split("/")[1]) if "/" in x else 0,
        [
            "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
        ].index(x.split("/")[0]) if "/" in x and x.split("/")[0] in [
            "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
        ] else 0
    )
)

estab_opcoes = ["Todos"] + sorted([x for x in df["_estabelecimento_norm"].unique().tolist() if x])
categoria_opcoes = ["Todas"] + sorted([x for x in df["_categoria_norm"].unique().tolist() if x])
entrada_opcoes = ["Todas"] + sorted([x for x in df["_entrada_norm"].unique().tolist() if x])

f1, f2, f3, f4 = st.columns(4)

with f1:
    st.markdown('<div class="filter-label">Mês</div>', unsafe_allow_html=True)
    filtro_mes = st.selectbox("Mês", meses_opcoes, label_visibility="collapsed")

with f2:
    st.markdown('<div class="filter-label">Estabelecimento</div>', unsafe_allow_html=True)
    filtro_estab = st.selectbox("Estabelecimento", estab_opcoes, label_visibility="collapsed")

with f3:
    st.markdown('<div class="filter-label">Categoria</div>', unsafe_allow_html=True)
    filtro_categoria = st.selectbox("Categoria", categoria_opcoes, label_visibility="collapsed")

with f4:
    st.markdown('<div class="filter-label">Entrada</div>', unsafe_allow_html=True)
    filtro_entrada = st.selectbox("Entrada", entrada_opcoes, label_visibility="collapsed")

df_filtrado = df.copy()

if filtro_mes != "Todos":
    df_filtrado = df_filtrado[df_filtrado["_mes_label"] == filtro_mes]

if filtro_estab != "Todos":
    df_filtrado = df_filtrado[df_filtrado["_estabelecimento_norm"] == filtro_estab]

if filtro_categoria != "Todas":
    df_filtrado = df_filtrado[df_filtrado["_categoria_norm"] == filtro_categoria]

if filtro_entrada != "Todas":
    df_filtrado = df_filtrado[df_filtrado["_entrada_norm"] == filtro_entrada]

# =========================================================
# KPIs
# =========================================================
total_registros = len(df_filtrado)

total_receitas = df_filtrado.loc[
    df_filtrado["_entrada_norm"].str.lower() == "receita", "_valor_num"
].sum()

total_despesas = df_filtrado.loc[
    df_filtrado["_entrada_norm"].str.lower() == "despesa", "_valor_num"
].sum()

saldo = total_receitas - total_despesas

total_pago = df_filtrado.loc[
    df_filtrado["_status_norm"].str.lower() == "pago", "_valor_num"
].sum()

total_apagar = df_filtrado.loc[
    df_filtrado["_status_norm"].str.lower() == "a pagar", "_valor_num"
].sum()

total_areceber = df_filtrado.loc[
    df_filtrado["_status_norm"].str.lower() == "a receber", "_valor_num"
].sum()

c1, c2, c3, c4, c5, c6 = st.columns(6)

cards = [
    (c1, "Registros", str(total_registros), "total de lançamentos filtrados", "roxo"),
    (c2, "Receitas", formatar_brl(total_receitas), "soma de entradas do tipo Receita", "verde"),
    (c3, "Despesas", formatar_brl(total_despesas), "soma de entradas do tipo Despesa", "rosa"),
    (c4, "Saldo", formatar_brl(saldo), "receitas menos despesas", "azul"),
    (c5, "A pagar", formatar_brl(total_apagar), "somatório do status A Pagar", "laranja"),
    (c6, "A receber", formatar_brl(total_areceber), "somatório do status A Receber", "roxo"),
]

for col, titulo, valor, legenda, cor in cards:
    with col:
        st.markdown(
            f"""
            <div class="kpi-card {cor}">
                <div class="kpi-title">{titulo}</div>
                <div class="kpi-value">{valor}</div>
                <div class="kpi-caption">{legenda}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

st.markdown("<br>", unsafe_allow_html=True)

c7, c8 = st.columns(2)

with c7:
    st.markdown(
        f"""
        <div class="kpi-card verde">
            <div class="kpi-title">Pago</div>
            <div class="kpi-value">{formatar_brl(total_pago)}</div>
            <div class="kpi-caption">somatório do status Pago</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with c8:
    qtd_pago = (df_filtrado["_status_norm"].str.lower() == "pago").sum()
    qtd_apagar = (df_filtrado["_status_norm"].str.lower() == "a pagar").sum()
    qtd_areceber = (df_filtrado["_status_norm"].str.lower() == "a receber").sum()

    st.markdown(
        f"""
        <div class="kpi-card rosa">
            <div class="kpi-title">Resumo de status</div>
            <div class="kpi-value">{qtd_pago} / {qtd_apagar} / {qtd_areceber}</div>
            <div class="kpi-caption">Pago / A Pagar / A Receber</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================================================
# GRÁFICOS
# =========================================================
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

g1, g2 = st.columns(2)

with g1:
    st.markdown('<div class="section-title">📊 Valor por categoria</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-text">Soma dos valores agrupados por categoria.</div>', unsafe_allow_html=True)

    base_categoria = (
        df_filtrado.groupby("_categoria_norm", dropna=False)["_valor_num"]
        .sum()
        .reset_index()
        .rename(columns={"_categoria_norm": "Categoria", "_valor_num": "Valor"})
    )
    base_categoria = base_categoria[base_categoria["Categoria"].astype(str).str.strip() != ""]

    if not base_categoria.empty:
        fig_cat = px.bar(
            base_categoria.sort_values("Valor", ascending=False),
            x="Categoria",
            y="Valor",
            text="Valor"
        )
        fig_cat.update_traces(
            texttemplate="R$ %{y:,.2f}",
            textposition="outside"
        )
        fig_cat.update_layout(
            height=420,
            showlegend=False,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(l=20, r=20, t=20, b=20),
            xaxis_title="",
            yaxis_title="Valor"
        )
        fig_cat.update_yaxes(tickprefix="R$ ")
        st.plotly_chart(fig_cat, use_container_width=True)
    else:
        st.info("Sem dados para o gráfico de categoria.")

with g2:
    st.markdown('<div class="section-title">💰 Valor por status</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-text">Distribuição financeira por status atual.</div>', unsafe_allow_html=True)

    base_status = (
        df_filtrado.groupby("_status_norm", dropna=False)["_valor_num"]
        .sum()
        .reset_index()
        .rename(columns={"_status_norm": "Status", "_valor_num": "Valor"})
    )
    base_status = base_status[base_status["Status"].astype(str).str.strip() != ""]

    if not base_status.empty:
        fig_status = px.pie(
            base_status,
            names="Status",
            values="Valor",
            hole=0.62
        )
        fig_status.update_traces(
            textinfo="label+value",
            texttemplate="%{label}<br>R$ %{value:,.2f}"
        )
        fig_status.update_layout(
            height=420,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(l=20, r=20, t=20, b=20),
            legend_title=""
        )
        st.plotly_chart(fig_status, use_container_width=True)
    else:
        st.info("Sem dados para o gráfico de status.")

# =========================================================
# TABELA
# =========================================================
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">📋 Lançamentos</div>', unsafe_allow_html=True)
st.markdown('<div class="section-text">Visualização dos dados filtrados da planilha.</div>', unsafe_allow_html=True)

colunas_tabela = [c for c in ["Mês", "Mes", "Estabelecimento", "Valor", "Entrada", "Categoria", "Status", "Detalhes", "Whatsapp"] if c in df_filtrado.columns]
df_tabela = df_filtrado[colunas_tabela].copy()

st.dataframe(df_tabela, use_container_width=True, hide_index=True)

# =========================================================
# ATUALIZAR STATUS
# =========================================================
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">✏️ Atualizar status</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-text">Altere o campo <b>Status</b> diretamente pelo dashboard. Use a busca para localizar por estabelecimento, categoria, detalhes ou whatsapp.</div>',
    unsafe_allow_html=True
)

busca = st.text_input(
    "Buscar lançamento",
    placeholder="Ex.: Valeria, OpenAI, internet, mídia, whatsapp...",
    label_visibility="collapsed"
)

df_update = df.copy()

if filtro_mes != "Todos":
    df_update = df_update[df_update["_mes_label"] == filtro_mes]
if filtro_estab != "Todos":
    df_update = df_update[df_update["_estabelecimento_norm"] == filtro_estab]
if filtro_categoria != "Todas":
    df_update = df_update[df_update["_categoria_norm"] == filtro_categoria]
if filtro_entrada != "Todas":
    df_update = df_update[df_update["_entrada_norm"] == filtro_entrada]

if busca.strip():
    termo = busca.strip().lower()
    mask = (
        df_update["_estabelecimento_norm"].str.lower().str.contains(termo, na=False) |
        df_update["_categoria_norm"].str.lower().str.contains(termo, na=False) |
        df_update["_detalhes_norm"].str.lower().str.contains(termo, na=False) |
        (
            df_update["Whatsapp"].astype(str).str.lower().str.contains(termo, na=False)
            if "Whatsapp" in df_update.columns else False
        )
    )
    df_update = df_update[mask]

max_itens = 40
df_update = df_update.head(max_itens)

if df_update.empty:
    st.info("Nenhum lançamento encontrado para atualização.")
else:
    for idx, row in df_update.iterrows():
        estabelecimento = row["Estabelecimento"] if "Estabelecimento" in row else "-"
        valor_txt = formatar_brl(row["_valor_num"])
        entrada = row["Entrada"] if "Entrada" in row else "-"
        categoria = row["Categoria"] if "Categoria" in row else "-"
        status_atual = row["Status"] if "Status" in row else "-"
        detalhes = row["Detalhes"] if "Detalhes" in row else ""
        whatsapp = row["Whatsapp"] if "Whatsapp" in row else ""
        mes_txt = row["Mês"] if "Mês" in row else (row["Mes"] if "Mes" in row else "-")
        sheet_row = int(row["_sheet_row"])

        st.markdown('<div class="update-card">', unsafe_allow_html=True)

        info1, info2, info3, b1, b2, b3 = st.columns([3.3, 1.2, 1.2, 1.1, 1.1, 1.1])

        with info1:
            st.markdown(f'<div class="item-title">{estabelecimento}</div>', unsafe_allow_html=True)
            st.markdown(
                f"""
                <div class="item-meta">
                    <b>Mês:</b> {mes_txt}&nbsp;&nbsp;&nbsp;
                    <b>Entrada:</b> {entrada}&nbsp;&nbsp;&nbsp;
                    <b>Categoria:</b> {categoria}<br>
                    <b>Detalhes:</b> {detalhes if detalhes else "-"}<br>
                    <b>Whatsapp:</b> {whatsapp if whatsapp else "-"}
                </div>
                """,
                unsafe_allow_html=True
            )

        with info2:
            st.markdown('<div class="item-value-label">Valor</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="item-value">{valor_txt}</div>', unsafe_allow_html=True)

        with info3:
            st.markdown('<div class="item-value-label">Status atual</div>', unsafe_allow_html=True)
            st.markdown(
                f'<span class="{status_class(status_atual)}">{status_atual if status_atual else "Sem status"}</span>',
                unsafe_allow_html=True
            )

        with b1:
            if st.button("Pago", key=f"pago_{sheet_row}", use_container_width=True):
                try:
                    atualizar_status_na_planilha(sheet_row, "Pago")
                    st.success(f"Status da linha {sheet_row} atualizado para Pago.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar para Pago: {e}")

        with b2:
            if st.button("A Pagar", key=f"apagar_{sheet_row}", use_container_width=True):
                try:
                    atualizar_status_na_planilha(sheet_row, "A Pagar")
                    st.success(f"Status da linha {sheet_row} atualizado para A Pagar.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar para A Pagar: {e}")

        with b3:
            if st.button("A Receber", key=f"areceber_{sheet_row}", use_container_width=True):
                try:
                    atualizar_status_na_planilha(sheet_row, "A Receber")
                    st.success(f"Status da linha {sheet_row} atualizado para A Receber.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar para A Receber: {e}")

        st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="small-note">Dica: compartilhe a planilha com o e-mail da service account do Google para o dashboard conseguir ler e atualizar o status.</div>', unsafe_allow_html=True)
