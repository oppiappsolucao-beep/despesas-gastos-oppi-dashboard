import base64
import mimetypes
import re
from pathlib import Path

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

SHEET_ID = "1cQU5tNwSoiepTPHx_Qc7ZF1PcaER2gstW_dZQ0eCrB4"
WORKSHEET_NAME = "Página1"

# pode ser png, jpg, jpeg ou webp
LOGO_CANDIDATES = [
    "logo_oppi.png",
    "logo_oppi.jpg",
    "logo_oppi.jpeg",
    "logo_oppi.webp",
]

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
        max-width: 1450px;
        padding-top: 2.4rem !important;
        padding-bottom: 2rem;
    }

    .logo-wrap {
        display: flex;
        justify-content: center;
        margin-bottom: 0.65rem;
    }

    .logo-wrap img {
        max-width: 120px;
        width: 100%;
        height: auto;
        display: block;
    }

    .main-title {
        text-align: center;
        font-size: 2.6rem;
        font-weight: 800;
        color: #14213d;
        margin-bottom: 0.2rem;
        line-height: 1.1;
    }

    .main-subtitle {
        text-align: center;
        font-size: 1.08rem;
        color: #667085;
        margin-bottom: 1.6rem;
    }

    .top-divider, .section-divider {
        width: 100%;
        height: 18px;
        background: #ffffff;
        border: 1px solid #ececf3;
        border-radius: 999px;
        margin: 0.8rem 0 1.35rem 0;
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
        padding: 1.05rem 1.15rem 0.95rem 1.15rem;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        min-height: 150px;
    }

    .kpi-card.roxo { border-left-color: #7c3aed; }
    .kpi-card.rosa { border-left-color: #e91e63; }
    .kpi-card.verde { border-left-color: #10b981; }
    .kpi-card.azul { border-left-color: #3b82f6; }
    .kpi-card.laranja { border-left-color: #f59e0b; }

    .kpi-title {
        font-size: 1rem;
        font-weight: 700;
        color: #28314f;
        margin-bottom: 0.8rem;
    }

    .kpi-value {
        font-size: 2rem;
        font-weight: 800;
        color: #081b4b;
        line-height: 1.05;
        margin-bottom: 0.72rem;
    }

    .kpi-caption {
        font-size: 0.92rem;
        color: #667085;
    }

    .section-title {
        font-size: 1.38rem;
        font-weight: 800;
        color: #14213d;
        margin-bottom: 0.3rem;
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
        padding: 1.15rem;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        margin-bottom: 1rem;
    }

    .item-title {
        font-size: 1.35rem;
        font-weight: 800;
        color: #0b1d4d;
        margin-bottom: 0.35rem;
    }

    .item-meta {
        color: #64748b;
        font-size: 0.96rem;
        line-height: 1.65;
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
        font-size: 1.28rem;
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
        font-size: 0.88rem;
        color: #6b7280;
        margin-top: 0.45rem;
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
def normalizar_coluna(col):
    col = str(col or "")
    col = col.replace("\ufeff", "").replace("\xa0", " ").strip()
    return col

def slug_coluna(col):
    col = normalizar_coluna(col).lower()
    col = (
        col.replace("ã", "a")
        .replace("á", "a")
        .replace("à", "a")
        .replace("â", "a")
        .replace("é", "e")
        .replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ô", "o")
        .replace("õ", "o")
        .replace("ú", "u")
        .replace("ç", "c")
    )
    col = re.sub(r"[^a-z0-9]+", "", col)
    return col

def encontrar_coluna(df, candidatos):
    mapa = {slug_coluna(c): c for c in df.columns}
    for cand in candidatos:
        slug = slug_coluna(cand)
        if slug in mapa:
            return mapa[slug]
    return None

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
    return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def parse_data_br(valor):
    if pd.isna(valor):
        return pd.NaT
    s = str(valor).strip()
    if not s:
        return pd.NaT
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
    s = str(status or "").strip().lower()
    if s == "pago":
        return "status-pill status-pago"
    if s == "a pagar":
        return "status-pill status-apagar"
    if s == "a receber":
        return "status-pill status-areceber"
    return "status-pill"

def encontrar_logo():
    for nome in LOGO_CANDIDATES:
        p = Path(nome)
        if p.exists() and p.is_file():
            return p
    return None

def render_logo():
    logo_path = encontrar_logo()
    if not logo_path:
        return

    try:
        img_bytes = logo_path.read_bytes()
        mime_type = mimetypes.guess_type(str(logo_path))[0] or "image/png"
        img_base64 = base64.b64encode(img_bytes).decode("utf-8")
        st.markdown(
            f"""
            <div class="logo-wrap">
                <img src="data:{mime_type};base64,{img_base64}" alt="Logo Oppi">
            </div>
            """,
            unsafe_allow_html=True
        )
    except Exception:
        # não quebra o app se a logo estiver inválida
        pass

# =========================================================
# GOOGLE SHEETS
# =========================================================
@st.cache_resource(show_spinner=False)
def conectar():
    creds = Credentials.from_service_account_info(
        st.secrets["google"],
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    return client

@st.cache_data(ttl=30, show_spinner=False)
def carregar():
    client = conectar()
    ws = client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame(), {}

    headers = [normalizar_coluna(h) for h in values[0]]
    rows = values[1:]

    rows_pad = []
    for row in rows:
        if len(row) < len(headers):
            row = row + [""] * (len(headers) - len(row))
        elif len(row) > len(headers):
            row = row[:len(headers)]
        rows_pad.append(row)

    df = pd.DataFrame(rows_pad, columns=headers)

    for c in df.columns:
        df[c] = df[c].astype(str).apply(lambda x: x.strip())

    col_mes = encontrar_coluna(df, ["Mês", "Mes", "Data", "Data do mês", "Data do mes"])
    col_estabelecimento = encontrar_coluna(df, ["Estabelecimento", "Empresa", "Nome"])
    col_valor = encontrar_coluna(df, ["Valor", "Valor total", "Preço", "Preco"])
    col_entrada = encontrar_coluna(df, ["Entrada", "Tipo", "Movimento"])
    col_categoria = encontrar_coluna(df, ["Categoria", "Grupo"])
    col_status = encontrar_coluna(df, ["Status", "Situação", "Situacao"])
    col_detalhes = encontrar_coluna(df, ["Detalhes", "Descrição", "Descricao", "Observação", "Observacao"])
    col_whatsapp = encontrar_coluna(df, ["Whatsapp", "WhatsApp", "Telefone"])

    df["_mes_raw"] = df[col_mes] if col_mes else ""
    df["_data_mes"] = df["_mes_raw"].apply(parse_data_br) if col_mes else pd.NaT
    df["_mes_label"] = df["_data_mes"].apply(extrair_mes_label) if col_mes else "Sem data"

    df["_estabelecimento"] = df[col_estabelecimento] if col_estabelecimento else ""
    df["_valor_num"] = df[col_valor].apply(parse_brl) if col_valor else 0.0
    df["_entrada"] = df[col_entrada].astype(str).str.strip() if col_entrada else ""
    df["_categoria"] = df[col_categoria].astype(str).str.strip() if col_categoria else ""
    df["_status"] = df[col_status].astype(str).str.strip() if col_status else ""
    df["_detalhes"] = df[col_detalhes].astype(str).str.strip() if col_detalhes else ""
    df["_whatsapp"] = df[col_whatsapp].astype(str).str.strip() if col_whatsapp else ""

    df["_sheet_row"] = range(2, len(df) + 2)

    meta = {
        "status_col_name": col_status,
    }
    return df, meta

def atualizar_status(sheet_row, novo_status):
    client = conectar()
    ws = client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)
    headers = [normalizar_coluna(h) for h in ws.row_values(1)]

    status_col_name = None
    for h in headers:
        if slug_coluna(h) in [slug_coluna("Status"), slug_coluna("Situação"), slug_coluna("Situacao")]:
            status_col_name = h
            break

    if not status_col_name:
        raise ValueError("Coluna 'Status' não encontrada na planilha.")

    status_col_idx = headers.index(status_col_name) + 1
    ws.update_cell(sheet_row, status_col_idx, novo_status)
    st.cache_data.clear()

# =========================================================
# HEADER
# =========================================================
render_logo()

st.markdown('<div class="main-title">Despesas & Gastos OPPI</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="main-subtitle">Gestão financeira de receitas, despesas e status de pagamento</div>',
    unsafe_allow_html=True
)
st.markdown('<div class="top-divider"></div>', unsafe_allow_html=True)

# =========================================================
# LOAD
# =========================================================
try:
    df, meta = carregar()
except Exception as e:
    st.error("Erro ao conectar com a planilha do Google Sheets.")
    st.exception(e)
    st.stop()

if df.empty:
    st.warning("A planilha está vazia.")
    st.stop()

if not meta.get("status_col_name"):
    st.error("A coluna 'Status' não foi encontrada na planilha. Confira o cabeçalho.")
    st.stop()

# =========================================================
# FILTROS
# =========================================================
meses_opcoes = ["Todos"] + sorted(
    [m for m in df["_mes_label"].dropna().unique().tolist() if m != "Sem data"]
)

estab_opcoes = ["Todos"] + sorted([x for x in df["_estabelecimento"].unique().tolist() if str(x).strip()])
categoria_opcoes = ["Todas"] + sorted([x for x in df["_categoria"].unique().tolist() if str(x).strip()])
entrada_opcoes = ["Todas"] + sorted([x for x in df["_entrada"].unique().tolist() if str(x).strip()])

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
    df_filtrado = df_filtrado[df_filtrado["_estabelecimento"] == filtro_estab]

if filtro_categoria != "Todas":
    df_filtrado = df_filtrado[df_filtrado["_categoria"] == filtro_categoria]

if filtro_entrada != "Todas":
    df_filtrado = df_filtrado[df_filtrado["_entrada"] == filtro_entrada]

# =========================================================
# KPIs
# =========================================================
total_registros = len(df_filtrado)

total_receitas = df_filtrado.loc[
    df_filtrado["_entrada"].str.lower() == "receita", "_valor_num"
].sum()

total_despesas = df_filtrado.loc[
    df_filtrado["_entrada"].str.lower() == "despesa", "_valor_num"
].sum()

saldo = total_receitas - total_despesas

total_pago = df_filtrado.loc[
    df_filtrado["_status"].str.lower() == "pago", "_valor_num"
].sum()

total_apagar = df_filtrado.loc[
    df_filtrado["_status"].str.lower() == "a pagar", "_valor_num"
].sum()

total_areceber = df_filtrado.loc[
    df_filtrado["_status"].str.lower() == "a receber", "_valor_num"
].sum()

c1, c2, c3, c4, c5, c6 = st.columns(6)

cards = [
    (c1, "Registros", str(total_registros), "total de lançamentos filtrados", "roxo"),
    (c2, "Receitas", formatar_brl(total_receitas), "soma das receitas", "verde"),
    (c3, "Despesas", formatar_brl(total_despesas), "soma das despesas", "rosa"),
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
    qtd_pago = (df_filtrado["_status"].str.lower() == "pago").sum()
    qtd_apagar = (df_filtrado["_status"].str.lower() == "a pagar").sum()
    qtd_areceber = (df_filtrado["_status"].str.lower() == "a receber").sum()

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
        df_filtrado.groupby("_categoria", dropna=False)["_valor_num"]
        .sum()
        .reset_index()
        .rename(columns={"_categoria": "Categoria", "_valor_num": "Valor"})
    )
    base_categoria = base_categoria[base_categoria["Categoria"].astype(str).str.strip() != ""]

    if not base_categoria.empty:
        fig_cat = px.bar(
            base_categoria.sort_values("Valor", ascending=False),
            x="Categoria",
            y="Valor",
            text="Valor"
        )
        fig_cat.update_traces(texttemplate="R$ %{y:,.2f}", textposition="outside")
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
        df_filtrado.groupby("_status", dropna=False)["_valor_num"]
        .sum()
        .reset_index()
        .rename(columns={"_status": "Status", "_valor_num": "Valor"})
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
    placeholder="Ex.: Valeria, OpenAI, internet, mídia...",
    label_visibility="collapsed"
)

df_update = df.copy()

if filtro_mes != "Todos":
    df_update = df_update[df_update["_mes_label"] == filtro_mes]
if filtro_estab != "Todos":
    df_update = df_update[df_update["_estabelecimento"] == filtro_estab]
if filtro_categoria != "Todas":
    df_update = df_update[df_update["_categoria"] == filtro_categoria]
if filtro_entrada != "Todas":
    df_update = df_update[df_update["_entrada"] == filtro_entrada]

if busca.strip():
    termo = busca.strip().lower()
    mask = (
        df_update["_estabelecimento"].astype(str).str.lower().str.contains(termo, na=False) |
        df_update["_categoria"].astype(str).str.lower().str.contains(termo, na=False) |
        df_update["_detalhes"].astype(str).str.lower().str.contains(termo, na=False) |
        df_update["_whatsapp"].astype(str).str.lower().str.contains(termo, na=False)
    )
    df_update = df_update[mask]

df_update = df_update.head(40)

if df_update.empty:
    st.info("Nenhum lançamento encontrado para atualização.")
else:
    for _, row in df_update.iterrows():
        estabelecimento = row["_estabelecimento"] if str(row["_estabelecimento"]).strip() else "-"
        valor_txt = formatar_brl(row["_valor_num"])
        entrada = row["_entrada"] if str(row["_entrada"]).strip() else "-"
        categoria = row["_categoria"] if str(row["_categoria"]).strip() else "-"
        status_atual = row["_status"] if str(row["_status"]).strip() else "Sem status"
        detalhes = row["_detalhes"] if str(row["_detalhes"]).strip() else "-"
        whatsapp = row["_whatsapp"] if str(row["_whatsapp"]).strip() else "-"
        mes_txt = row["_mes_raw"] if str(row["_mes_raw"]).strip() else "-"
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
                    <b>Detalhes:</b> {detalhes}<br>
                    <b>Whatsapp:</b> {whatsapp}
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
                f'<span class="{status_class(status_atual)}">{status_atual}</span>',
                unsafe_allow_html=True
            )

        with b1:
            if st.button("Pago", key=f"pago_{sheet_row}", use_container_width=True):
                try:
                    atualizar_status(sheet_row, "Pago")
                    st.success(f"Status da linha {sheet_row} atualizado para Pago.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar para Pago: {e}")

        with b2:
            if st.button("A Pagar", key=f"apagar_{sheet_row}", use_container_width=True):
                try:
                    atualizar_status(sheet_row, "A Pagar")
                    st.success(f"Status da linha {sheet_row} atualizado para A Pagar.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar para A Pagar: {e}")

        with b3:
            if st.button("A Receber", key=f"areceber_{sheet_row}", use_container_width=True):
                try:
                    atualizar_status(sheet_row, "A Receber")
                    st.success(f"Status da linha {sheet_row} atualizado para A Receber.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar para A Receber: {e}")

        st.markdown('</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="small-note">Se algo não carregar, confira se a planilha continua compartilhada com o e-mail da service account como Editor.</div>',
    unsafe_allow_html=True
)
