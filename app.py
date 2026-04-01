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
    .stApp { background: #f6f7fb; }

    .block-container {
        max-width: 1450px;
        padding-top: 2.4rem !important;
    }

    .main-title {
        text-align: center;
        font-size: 2.6rem;
        font-weight: 800;
        color: #14213d;
    }

    .main-subtitle {
        text-align: center;
        color: #667085;
        margin-bottom: 1.5rem;
    }

    .kpi-card {
        background: #fff;
        border-left: 5px solid #e91e63;
        border-radius: 18px;
        padding: 1rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }

    .kpi-title {
        font-size: 0.9rem;
        color: #666;
    }

    .kpi-value {
        font-size: 1.8rem;
        font-weight: 800;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# LOGO REDONDA (🔥 AQUI ESTÁ O SEGREDO)
# =========================================================
def encontrar_logo():
    for nome in LOGO_CANDIDATES:
        p = Path(nome)
        if p.exists():
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
            <div style="display:flex; justify-content:center; margin-bottom:15px;">
                <div style="
                    width:110px;
                    height:110px;
                    background:#000;
                    border-radius:50%;
                    display:flex;
                    align-items:center;
                    justify-content:center;
                    box-shadow: 0 10px 25px rgba(0,0,0,0.25);
                ">
                    <img src="data:{mime_type};base64,{img_base64}"
                         style="width:55%; height:auto;">
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    except:
        pass

# =========================================================
# GOOGLE
# =========================================================
@st.cache_resource
def conectar():
    creds = Credentials.from_service_account_info(
        st.secrets["google"], scopes=SCOPES
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=30)
def carregar():
    client = conectar()
    ws = client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)
    df = pd.DataFrame(ws.get_all_records())

    df["_valor"] = (
        df["Valor"]
        .astype(str)
        .str.replace("R$", "")
        .str.replace(".", "")
        .str.replace(",", ".")
        .astype(float)
    )

    return df, ws

def atualizar(ws, row, status):
    headers = ws.row_values(1)
    col = headers.index("Status") + 1
    ws.update_cell(row, col, status)
    st.cache_data.clear()

# =========================================================
# HEADER
# =========================================================
render_logo()

st.markdown('<div class="main-title">Despesas & Gastos OPPI</div>', unsafe_allow_html=True)
st.markdown('<div class="main-subtitle">Gestão financeira inteligente</div>', unsafe_allow_html=True)

# =========================================================
# LOAD
# =========================================================
df, ws = carregar()

# =========================================================
# KPIs
# =========================================================
receita = df[df["Entrada"]=="Receita"]["_valor"].sum()
despesa = df[df["Entrada"]=="Despesa"]["_valor"].sum()
saldo = receita - despesa

c1, c2, c3 = st.columns(3)

c1.markdown(f'<div class="kpi-card"><div class="kpi-title">Receita</div><div class="kpi-value">R$ {receita:,.2f}</div></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="kpi-card"><div class="kpi-title">Despesa</div><div class="kpi-value">R$ {despesa:,.2f}</div></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="kpi-card"><div class="kpi-title">Saldo</div><div class="kpi-value">R$ {saldo:,.2f}</div></div>', unsafe_allow_html=True)

# =========================================================
# GRÁFICO
# =========================================================
st.markdown("### 📊 Categoria")
cat = df.groupby("Categoria")["_valor"].sum().reset_index()
fig = px.bar(cat, x="Categoria", y="_valor")
st.plotly_chart(fig, use_container_width=True)

# =========================================================
# ATUALIZAR STATUS
# =========================================================
st.markdown("### ✏️ Atualizar Status")

for i, row in df.iterrows():
    c1, c2, c3, c4 = st.columns([3,1,1,1])

    c1.write(f"{row['Estabelecimento']} - R$ {row['_valor']:,.2f} - {row['Status']}")

    if c2.button("Pago", key=f"p{i}"):
        atualizar(ws, i+2, "Pago")
        st.rerun()

    if c3.button("A Pagar", key=f"a{i}"):
        atualizar(ws, i+2, "A Pagar")
        st.rerun()

    if c4.button("A Receber", key=f"r{i}"):
        atualizar(ws, i+2, "A Receber")
        st.rerun()
