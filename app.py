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
    page_title="Gestão Financeira Oppi",
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
# ESTILO (NÃO ALTERADO)
# =========================================================
st.markdown("""<style>
    .stApp { background: #f6f7fb; }
    .block-container {
        max-width: 1450px;
        padding-top: 3.6rem !important;
        padding-bottom: 2rem;
    }
    .logo-wrap {
        display: flex;
        justify-content: center;
        margin-top: 0.35rem;
        margin-bottom: 0.8rem;
    }
    .logo-circle {
        width: 142px;
        height: 142px;
        border-radius: 50%;
        overflow: hidden;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 8px 22px rgba(0, 0, 0, 0.12);
    }
    .logo-circle img {
        width: 100%;
        height: 100%;
        object-fit: cover;
    }
    .main-title {
        text-align: center;
        font-size: 2.6rem;
        font-weight: 800;
        color: #14213d;
        margin-bottom: 0.2rem;
    }
    .main-subtitle {
        text-align: center;
        font-size: 1.08rem;
        color: #667085;
        margin-bottom: 1.6rem;
    }
</style>""", unsafe_allow_html=True)

# =========================================================
# LOGO (NÃO ALTERADO)
# =========================================================
def encontrar_logo():
    for nome in LOGO_CANDIDATES:
        p = Path(nome)
        if p.exists():
            return p
    return None

def render_logo():
    logo = encontrar_logo()
    if not logo:
        return

    img_bytes = logo.read_bytes()
    mime = mimetypes.guess_type(str(logo))[0]
    b64 = base64.b64encode(img_bytes).decode()

    st.markdown(f"""
    <div class="logo-wrap">
        <div class="logo-circle">
            <img src="data:{mime};base64,{b64}">
        </div>
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# HEADER (ALTERADO AQUI)
# =========================================================
render_logo()

st.markdown('<div class="main-title">Gestão Financeira Oppi</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="main-subtitle">Gestão financeira de receitas, despesas e status de pagamento</div>',
    unsafe_allow_html=True
)

# =========================================================
# GOOGLE SHEETS (NÃO ALTERADO)
# =========================================================
@st.cache_resource
def conectar():
    creds = Credentials.from_service_account_info(
        st.secrets["google"],
        scopes=SCOPES
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=30)
def carregar():
    client = conectar()
    ws = client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)
    df = pd.DataFrame(ws.get_all_records())
    return df

# =========================================================
# LOAD
# =========================================================
try:
    df = carregar()
except Exception as e:
    st.error("Erro ao conectar com a planilha.")
    st.exception(e)
    st.stop()

st.dataframe(df)
