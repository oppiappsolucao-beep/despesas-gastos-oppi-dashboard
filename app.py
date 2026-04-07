import base64
import html
import mimetypes
import re
from pathlib import Path
from datetime import datetime, timedelta
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
# ESTILO (AJUSTADO COM SCROLL NO HOVER)
# =========================================================
st.markdown("""
<style>
.status-hover-box {
    position: absolute;
    left: 0;
    top: calc(100% + 8px);
    width: 320px;
    max-height: 360px;
    overflow-y: auto;
    overflow-x: hidden;
    background: #ffffff;
    border: 1px solid #e8eaf2;
    border-radius: 16px;
    box-shadow: 0 14px 28px rgba(20, 20, 43, 0.12);
    padding: 0.9rem 1rem;
    z-index: 9999;
    opacity: 0;
    visibility: hidden;
    transform: translateY(6px);
    transition: opacity 0.15s ease, transform 0.15s ease, visibility 0.15s ease;
    pointer-events: auto;
}

.status-hover-box::-webkit-scrollbar {
    width: 8px;
}

.status-hover-box::-webkit-scrollbar-thumb {
    background: #cfd5e3;
    border-radius: 999px;
}

.status-hover-box::-webkit-scrollbar-track {
    background: transparent;
}

.status-mini-wrap:hover .status-hover-box,
.status-hover-box:hover {
    opacity: 1;
    visibility: visible;
    transform: translateY(0);
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# HELPERS
# =========================================================
def normalizar_coluna(col):
    return str(col or "").replace("\ufeff", "").replace("\xa0", " ").strip()

def slug_coluna(col):
    col = normalizar_coluna(col).lower()
    col = re.sub(r"[^a-z0-9]+", "", col)
    return col

def encontrar_coluna(df, candidatos):
    mapa = {slug_coluna(c): c for c in df.columns}
    for cand in candidatos:
        if slug_coluna(cand) in mapa:
            return mapa[slug_coluna(cand)]
    return None

def parse_brl(valor):
    try:
        return float(str(valor).replace("R$", "").replace(".", "").replace(",", "."))
    except:
        return 0.0

def formatar_brl(v):
    return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def montar_detalhes_status_html(df_base, status_nome):
    base = df_base[df_base["_status"].str.lower() == status_nome.lower()].copy()

    qtd = len(base)
    valor_total = formatar_brl(base["_valor_num"].sum())

    itens_html = ""
    for _, r in base.iterrows():
        nome = html.escape(str(r["_estabelecimento"]) or "-")
        valor = formatar_brl(r["_valor_num"])
        itens_html += f'<div class="status-hover-item">• {nome}: {valor}</div>'

    return f"""
    <div class="status-hover-title">{status_nome}</div>
    <div class="status-hover-line">Qtd: {qtd}</div>
    <div class="status-hover-line">Valor total: {valor_total}</div>
    <div class="status-hover-subtitle">Todos os registros:</div>
    {itens_html}
    """

# =========================================================
# GOOGLE SHEETS
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
    ws = conectar().open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    df["_valor_num"] = df["Valor"].apply(parse_brl)
    df["_status"] = df["Status"]
    df["_estabelecimento"] = df["Estabelecimento"]
    df["_entrada"] = df["Entrada"]
    df["_categoria"] = df["Categoria"]

    return df

# =========================================================
# LOAD
# =========================================================
df = carregar()

# =========================================================
# KPIs
# =========================================================
total_pago = df[df["_status"].str.lower() == "pago"]["_valor_num"].sum()
qtd_pago = len(df[df["_status"].str.lower() == "pago"])

total_apagar = df[df["_status"].str.lower() == "a pagar"]["_valor_num"].sum()
qtd_apagar = len(df[df["_status"].str.lower() == "a pagar"])

total_areceber = df[df["_status"].str.lower() == "a receber"]["_valor_num"].sum()
qtd_areceber = len(df[df["_status"].str.lower() == "a receber"])

# =========================================================
# HOVERS
# =========================================================
hover_pago = montar_detalhes_status_html(df, "Pago")
hover_apagar = montar_detalhes_status_html(df, "A Pagar")
hover_areceber = montar_detalhes_status_html(df, "A Receber")

# =========================================================
# UI
# =========================================================
c1, c2, c3 = st.columns(3)

with c1:
    st.markdown(f"""
    <div class="status-mini-wrap">
        <div class="status-mini-card pago">
            <div class="status-mini-title">Pago</div>
            <div class="status-mini-value">{qtd_pago}</div>
            <div class="status-mini-caption">Passe o mouse</div>
        </div>
        <div class="status-hover-box">{hover_pago}</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="status-mini-wrap">
        <div class="status-mini-card apagar">
            <div class="status-mini-title">A Pagar</div>
            <div class="status-mini-value">{qtd_apagar}</div>
            <div class="status-mini-caption">Passe o mouse</div>
        </div>
        <div class="status-hover-box">{hover_apagar}</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="status-mini-wrap">
        <div class="status-mini-card areceber">
            <div class="status-mini-title">A Receber</div>
            <div class="status-mini-value">{qtd_areceber}</div>
            <div class="status-mini-caption">Passe o mouse</div>
        </div>
        <div class="status-hover-box">{hover_areceber}</div>
    </div>
    """, unsafe_allow_html=True)
