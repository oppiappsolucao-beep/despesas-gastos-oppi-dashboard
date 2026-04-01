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
.stApp { background: #f6f7fb; }
.main-title { text-align:center; font-size:2.6rem; font-weight:800; }
.kpi-card {
    background:white; border-radius:20px; padding:1rem;
    border-left:5px solid #e91e63;
}
.kpi-card.roxo { border-left-color:#7c3aed; }
.kpi-card.verde { border-left-color:#10b981; }
.kpi-card.rosa { border-left-color:#e91e63; }
</style>
""", unsafe_allow_html=True)

# =========================================================
# HELPERS
# =========================================================
def parse_brl(valor):
    try:
        return float(str(valor).replace("R$", "").replace(".", "").replace(",", "."))
    except:
        return 0.0

def formatar_brl(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def parse_data_br(valor):
    try:
        return pd.to_datetime(valor, dayfirst=True)
    except:
        return pd.NaT

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
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    df["_valor"] = df["Valor"].apply(parse_brl)
    df["_status"] = df["Status"].astype(str)
    df["_entrada"] = df["Entrada"].astype(str)

    return df, ws

def atualizar(ws, row, status):
    headers = ws.row_values(1)
    col = headers.index("Status") + 1
    ws.update_cell(row, col, status)
    st.cache_data.clear()

# =========================================================
# HEADER
# =========================================================
st.markdown('<div class="main-title">Despesas & Gastos OPPI</div>', unsafe_allow_html=True)

df, ws = carregar()

# =========================================================
# KPIs
# =========================================================
receita = df[df["_entrada"]=="Receita"]["_valor"].sum()
despesa = df[df["_entrada"]=="Despesa"]["_valor"].sum()
saldo = receita - despesa

col1,col2,col3 = st.columns(3)

col1.metric("Receita", formatar_brl(receita))
col2.metric("Despesa", formatar_brl(despesa))
col3.metric("Saldo", formatar_brl(saldo))

# =========================================================
# GRÁFICO
# =========================================================
st.markdown("## 📊 Categoria")

cat = df.groupby("Categoria")["_valor"].sum().reset_index()

fig = px.bar(cat, x="Categoria", y="_valor")
st.plotly_chart(fig, use_container_width=True)

# =========================================================
# ATUALIZAR STATUS
# =========================================================
st.markdown("## ✏️ Atualizar Status")

for i, row in df.iterrows():
    c1,c2,c3,c4 = st.columns([3,1,1,1])

    c1.write(f"{row['Estabelecimento']} - {formatar_brl(row['_valor'])} - {row['Status']}")

    if c2.button("Pago", key=f"p{i}"):
        atualizar(ws, i+2, "Pago")
        st.rerun()

    if c3.button("A Pagar", key=f"a{i}"):
        atualizar(ws, i+2, "A Pagar")
        st.rerun()

    if c4.button("A Receber", key=f"r{i}"):
        atualizar(ws, i+2, "A Receber")
        st.rerun()
