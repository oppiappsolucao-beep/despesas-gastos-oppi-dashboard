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

TZ = ZoneInfo("America/Sao_Paulo")

GOOGLE_CREDS = {
    "type": "service_account",
    "project_id": "dashboard-despesas-gastos-oppi",
    "private_key_id": "3974a0efe71f649841a6f316605c3bd5a50754ef",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDiEtHKfNBdk8Li\nETrtw0TUPbmT9JT8jGxKTX2+40+O5G/jC9xU1TokbWpenbqYTH//ZmWj1vu5DLQD\n0eaXD477uOU8UZZn2+ZLCVOGhQH/4bR0RM5n76y5uL5WLjJynf/dSiplfBzDuQXR\nHzmGyx/PSj+gvTQyW6D92LQ/tCKdOy+SrpWuihiFBSsjfaOSGS/yWaLnDECfX6DR\nDuKAZofmlPJge8fCa9sP/G0yYikC1rjB5NOwcGpMtVgG3rX6rFBJit/bg8kVblZD\nztE7dEDmMYc8pa56N2FlSXbyq6ve4BD3hsau2mdyKMP8XJpsdbdNo2s/sKgGR4A2\n+PKuMNg7AgMBAAECggEAFN5+rH3rHWHnhC22mxfPtPWU7QCPkRfeMAzvn7ByIEaC\njmecFkSv6+7uk8wxqAXf08UxesQv7O9fd4ubCRs8xfK8G15Ytdh9BzSil9dgmnnM\nyhlHIh19r1zlGfJRYkOnLySKipKNDgN4pjGHHCQiGP7Jct6KzZJ1ARnndPEwnmmz\nP2vswvQFn43dp6/h73As8ofQig/xi0hovH47KGqgUZdDOfC0NZMucE03VpGbZ1TN\n4bRwNh4JhORUEPg3H9pxHfzENAQKobCWFQzG6zLkPfnrPDsB3RWF7v6KPcPu+Olw\nI4BQbSJpEcQPsq4294KNgx8nsS32u7z//I74qrp+UQKBgQD0RFKB3qhEwO58Bj9s\nADsD7j3YNZ7pRiI4uKIKPnTOzJ7hyDeEIApq9nIfGU23O2GLor8xpIIsJTLkUn58\nthGvWoFosl95w7qZ0Ym6rxG8BKCPU2Arz83gjYltIHEJAAfFeaBwcfNdFvqeOjQz\noi1dp95Hvb3lzfoaRHUQ3j8JIwKBgQDs7sdZB0G/tZ8VUQXaGSvJnHie1YbzQ1KL\n7qlE4Rw4DjTtjIxYshJTCLRw7JmlbGuAnShG6v2ADQsh9JlLeByfehLTd+fqTFPI\nZQxLKhgVn1pL3gFevpSJkDENjrP/4upLMrp1dMhg5pZTTypAwHOwUJyTFExRmxHF\nPlRNMA/CCQKBgF7fAmSqhBRgEsBc9NkPpdw69g45lUTpFnWNUHJGG7wOQU9UIivQ\n/frZSS3G+CZIi/Rd+4BecqiOsht35uStGmVO86AkV2zFln4Toji9sleiPHIuYdXi\nWgXzMwMNbJmgR2RtfuDtgSYQvLojxQ6g2Jndjzmx+kV9ILx/BjDNARKdAoGBAONu\njaLTCXT55Vvz63cgpFyiO1LUSvcmD43NKWS55XmVgY7pVCsru9VCzNp88zvMqCDM\nOsZgebg6TQ5qGeBMysT2zC17sv3ACMia3sMkA/x1e5rJ32zP6gtmgv+tlPEzI43N\ngxiOYm5JydDsc/W2BxcfOj0gxeWrwdIhc5CoaufpAoGATl1DvcRsjSun6PELyZ68\nXfL0sAphLUyal6Gli2ez9C9uoe5fQ6+HRCO9o/uQnnkQtljrPajr6aGQ3wzJ7dc+\nPT8y3JPBO/VBg4BUPWrdU1h011TKqw/BBYXJKqepkQ2ftNihhYyVuY3U7IW+SXEQ\nnwynqWw9XKwws5EkPW1TdkE=\n-----END PRIVATE KEY-----\n",
    "client_email": "streamlit-dashboard@dashboard-despesas-gastos-oppi.iam.gserviceaccount.com",
    "client_id": "107995700087066475943",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/streamlit-dashboard%40dashboard-despesas-gastos-oppi.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

@st.cache_resource(show_spinner=False)
def conectar():
    creds = Credentials.from_service_account_info(
        GOOGLE_CREDS,
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    return client

# =========================================================
# ESTILO
# =========================================================
st.markdown("""
<style>
    .stApp {
        background: #f6f7fb;
    }

    .block-container {
        max-width: 1500px;
        padding-top: 3.2rem !important;
        padding-bottom: 2rem;
    }

    .logo-wrap {
        display: flex;
        justify-content: center;
        margin-top: 0.2rem;
        margin-bottom: 0.8rem;
    }

    .logo-circle {
        width: 138px;
        height: 138px;
        border-radius: 50%;
        overflow: hidden;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 10px 24px rgba(0, 0, 0, 0.12);
        background: #fff;
    }

    .logo-circle img {
        width: 100%;
        height: 100%;
        object-fit: cover;
        object-position: center center;
        display: block;
    }

    .main-title {
        text-align: center;
        font-size: 2.55rem;
        font-weight: 800;
        color: #14213d;
        margin-bottom: 0.2rem;
        line-height: 1.1;
    }

    .main-subtitle {
        text-align: center;
        font-size: 1.02rem;
        color: #667085;
        margin-bottom: 1.55rem;
    }

    .top-divider, .section-divider {
        width: 100%;
        height: 16px;
        background: #ffffff;
        border: 1px solid #ececf3;
        border-radius: 999px;
        margin: 0.8rem 0 1.25rem 0;
    }

    .section-title {
        font-size: 1.34rem;
        font-weight: 800;
        color: #14213d;
        margin-bottom: 0.22rem;
    }

    .section-text {
        color: #677185;
        font-size: 0.96rem;
        margin-bottom: 1rem;
    }

    .filter-label {
        font-size: 0.94rem;
        color: #2f3552;
        font-weight: 700;
        margin-bottom: 0.3rem;
    }

    .section-chip {
        display: inline-block;
        padding: 0.35rem 0.8rem;
        border-radius: 999px;
        background: #eef2ff;
        color: #4938b7;
        font-size: 0.82rem;
        font-weight: 700;
        margin-bottom: 0.65rem;
    }

    .kpi-card {
        background: #ffffff;
        border: 1px solid #ececf3;
        border-left: 6px solid #e91e63;
        border-radius: 22px;
        padding: 1rem 1rem 0.95rem 1rem;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        min-height: 180px;
        height: 180px;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
        overflow: hidden;
    }

    .kpi-card.compacto {
        min-height: 168px;
        height: 168px;
    }

    .kpi-card.alto {
        min-height: 208px;
        height: 208px;
    }

    .kpi-card.roxo { border-left-color: #7c3aed; }
    .kpi-card.rosa { border-left-color: #e91e63; }
    .kpi-card.verde { border-left-color: #10b981; }
    .kpi-card.azul { border-left-color: #3b82f6; }
    .kpi-card.laranja { border-left-color: #f59e0b; }
    .kpi-card.vermelho { border-left-color: #ef4444; }
    .kpi-card.cinza { border-left-color: #64748b; }

    .kpi-title {
        font-size: 0.96rem;
        font-weight: 700;
        color: #28314f;
        margin-bottom: 0.78rem;
        line-height: 1.15;
    }

    .kpi-value {
        font-size: clamp(1.15rem, 1.5vw, 1.7rem);
        font-weight: 800;
        color: #081b4b;
        line-height: 1.12;
        margin-bottom: 0.58rem;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: clip;
        max-width: 100%;
    }

    .kpi-value.small {
        font-size: clamp(1.05rem, 1.3vw, 1.4rem);
    }

    .kpi-caption {
        font-size: 0.9rem;
        color: #667085;
        line-height: 1.4;
    }

    .kpi-helper {
        font-size: 0.84rem;
        color: #8a93a5;
        margin-top: 0.3rem;
        line-height: 1.35;
    }

    .saldo-pos {
        color: #0f9f6f !important;
    }

    .saldo-neg {
        color: #d92d20 !important;
    }

    .alert-card {
        background: linear-gradient(135deg, #ffffff 0%, #fbfbff 100%);
        border: 1px solid #ececf3;
        border-left: 6px solid #ef4444;
        border-radius: 22px;
        padding: 1rem 1rem 0.95rem 1rem;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        margin-bottom: 1rem;
    }

    .alert-title {
        font-size: 1.02rem;
        font-weight: 800;
        color: #14213d;
        margin-bottom: 0.55rem;
    }

    .alert-line {
        font-size: 0.93rem;
        color: #5f6b7a;
        margin-bottom: 0.28rem;
        line-height: 1.45;
    }

    .next-due-card {
        background: #ffffff;
        border: 1px solid #ececf3;
        border-left: 6px solid #3b82f6;
        border-radius: 22px;
        padding: 1rem;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        min-height: 208px;
        height: 208px;
        overflow: hidden;
    }

    .next-due-main {
        font-size: 1.42rem;
        font-weight: 800;
        color: #081b4b;
        margin-bottom: 0.28rem;
        line-height: 1.1;
    }

    .next-due-sub {
        font-size: 0.94rem;
        color: #667085;
        margin-bottom: 0.6rem;
        line-height: 1.4;
    }

    .next-due-list {
        font-size: 0.88rem;
        color: #5f6b7a;
        line-height: 1.48;
        max-height: 96px;
        overflow-y: auto;
        padding-right: 0.25rem;
    }

    .status-mini-wrap {
        position: relative;
        min-width: 0;
    }

    .status-mini-card {
        background: #ffffff;
        border: 1px solid #ececf3;
        border-radius: 18px;
        padding: 0.9rem 0.85rem;
        min-height: 150px;
        height: 150px;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        display: flex;
        flex-direction: column;
        justify-content: center;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
        cursor: default;
        overflow: hidden;
    }

    .status-mini-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 24px rgba(20, 20, 43, 0.09);
    }

    .status-mini-card.pago {
        border-left: 5px solid #10b981;
    }

    .status-mini-card.apagar {
        border-left: 5px solid #f59e0b;
    }

    .status-mini-card.areceber {
        border-left: 5px solid #7c3aed;
    }

    .status-mini-card.vencido {
        border-left: 5px solid #ef4444;
    }

    .status-mini-card.recebido {
        border-left: 5px solid #10b981;
    }

    .status-mini-title {
        font-size: 0.9rem;
        font-weight: 700;
        color: #28314f;
        margin-bottom: 0.45rem;
        line-height: 1.1;
    }

    .status-mini-value {
        font-size: 1.68rem;
        font-weight: 800;
        color: #081b4b;
        line-height: 1.05;
        margin-bottom: 0.35rem;
    }

    .status-mini-caption {
        font-size: 0.78rem;
        color: #667085;
        line-height: 1.2;
    }

    .status-hover-box {
        position: absolute;
        left: 0;
        top: calc(100% + 8px);
        width: 330px;
        max-height: 360px;
        overflow-y: auto;
        background: #ffffff;
        border: 1px solid #e8eaf2;
        border-radius: 16px;
        box-shadow: 0 14px 28px rgba(20, 20, 43, 0.12);
        padding: 0.9rem 1rem;
        z-index: 60;
        opacity: 0;
        visibility: hidden;
        transform: translateY(6px);
        transition: opacity 0.15s ease, transform 0.15s ease, visibility 0.15s ease;
        pointer-events: none;
    }

    .status-mini-wrap:hover .status-hover-box {
        opacity: 1;
        visibility: visible;
        transform: translateY(0);
    }

    .status-hover-title {
        font-size: 0.95rem;
        font-weight: 800;
        color: #14213d;
        margin-bottom: 0.45rem;
    }

    .status-hover-line {
        font-size: 0.87rem;
        color: #5f6b7a;
        margin-bottom: 0.15rem;
    }

    .status-hover-subtitle {
        font-size: 0.88rem;
        font-weight: 700;
        color: #28314f;
        margin-top: 0.55rem;
        margin-bottom: 0.3rem;
    }

    .status-hover-item {
        font-size: 0.84rem;
        color: #667085;
        margin-bottom: 0.14rem;
    }

    .update-card {
        background: #ffffff;
        border: 1px solid #ececf3;
        border-radius: 24px;
        padding: 1.15rem;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        margin-bottom: 1rem;
    }

    .create-card {
        background: #ffffff;
        border: 1px solid #ececf3;
        border-radius: 24px;
        padding: 1.2rem;
        box-shadow: 0 6px 18px rgba(20, 20, 43, 0.05);
        margin-bottom: 1rem;
    }

    .item-title {
        font-size: 1.25rem;
        font-weight: 800;
        color: #0b1d4d;
        margin-bottom: 0.35rem;
    }

    .item-meta {
        color: #64748b;
        font-size: 0.95rem;
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
        font-size: 1.22rem;
        font-weight: 800;
        color: #081b4b;
        margin-bottom: 0.45rem;
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
        background: #fff1dc;
        color: #b45309;
    }

    .status-areceber {
        background: #efe3ff;
        color: #6d28d9;
    }

    .status-vencido {
        background: #fee4e2;
        color: #c62828;
    }

    .status-recebido {
        background: #dff7e8;
        color: #118a43;
    }

    .small-note {
        font-size: 0.88rem;
        color: #6b7280;
        margin-top: 0.45rem;
    }

    .edit-hint {
        font-size: 0.78rem;
        color: #8b93a7;
        margin-top: 0.2rem;
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

    div[data-testid="stDataFrame"] {
        border: 1px solid #ececf3;
        border-radius: 18px;
        overflow: hidden;
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

def formatar_valor_planilha(v):
    return f"{float(v):.2f}".replace(".", ",")

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

def formatar_data_curta(data):
    if pd.isna(data):
        return "-"
    try:
        return pd.to_datetime(data).strftime("%d/%m/%Y")
    except Exception:
        return "-"

def cor_saldo(valor):
    return "saldo-pos" if valor >= 0 else "saldo-neg"

def texto_plural(qtd, singular, plural=None):
    if plural is None:
        plural = singular + "s"
    return singular if qtd == 1 else plural

def normalizar_entrada(valor):
    s = str(valor or "").strip().lower()
    if "receita" in s:
        return "receita"
    if "despesa" in s:
        return "despesa"
    return s

def normalizar_status_base(valor):
    s = str(valor or "").strip().lower()
    if s in ["pago", "paga", "pagamento efetuado"]:
        return "pago"
    if s in ["a pagar", "apagar"]:
        return "a pagar"
    if s in ["recebido", "recebida"]:
        return "recebido"
    if s in ["a receber", "areceber"]:
        return "a receber"
    return s

def status_exibicao_por_tipo(entrada_norm, status_base, data_ref, hoje_ref):
    if entrada_norm == "receita":
        if status_base in ["pago", "recebido"]:
            return "Recebido"
        if status_base == "a receber":
            return "A Receber"
        return "A Receber"

    if entrada_norm == "despesa":
        if status_base == "pago":
            return "Pago"
        if status_base == "a pagar":
            if pd.notna(data_ref) and data_ref < hoje_ref:
                return "Vencido"
            return "A Pagar"
        return "A Pagar"

    return str(status_base).title() if status_base else "Sem status"

def status_class(status):
    s = str(status or "").strip().lower()
    if s == "pago":
        return "status-pill status-pago"
    if s == "a pagar":
        return "status-pill status-apagar"
    if s == "a receber":
        return "status-pill status-areceber"
    if s == "vencido":
        return "status-pill status-vencido"
    if s == "recebido":
        return "status-pill status-recebido"
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
                <div class="logo-circle">
                    <img src="data:{mime_type};base64,{img_base64}" alt="Logo Oppi">
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    except Exception:
        pass

def montar_detalhes_status_html(df_base, status_nome):
    base = df_base[df_base["_status_exibicao"].str.lower() == status_nome.lower()].copy()
    if base.empty:
        return """
        <div class="status-hover-title">Sem registros</div>
        <div class="status-hover-line">Nenhum item encontrado nesse status.</div>
        """
    qtd = len(base)
    valor_total = formatar_brl(base["_valor_num"].sum())
    itens_html = ""
    for _, r in base.iterrows():
        nome = html.escape(str(r["_estabelecimento"]).strip() or "-")
        data_txt = formatar_data_curta(r["_data_mes"])
        valor = formatar_brl(r["_valor_num"])
        itens_html += f'<div class="status-hover-item">• {nome} | {data_txt} | {valor}</div>'

    html_final = f"""
    <div class="status-hover-title">{html.escape(status_nome)}</div>
    <div class="status-hover-line">Qtd: {qtd}</div>
    <div class="status-hover-line">Valor total: {valor_total}</div>
    <div class="status-hover-subtitle">Todos os registros:</div>
    {itens_html}
    """
    return html_final.strip()

def opcoes_status_por_tipo(tipo):
    return ["Recebido", "A Receber"] if tipo == "Receita" else ["Pago", "A Pagar"]

# =========================================================
# GOOGLE SHEETS
# =========================================================
@st.cache_resource(show_spinner=False)
def conectar():
    creds = Credentials.from_service_account_info(
        GOOGLE_CREDS,
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
    df["_data_ref"] = pd.to_datetime(df["_data_mes"], errors="coerce").dt.date

    meta = {
        "status_col_name": col_status,
        "valor_col_name": col_valor,
        "mes_col_name": col_mes,
        "estab_col_name": col_estabelecimento,
        "entrada_col_name": col_entrada,
        "categoria_col_name": col_categoria,
        "detalhes_col_name": col_detalhes,
        "whatsapp_col_name": col_whatsapp,
        "headers": headers,
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

def atualizar_valor(sheet_row, novo_valor):
    if novo_valor is None or str(novo_valor).strip() == "":
        raise ValueError("Digite um valor antes de salvar.")

    valor_num = parse_brl(novo_valor)

    client = conectar()
    ws = client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)
    headers = [normalizar_coluna(h) for h in ws.row_values(1)]

    valor_col_name = None
    for h in headers:
        if slug_coluna(h) in [
            slug_coluna("Valor"),
            slug_coluna("Valor total"),
            slug_coluna("Preço"),
            slug_coluna("Preco"),
        ]:
            valor_col_name = h
            break

    if not valor_col_name:
        raise ValueError("Coluna 'Valor' não encontrada na planilha.")

    valor_col_idx = headers.index(valor_col_name) + 1
    ws.update_cell(sheet_row, valor_col_idx, formatar_valor_planilha(valor_num))
    st.cache_data.clear()

def adicionar_lancamento(meta, data_str, estabelecimento, valor, tipo, categoria, status, detalhes, whatsapp):
    client = conectar()
    ws = client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)

    headers = meta["headers"]
    nova_linha = [""] * len(headers)

    mapa_valores = {
        meta.get("mes_col_name"): data_str,
        meta.get("estab_col_name"): estabelecimento,
        meta.get("valor_col_name"): formatar_valor_planilha(parse_brl(valor)),
        meta.get("entrada_col_name"): tipo,
        meta.get("categoria_col_name"): categoria,
        meta.get("status_col_name"): status,
        meta.get("detalhes_col_name"): detalhes,
        meta.get("whatsapp_col_name"): whatsapp,
    }

    for col_name, valor_coluna in mapa_valores.items():
        if col_name and col_name in headers:
            idx = headers.index(col_name)
            nova_linha[idx] = valor_coluna

    ws.append_row(nova_linha, value_input_option="USER_ENTERED")
    st.cache_data.clear()

# =========================================================
# HEADER
# =========================================================
render_logo()

st.markdown('<div class="main-title">Gestão Financeira Oppi</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="main-subtitle">Painel financeiro com visão de realizado, projetado, vencimentos, previsões e atualização direta na planilha</div>',
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

if not meta.get("valor_col_name"):
    st.error("A coluna 'Valor' não foi encontrada na planilha. Confira o cabeçalho.")
    st.stop()

# =========================================================
# SESSION STATE - NOVO LANÇAMENTO
# =========================================================
hoje = datetime.now(TZ).date()

if "novo_tipo" not in st.session_state:
    st.session_state["novo_tipo"] = "Receita"
if "novo_status" not in st.session_state:
    st.session_state["novo_status"] = "Recebido"
if "nova_data" not in st.session_state:
    st.session_state["nova_data"] = hoje
if "novo_estabelecimento" not in st.session_state:
    st.session_state["novo_estabelecimento"] = ""
if "novo_valor" not in st.session_state:
    st.session_state["novo_valor"] = ""
if "nova_categoria" not in st.session_state:
    st.session_state["nova_categoria"] = ""
if "novo_whatsapp" not in st.session_state:
    st.session_state["novo_whatsapp"] = ""
if "novo_detalhes" not in st.session_state:
    st.session_state["novo_detalhes"] = ""

def ao_mudar_tipo():
    tipo = st.session_state["novo_tipo"]
    opcoes = opcoes_status_por_tipo(tipo)
    if st.session_state.get("novo_status") not in opcoes:
        st.session_state["novo_status"] = opcoes[0]

amanha = hoje + timedelta(days=1)
fim_7_dias = hoje + timedelta(days=7)

# =========================================================
# PRÉ-PROCESSAMENTO
# =========================================================
df["_entrada_norm"] = df["_entrada"].apply(normalizar_entrada)
df["_status_base"] = df["_status"].apply(normalizar_status_base)

df["_status_exibicao"] = df.apply(
    lambda r: status_exibicao_por_tipo(
        r["_entrada_norm"],
        r["_status_base"],
        r["_data_ref"],
        hoje
    ),
    axis=1
)

# =========================================================
# FILTROS
# =========================================================
meses_opcoes = ["Todos"] + sorted(
    [m for m in df["_mes_label"].dropna().unique().tolist() if m != "Sem data"]
)
estab_opcoes = ["Todos"] + sorted([x for x in df["_estabelecimento"].unique().tolist() if str(x).strip()])
categoria_opcoes = ["Todas"] + sorted([x for x in df["_categoria"].unique().tolist() if str(x).strip()])
entrada_opcoes = ["Todas"] + sorted([x.title() for x in df["_entrada_norm"].unique().tolist() if str(x).strip()])

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
    df_filtrado = df_filtrado[df_filtrado["_entrada_norm"] == filtro_entrada.lower()]

# =========================================================
# KPIs - RESUMO DO MÊS
# =========================================================
total_registros = len(df_filtrado)
qtd_receitas = int((df_filtrado["_entrada_norm"] == "receita").sum())
qtd_despesas = int((df_filtrado["_entrada_norm"] == "despesa").sum())

total_receitas = df_filtrado.loc[df_filtrado["_entrada_norm"] == "receita", "_valor_num"].sum()
total_despesas = df_filtrado.loc[df_filtrado["_entrada_norm"] == "despesa", "_valor_num"].sum()

resultado_mes = total_receitas - total_despesas
margem_mes = (resultado_mes / total_receitas * 100) if total_receitas > 0 else 0.0

ticket_medio_receita = total_receitas / qtd_receitas if qtd_receitas > 0 else 0.0
ticket_medio_despesa = total_despesas / qtd_despesas if qtd_despesas > 0 else 0.0

maior_despesa = df_filtrado.loc[df_filtrado["_entrada_norm"] == "despesa", "_valor_num"].max() if qtd_despesas > 0 else 0.0

base_cat_desp = (
    df_filtrado[df_filtrado["_entrada_norm"] == "despesa"]
    .groupby("_categoria", dropna=False)["_valor_num"]
    .sum()
    .reset_index()
    .sort_values("_valor_num", ascending=False)
)

if not base_cat_desp.empty:
    maior_categoria_nome = str(base_cat_desp.iloc[0]["_categoria"]).strip() or "Sem categoria"
    maior_categoria_valor = float(base_cat_desp.iloc[0]["_valor_num"])
else:
    maior_categoria_nome = "-"
    maior_categoria_valor = 0.0

# =========================================================
# KPIs - REALIZADO / PROJETADO / STATUS
# =========================================================
despesa_paga = df_filtrado.loc[
    (df_filtrado["_entrada_norm"] == "despesa") &
    (df_filtrado["_status_base"] == "pago"),
    "_valor_num"
].sum()

receita_recebida = df_filtrado.loc[
    (df_filtrado["_entrada_norm"] == "receita") &
    (df_filtrado["_status_base"].isin(["pago", "recebido"])),
    "_valor_num"
].sum()

saldo_realizado = receita_recebida - despesa_paga

total_pago = despesa_paga
total_recebido = receita_recebida

total_apagar = df_filtrado.loc[
    (df_filtrado["_entrada_norm"] == "despesa") &
    (df_filtrado["_status_base"] == "a pagar"),
    "_valor_num"
].sum()

total_areceber = df_filtrado.loc[
    (df_filtrado["_entrada_norm"] == "receita") &
    (df_filtrado["_status_base"] == "a receber"),
    "_valor_num"
].sum()

total_vencido = df_filtrado.loc[
    (df_filtrado["_entrada_norm"] == "despesa") &
    (df_filtrado["_status_base"] == "a pagar") &
    (df_filtrado["_data_ref"].notna()) &
    (df_filtrado["_data_ref"] < hoje),
    "_valor_num"
].sum()

qtd_pago = int(((df_filtrado["_entrada_norm"] == "despesa") & (df_filtrado["_status_base"] == "pago")).sum())
qtd_recebido = int(((df_filtrado["_entrada_norm"] == "receita") & (df_filtrado["_status_base"].isin(["pago", "recebido"]))).sum())
qtd_apagar = int(((df_filtrado["_entrada_norm"] == "despesa") & (df_filtrado["_status_base"] == "a pagar")).sum())
qtd_areceber = int(((df_filtrado["_entrada_norm"] == "receita") & (df_filtrado["_status_base"] == "a receber")).sum())

qtd_vencido = int((
    (df_filtrado["_entrada_norm"] == "despesa") &
    (df_filtrado["_status_base"] == "a pagar") &
    (df_filtrado["_data_ref"].notna()) &
    (df_filtrado["_data_ref"] < hoje)
).sum())

saldo_projetado_mes = saldo_realizado + total_areceber - total_apagar

# =========================================================
# PREVISÕES
# =========================================================
df_prev = df_filtrado.copy()
df_prev = df_prev[df_prev["_data_ref"].notna()].copy()

ganhos_hoje = df_prev.loc[
    (df_prev["_entrada_norm"] == "receita") &
    (df_prev["_data_ref"] == hoje),
    "_valor_num"
].sum()

despesas_hoje = df_prev.loc[
    (df_prev["_entrada_norm"] == "despesa") &
    (df_prev["_data_ref"] == hoje),
    "_valor_num"
].sum()

saldo_prev_hoje = ganhos_hoje - despesas_hoje

ganhos_7_dias = df_prev.loc[
    (df_prev["_entrada_norm"] == "receita") &
    (df_prev["_data_ref"] >= amanha) &
    (df_prev["_data_ref"] <= fim_7_dias),
    "_valor_num"
].sum()

despesas_7_dias = df_prev.loc[
    (df_prev["_entrada_norm"] == "despesa") &
    (df_prev["_data_ref"] >= amanha) &
    (df_prev["_data_ref"] <= fim_7_dias),
    "_valor_num"
].sum()

saldo_prev_7_dias = ganhos_7_dias - despesas_7_dias

base_prox_venc = df_filtrado[
    (df_filtrado["_entrada_norm"] == "despesa") &
    (df_filtrado["_status_base"] == "a pagar") &
    (df_filtrado["_data_ref"].notna()) &
    (df_filtrado["_data_ref"] >= hoje)
].copy().sort_values("_data_ref", ascending=True)

if not base_prox_venc.empty:
    prox = base_prox_venc.iloc[0]
    prox_data = formatar_data_curta(prox["_data_mes"])
    prox_estabelecimento = str(prox["_estabelecimento"]).strip() or "-"
    prox_valor = formatar_brl(prox["_valor_num"])
else:
    prox_data = "-"
    prox_estabelecimento = "Sem contas futuras"
    prox_valor = "R$ 0,00"

top5_vencer = base_prox_venc.head(5).copy()

# =========================================================
# ALERTAS INTELIGENTES
# =========================================================
contas_hoje_df = df_filtrado[
    (df_filtrado["_entrada_norm"] == "despesa") &
    (df_filtrado["_status_base"] == "a pagar") &
    (df_filtrado["_data_ref"] == hoje)
].copy()

contas_amanha_df = df_filtrado[
    (df_filtrado["_entrada_norm"] == "despesa") &
    (df_filtrado["_status_base"] == "a pagar") &
    (df_filtrado["_data_ref"] == amanha)
].copy()

receb_atrasado_df = df_filtrado[
    (df_filtrado["_entrada_norm"] == "receita") &
    (df_filtrado["_status_base"] == "a receber") &
    (df_filtrado["_data_ref"].notna()) &
    (df_filtrado["_data_ref"] < hoje)
].copy()

qtd_contas_hoje = len(contas_hoje_df)
valor_contas_hoje = contas_hoje_df["_valor_num"].sum()
qtd_contas_amanha = len(contas_amanha_df)
valor_contas_amanha = contas_amanha_df["_valor_num"].sum()
valor_receb_atrasado = receb_atrasado_df["_valor_num"].sum()

categoria_alerta_txt = ""
if total_despesas > 0 and maior_categoria_valor > 0:
    perc_maior_cat = (maior_categoria_valor / total_despesas) * 100
    if perc_maior_cat >= 25:
        categoria_alerta_txt = (
            f'Categoria "{maior_categoria_nome}" já representa '
            f'{perc_maior_cat:.1f}% das despesas.'
        ).replace(".", ",")

alertas = []
if qtd_contas_hoje > 0:
    alertas.append(
        f"Hoje vencem {qtd_contas_hoje} {texto_plural(qtd_contas_hoje, 'conta')} somando {formatar_brl(valor_contas_hoje)}."
    )
if qtd_contas_amanha > 0:
    alertas.append(
        f"Amanhã vencem {qtd_contas_amanha} {texto_plural(qtd_contas_amanha, 'conta')} somando {formatar_brl(valor_contas_amanha)}."
    )
if total_vencido > 0:
    alertas.append(f"Você tem {formatar_brl(total_vencido)} em despesas vencidas.")
if valor_receb_atrasado > 0:
    alertas.append(f"Há {formatar_brl(valor_receb_atrasado)} em recebimentos atrasados.")
if total_despesas > total_receitas and total_receitas > 0:
    alertas.append("As despesas já ultrapassaram as receitas no período filtrado.")
if categoria_alerta_txt:
    alertas.append(categoria_alerta_txt)

# =========================================================
# LINHA 1 — RESUMO DO MÊS
# =========================================================
st.markdown('<div class="section-chip">Resumo do período</div>', unsafe_allow_html=True)
r1, r2, r3, r4, r5 = st.columns(5)

with r1:
    st.markdown(
        f"""
        <div class="kpi-card roxo compacto">
            <div class="kpi-title">Registros</div>
            <div class="kpi-value">{total_registros}</div>
            <div class="kpi-caption">{qtd_receitas} receitas • {qtd_despesas} despesas</div>
            <div class="kpi-helper">Total de lançamentos no filtro atual</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with r2:
    st.markdown(
        f"""
        <div class="kpi-card verde compacto">
            <div class="kpi-title">Receitas</div>
            <div class="kpi-value">{formatar_brl(total_receitas)}</div>
            <div class="kpi-caption">Ticket médio: {formatar_brl(ticket_medio_receita)}</div>
            <div class="kpi-helper">Total de entradas registradas</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with r3:
    st.markdown(
        f"""
        <div class="kpi-card rosa compacto">
            <div class="kpi-title">Despesas</div>
            <div class="kpi-value">{formatar_brl(total_despesas)}</div>
            <div class="kpi-caption">Ticket médio: {formatar_brl(ticket_medio_despesa)}</div>
            <div class="kpi-helper">Total de saídas registradas</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with r4:
    st.markdown(
        f"""
        <div class="kpi-card {'verde' if resultado_mes >= 0 else 'vermelho'} compacto">
            <div class="kpi-title">Saldo</div>
            <div class="kpi-value {cor_saldo(resultado_mes)}">{formatar_brl(resultado_mes)}</div>
            <div class="kpi-caption">Resultado do período</div>
            <div class="kpi-helper">Receitas menos despesas</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with r5:
    st.markdown(
        f"""
        <div class="kpi-card azul compacto">
            <div class="kpi-title">Margem %</div>
            <div class="kpi-value {'saldo-pos' if margem_mes >= 0 else 'saldo-neg'}">{str(round(margem_mes, 1)).replace('.', ',')}%</div>
            <div class="kpi-caption">Saúde do período</div>
            <div class="kpi-helper">Resultado ÷ receitas</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================================================
# ALERTAS
# =========================================================
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    f"""
    <div class="alert-card">
        <div class="alert-title">⚠️ Alertas inteligentes</div>
        {"".join([f'<div class="alert-line">• {html.escape(a)}</div>' for a in alertas]) if alertas else '<div class="alert-line">Nenhum alerta crítico encontrado no filtro atual.</div>'}
    </div>
    """,
    unsafe_allow_html=True
)

# =========================================================
# LINHA 2 — SITUAÇÃO FINANCEIRA
# =========================================================
hover_pago = montar_detalhes_status_html(df_filtrado, "Pago")
hover_recebido = montar_detalhes_status_html(df_filtrado, "Recebido")
hover_apagar = montar_detalhes_status_html(df_filtrado, "A Pagar")
hover_areceber = montar_detalhes_status_html(df_filtrado, "A Receber")
hover_vencido = montar_detalhes_status_html(df_filtrado, "Vencido")

st.markdown('<div class="section-chip">Situação financeira</div>', unsafe_allow_html=True)
s1, s2, s3, s4, s5 = st.columns(5)

with s1:
    st.markdown(
        f"""
        <div class="kpi-card verde compacto">
            <div class="kpi-title">Total já pago</div>
            <div class="kpi-value">{formatar_brl(total_pago)}</div>
            <div class="kpi-caption">{qtd_pago} despesas pagas</div>
            <div class="kpi-helper">Somente despesas com status Pago</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with s2:
    st.markdown(
        f"""
        <div class="kpi-card azul compacto">
            <div class="kpi-title">Total recebido</div>
            <div class="kpi-value">{formatar_brl(total_recebido)}</div>
            <div class="kpi-caption">{qtd_recebido} receitas recebidas</div>
            <div class="kpi-helper">Somente receitas com status Recebido</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with s3:
    st.markdown(
        f"""
        <div class="kpi-card laranja compacto">
            <div class="kpi-title">Total em aberto para pagar</div>
            <div class="kpi-value">{formatar_brl(total_apagar)}</div>
            <div class="kpi-caption">{qtd_apagar} lançamentos</div>
            <div class="kpi-helper">Despesas ainda pendentes</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with s4:
    st.markdown(
        f"""
        <div class="kpi-card roxo compacto">
            <div class="kpi-title">Total a receber</div>
            <div class="kpi-value">{formatar_brl(total_areceber)}</div>
            <div class="kpi-caption">{qtd_areceber} lançamentos</div>
            <div class="kpi-helper">Receitas ainda previstas</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with s5:
    st.markdown(
        f"""
        <div class="kpi-card {'verde' if saldo_projetado_mes >= 0 else 'vermelho'} compacto">
            <div class="kpi-title">Saldo projetado</div>
            <div class="kpi-value {cor_saldo(saldo_projetado_mes)}">{formatar_brl(saldo_projetado_mes)}</div>
            <div class="kpi-caption">Recebido + a receber - a pagar</div>
            <div class="kpi-helper">Visão projetada do período filtrado</div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("<br>", unsafe_allow_html=True)
mini1, mini2, mini3, mini4, mini5 = st.columns(5)

with mini1:
    st.markdown(
        f"""
        <div class="status-mini-wrap">
            <div class="status-mini-card pago">
                <div class="status-mini-title">Pago</div>
                <div class="status-mini-value">{qtd_pago}</div>
                <div class="status-mini-caption">Despesas pagas</div>
            </div>
            <div class="status-hover-box">{hover_pago}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with mini2:
    st.markdown(
        f"""
        <div class="status-mini-wrap">
            <div class="status-mini-card recebido">
                <div class="status-mini-title">Recebido</div>
                <div class="status-mini-value">{qtd_recebido}</div>
                <div class="status-mini-caption">Receitas recebidas</div>
            </div>
            <div class="status-hover-box">{hover_recebido}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with mini3:
    st.markdown(
        f"""
        <div class="status-mini-wrap">
            <div class="status-mini-card apagar">
                <div class="status-mini-title">A Pagar</div>
                <div class="status-mini-value">{qtd_apagar}</div>
                <div class="status-mini-caption">Despesas pendentes</div>
            </div>
            <div class="status-hover-box">{hover_apagar}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with mini4:
    st.markdown(
        f"""
        <div class="status-mini-wrap">
            <div class="status-mini-card areceber">
                <div class="status-mini-title">A Receber</div>
                <div class="status-mini-value">{qtd_areceber}</div>
                <div class="status-mini-caption">Receitas pendentes</div>
            </div>
            <div class="status-hover-box">{hover_areceber}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with mini5:
    st.markdown(
        f"""
        <div class="status-mini-wrap">
            <div class="status-mini-card vencido">
                <div class="status-mini-title">Vencido</div>
                <div class="status-mini-value">{qtd_vencido}</div>
                <div class="status-mini-caption">Despesas atrasadas</div>
            </div>
            <div class="status-hover-box">{hover_vencido}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================================================
# LINHA 3 — REALIZADO / PROJETADO / PREVISÕES
# =========================================================
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-chip">Realizado e previsões</div>', unsafe_allow_html=True)

p1, p2, p3, p4, p5 = st.columns(5)

with p1:
    st.markdown(
        f"""
        <div class="kpi-card azul alto">
            <div class="kpi-title">Receita recebida</div>
            <div class="kpi-value">{formatar_brl(receita_recebida)}</div>
            <div class="kpi-caption">Entradas já realizadas</div>
            <div class="kpi-helper">Somente receitas com status Recebido</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with p2:
    st.markdown(
        f"""
        <div class="kpi-card rosa alto">
            <div class="kpi-title">Despesa paga</div>
            <div class="kpi-value">{formatar_brl(despesa_paga)}</div>
            <div class="kpi-caption">Saídas já realizadas</div>
            <div class="kpi-helper">Somente despesas com status Pago</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with p3:
    st.markdown(
        f"""
        <div class="kpi-card {'verde' if saldo_realizado >= 0 else 'vermelho'} alto">
            <div class="kpi-title">Saldo realizado</div>
            <div class="kpi-value {cor_saldo(saldo_realizado)}">{formatar_brl(saldo_realizado)}</div>
            <div class="kpi-caption">Recebido menos pago</div>
            <div class="kpi-helper">Visão real do caixa já realizado</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with p4:
    st.markdown(
        f"""
        <div class="kpi-card verde alto">
            <div class="kpi-title">Hoje</div>
            <div class="kpi-caption">Entradas previstas: {formatar_brl(ganhos_hoje)}</div>
            <div class="kpi-caption">Saídas previstas: {formatar_brl(despesas_hoje)}</div>
            <div class="kpi-value small {cor_saldo(saldo_prev_hoje)}">Saldo: {formatar_brl(saldo_prev_hoje)}</div>
            <div class="kpi-helper">Baseado nas datas de hoje</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with p5:
    st.markdown(
        f"""
        <div class="kpi-card laranja alto">
            <div class="kpi-title">Próximos 7 dias</div>
            <div class="kpi-caption">Entradas previstas: {formatar_brl(ganhos_7_dias)}</div>
            <div class="kpi-caption">Saídas previstas: {formatar_brl(despesas_7_dias)}</div>
            <div class="kpi-value small {cor_saldo(saldo_prev_7_dias)}">Saldo: {formatar_brl(saldo_prev_7_dias)}</div>
            <div class="kpi-helper">De amanhã até +7 dias</div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("<br>", unsafe_allow_html=True)
nv1, nv2 = st.columns([1.25, 1])

with nv1:
    if top5_vencer.empty:
        top5_html = '<div class="next-due-list">Nenhuma conta futura encontrada.</div>'
    else:
        linhas = []
        for _, row in top5_vencer.iterrows():
            dt = formatar_data_curta(row["_data_mes"])
            est = html.escape(str(row["_estabelecimento"]).strip() or "-")
            val = formatar_brl(row["_valor_num"])
            linhas.append(f"• {dt} — {est} — {val}")
        top5_html = f'<div class="next-due-list">{"<br>".join(linhas)}</div>'

    st.markdown(
        f"""
        <div class="next-due-card">
            <div class="kpi-title">Próximo vencimento</div>
            <div class="next-due-main">{prox_data}</div>
            <div class="next-due-sub">{html.escape(prox_estabelecimento)} • {prox_valor}</div>
            <div class="kpi-helper" style="margin-bottom:0.35rem;">Top 5 contas a vencer</div>
            {top5_html}
        </div>
        """,
        unsafe_allow_html=True
    )

with nv2:
    st.markdown(
        f"""
        <div class="next-due-card" style="border-left-color:#7c3aed;">
            <div class="kpi-title">Indicadores de gestão</div>
            <div class="kpi-caption">Maior despesa: <b>{formatar_brl(maior_despesa)}</b></div>
            <div class="kpi-caption">Maior categoria: <b>{html.escape(maior_categoria_nome)}</b></div>
            <div class="kpi-caption">Valor da categoria líder: <b>{formatar_brl(maior_categoria_valor)}</b></div>
            <div class="kpi-caption">Recebimentos atrasados: <b>{formatar_brl(valor_receb_atrasado)}</b></div>
            <div class="kpi-caption">Total previsto a entrar: <b>{formatar_brl(total_areceber)}</b></div>
            <div class="kpi-helper">Leitura rápida para decisão</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================================================
# NOVO LANÇAMENTO
# =========================================================
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">➕ Adicionar novo lançamento</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-text">Cadastre uma nova receita ou despesa e envie direto para a planilha.</div>',
    unsafe_allow_html=True
)

with st.container():
    st.markdown('<div class="create-card">', unsafe_allow_html=True)

    nl1, nl2, nl3, nl4 = st.columns(4)

    with nl1:
        nova_data = st.date_input(
            "Data",
            key="nova_data",
            format="DD/MM/YYYY"
        )

    with nl2:
        novo_estabelecimento = st.text_input(
            "Estabelecimento",
            key="novo_estabelecimento"
        )

    with nl3:
        novo_valor = st.text_input(
            "Valor",
            key="novo_valor",
            placeholder="Ex.: 1574,00"
        )

    with nl4:
        novo_tipo = st.selectbox(
            "Tipo",
            ["Receita", "Despesa"],
            key="novo_tipo",
            on_change=ao_mudar_tipo
        )

    opcoes_status = opcoes_status_por_tipo(st.session_state["novo_tipo"])
    if st.session_state["novo_status"] not in opcoes_status:
        st.session_state["novo_status"] = opcoes_status[0]

    nl5, nl6, nl7 = st.columns(3)

    with nl5:
        nova_categoria = st.text_input(
            "Categoria",
            key="nova_categoria"
        )

    with nl6:
        novo_status = st.selectbox(
            "Status",
            opcoes_status,
            key="novo_status"
        )

    with nl7:
        novo_whatsapp = st.text_input(
            "Whatsapp",
            key="novo_whatsapp"
        )

    novo_detalhes = st.text_area(
        "Detalhes",
        key="novo_detalhes",
        height=90
    )

    if st.button("Salvar lançamento", use_container_width=True, key="btn_salvar_novo_lancamento"):
        try:
            if not str(st.session_state["novo_estabelecimento"]).strip():
                raise ValueError("Preencha o estabelecimento.")

            if not str(st.session_state["novo_valor"]).strip():
                raise ValueError("Preencha o valor.")

            if parse_brl(st.session_state["novo_valor"]) <= 0:
                raise ValueError("Digite um valor maior que zero.")

            if not str(st.session_state["nova_categoria"]).strip():
                raise ValueError("Preencha a categoria.")

            data_formatada = pd.to_datetime(st.session_state["nova_data"]).strftime("%d/%m/%Y")

            adicionar_lancamento(
                meta=meta,
                data_str=data_formatada,
                estabelecimento=str(st.session_state["novo_estabelecimento"]).strip(),
                valor=st.session_state["novo_valor"],
                tipo=st.session_state["novo_tipo"],
                categoria=str(st.session_state["nova_categoria"]).strip(),
                status=st.session_state["novo_status"],
                detalhes=str(st.session_state["novo_detalhes"]).strip(),
                whatsapp=str(st.session_state["novo_whatsapp"]).strip(),
            )

            st.success("Novo lançamento adicionado com sucesso.")

            st.session_state["nova_data"] = hoje
            st.session_state["novo_estabelecimento"] = ""
            st.session_state["novo_valor"] = ""
            st.session_state["novo_tipo"] = "Receita"
            st.session_state["novo_status"] = "Recebido"
            st.session_state["nova_categoria"] = ""
            st.session_state["novo_whatsapp"] = ""
            st.session_state["novo_detalhes"] = ""

            st.rerun()

        except Exception as e:
            st.error(f"Erro ao adicionar lançamento: {e}")

    st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# GRÁFICOS
# =========================================================
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-chip">Visualização gráfica</div>', unsafe_allow_html=True)

g1, g2 = st.columns(2)

with g1:
    st.markdown('<div class="section-title">📊 Despesas por categoria</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-text">Barras horizontais facilitam a leitura das maiores categorias de gasto.</div>', unsafe_allow_html=True)

    base_categoria = (
        df_filtrado[df_filtrado["_entrada_norm"] == "despesa"]
        .groupby("_categoria", dropna=False)["_valor_num"]
        .sum()
        .reset_index()
        .rename(columns={"_categoria": "Categoria", "_valor_num": "Valor"})
    )
    base_categoria = base_categoria[base_categoria["Categoria"].astype(str).str.strip() != ""]

    if not base_categoria.empty:
        base_categoria = base_categoria.sort_values("Valor", ascending=True)
        fig_cat = px.bar(
            base_categoria,
            x="Valor",
            y="Categoria",
            orientation="h",
            text="Valor"
        )
        fig_cat.update_traces(texttemplate="R$ %{x:,.2f}", textposition="outside")
        fig_cat.update_layout(
            height=450,
            showlegend=False,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(l=20, r=20, t=20, b=20),
            xaxis_title="Valor",
            yaxis_title=""
        )
        fig_cat.update_xaxes(tickprefix="R$ ")
        st.plotly_chart(fig_cat, use_container_width=True)
    else:
        st.info("Sem dados de despesas para o gráfico de categoria.")

with g2:
    st.markdown('<div class="section-title">🍩 Distribuição por status</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-text">Veja a distribuição financeira por valor ou por quantidade.</div>', unsafe_allow_html=True)

    modo_status = st.radio(
        "Modo de exibição do status",
        ["Por valor", "Por quantidade"],
        horizontal=True,
        label_visibility="collapsed",
        key="modo_status_grafico"
    )

    if modo_status == "Por valor":
        base_status = (
            df_filtrado.groupby("_status_exibicao", dropna=False)["_valor_num"]
            .sum()
            .reset_index()
            .rename(columns={"_status_exibicao": "Status", "_valor_num": "Valor"})
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
                height=450,
                plot_bgcolor="white",
                paper_bgcolor="white",
                margin=dict(l=20, r=20, t=20, b=20),
                legend_title=""
            )
            st.plotly_chart(fig_status, use_container_width=True)
        else:
            st.info("Sem dados para o gráfico de status.")
    else:
        base_status_qtd = (
            df_filtrado.groupby("_status_exibicao", dropna=False)
            .size()
            .reset_index(name="Quantidade")
            .rename(columns={"_status_exibicao": "Status"})
        )
        base_status_qtd = base_status_qtd[base_status_qtd["Status"].astype(str).str.strip() != ""]
        if not base_status_qtd.empty:
            fig_status_qtd = px.pie(
                base_status_qtd,
                names="Status",
                values="Quantidade",
                hole=0.62
            )
            fig_status_qtd.update_traces(
                textinfo="label+value",
                texttemplate="%{label}<br>%{value}"
            )
            fig_status_qtd.update_layout(
                height=450,
                plot_bgcolor="white",
                paper_bgcolor="white",
                margin=dict(l=20, r=20, t=20, b=20),
                legend_title=""
            )
            st.plotly_chart(fig_status_qtd, use_container_width=True)
        else:
            st.info("Sem dados para o gráfico de status.")

st.markdown("<br>", unsafe_allow_html=True)
g3, g4 = st.columns(2)

with g3:
    st.markdown('<div class="section-title">📈 Fluxo por dia</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-text">Entradas e saídas por data para acompanhar o comportamento do período.</div>', unsafe_allow_html=True)

    base_fluxo = df_filtrado[df_filtrado["_data_mes"].notna()].copy()
    if not base_fluxo.empty:
        base_fluxo["Data"] = pd.to_datetime(base_fluxo["_data_mes"]).dt.strftime("%d/%m/%Y")
        base_fluxo["Tipo"] = base_fluxo["_entrada"].replace("", "Sem tipo")
        base_fluxo_fluxo = (
            base_fluxo.groupby(["Data", "Tipo"], dropna=False)["_valor_num"]
            .sum()
            .reset_index()
            .rename(columns={"_valor_num": "Valor"})
        )

        fig_fluxo = px.bar(
            base_fluxo_fluxo,
            x="Data",
            y="Valor",
            color="Tipo",
            barmode="group"
        )
        fig_fluxo.update_layout(
            height=430,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(l=20, r=20, t=20, b=20),
            xaxis_title="",
            yaxis_title="Valor",
            legend_title=""
        )
        fig_fluxo.update_yaxes(tickprefix="R$ ")
        st.plotly_chart(fig_fluxo, use_container_width=True)
    else:
        st.info("Sem datas válidas para montar o fluxo diário.")

with g4:
    st.markdown('<div class="section-title">🏆 Top 10 maiores despesas</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-text">As maiores saídas do período para identificar os pesos do caixa.</div>', unsafe_allow_html=True)

    base_top_desp = df_filtrado[df_filtrado["_entrada_norm"] == "despesa"].copy()
    if not base_top_desp.empty:
        base_top_desp["Label"] = base_top_desp["_estabelecimento"].replace("", "-")
        base_top_desp = (
            base_top_desp.sort_values("_valor_num", ascending=False)
            .head(10)
            .copy()
        )
        base_top_desp = base_top_desp.sort_values("_valor_num", ascending=True)
        fig_top = px.bar(
            base_top_desp,
            x="_valor_num",
            y="Label",
            orientation="h",
            text="_valor_num"
        )
        fig_top.update_traces(texttemplate="R$ %{x:,.2f}", textposition="outside")
        fig_top.update_layout(
            height=430,
            showlegend=False,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(l=20, r=20, t=20, b=20),
            xaxis_title="Valor",
            yaxis_title=""
        )
        fig_top.update_xaxes(tickprefix="R$ ")
        st.plotly_chart(fig_top, use_container_width=True)
    else:
        st.info("Sem despesas para montar o Top 10.")

# =========================================================
# TABELA DETALHADA
# =========================================================
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">📋 Lançamentos detalhados</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-text">Use os filtros rápidos para ver somente vencidos, a pagar, a receber, hoje, esta semana ou o período completo.</div>',
    unsafe_allow_html=True
)

t1, t2 = st.columns([1.1, 2.2])

with t1:
    filtro_rapido = st.selectbox(
        "Filtro rápido",
        [
            "Todos",
            "Só vencidos",
            "Só a pagar",
            "Só a receber",
            "Só recebidos",
            "Só pagos",
            "Hoje",
            "Próximos 7 dias",
            "Este mês"
        ]
    )

with t2:
    busca_tabela = st.text_input(
        "Buscar na tabela",
        placeholder="Ex.: fornecedor, categoria, whatsapp, observação..."
    )

df_tabela = df_filtrado.copy()

if filtro_rapido == "Só vencidos":
    df_tabela = df_tabela[df_tabela["_status_exibicao"] == "Vencido"]
elif filtro_rapido == "Só a pagar":
    df_tabela = df_tabela[df_tabela["_status_exibicao"] == "A Pagar"]
elif filtro_rapido == "Só a receber":
    df_tabela = df_tabela[df_tabela["_status_exibicao"] == "A Receber"]
elif filtro_rapido == "Só recebidos":
    df_tabela = df_tabela[df_tabela["_status_exibicao"] == "Recebido"]
elif filtro_rapido == "Só pagos":
    df_tabela = df_tabela[df_tabela["_status_exibicao"] == "Pago"]
elif filtro_rapido == "Hoje":
    df_tabela = df_tabela[df_tabela["_data_ref"] == hoje]
elif filtro_rapido == "Próximos 7 dias":
    df_tabela = df_tabela[
        (df_tabela["_data_ref"].notna()) &
        (df_tabela["_data_ref"] >= hoje) &
        (df_tabela["_data_ref"] <= fim_7_dias)
    ]
elif filtro_rapido == "Este mês":
    df_tabela = df_tabela[
        (df_tabela["_data_mes"].notna()) &
        (pd.to_datetime(df_tabela["_data_mes"]).dt.month == hoje.month) &
        (pd.to_datetime(df_tabela["_data_mes"]).dt.year == hoje.year)
    ]

if busca_tabela.strip():
    termo = busca_tabela.strip().lower()
    mask = (
        df_tabela["_estabelecimento"].astype(str).str.lower().str.contains(termo, na=False) |
        df_tabela["_categoria"].astype(str).str.lower().str.contains(termo, na=False) |
        df_tabela["_entrada"].astype(str).str.lower().str.contains(termo, na=False) |
        df_tabela["_status_exibicao"].astype(str).str.lower().str.contains(termo, na=False) |
        df_tabela["_detalhes"].astype(str).str.lower().str.contains(termo, na=False) |
        df_tabela["_whatsapp"].astype(str).str.lower().str.contains(termo, na=False)
    )
    df_tabela = df_tabela[mask]

if df_tabela.empty:
    st.info("Nenhum lançamento encontrado na tabela detalhada.")
else:
    tabela_exibir = df_tabela.copy()
    tabela_exibir["Data"] = tabela_exibir["_data_mes"].apply(formatar_data_curta)
    tabela_exibir["Estabelecimento"] = tabela_exibir["_estabelecimento"]
    tabela_exibir["Categoria"] = tabela_exibir["_categoria"]
    tabela_exibir["Tipo"] = tabela_exibir["_entrada"].str.title()
    tabela_exibir["Status"] = tabela_exibir["_status_exibicao"]
    tabela_exibir["Valor"] = tabela_exibir["_valor_num"].apply(formatar_brl)
    tabela_exibir["Observação"] = tabela_exibir["_detalhes"]
    tabela_exibir["Whatsapp"] = tabela_exibir["_whatsapp"]

    st.dataframe(
        tabela_exibir[
            ["Data", "Estabelecimento", "Categoria", "Tipo", "Status", "Valor", "Observação", "Whatsapp"]
        ],
        use_container_width=True,
        hide_index=True
    )

# =========================================================
# ATUALIZAR STATUS / VALOR
# =========================================================
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">✏️ Atualizar status e valor</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-text">Altere <b>Status</b> e <b>Valor</b> diretamente pelo dashboard. Receitas usam <b>Recebido / A Receber</b> e despesas usam <b>Pago / A Pagar</b>.</div>',
    unsafe_allow_html=True
)

busca = st.text_input(
    "Buscar lançamento para editar",
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
    df_update = df_update[df_update["_entrada_norm"] == filtro_entrada.lower()]

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
        entrada = row["_entrada_norm"] if str(row["_entrada_norm"]).strip() else "-"
        categoria = row["_categoria"] if str(row["_categoria"]).strip() else "-"
        status_atual = row["_status_exibicao"] if str(row["_status_exibicao"]).strip() else "Sem status"
        detalhes = row["_detalhes"] if str(row["_detalhes"]).strip() else "-"
        whatsapp = row["_whatsapp"] if str(row["_whatsapp"]).strip() else "-"
        mes_txt = row["_mes_raw"] if str(row["_mes_raw"]).strip() else "-"
        sheet_row = int(row["_sheet_row"])

        if entrada == "receita":
            btn_ok_label = "Recebido"
            btn_pendente_label = "A Receber"
            status_ok_gravar = "Recebido"
            status_pendente_gravar = "A Receber"
        else:
            btn_ok_label = "Pago"
            btn_pendente_label = "A Pagar"
            status_ok_gravar = "Pago"
            status_pendente_gravar = "A Pagar"

        st.markdown('<div class="update-card">', unsafe_allow_html=True)

        info1, info2, info3, b0, b1, b2 = st.columns([3.2, 1.45, 1.15, 1.35, 1.0, 1.1])

        with info1:
            st.markdown(f'<div class="item-title">{estabelecimento}</div>', unsafe_allow_html=True)
            st.markdown(
                f"""
                <div class="item-meta">
                    <b>Data:</b> {mes_txt}&nbsp;&nbsp;&nbsp;
                    <b>Tipo:</b> {entrada.title()}&nbsp;&nbsp;&nbsp;
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

            novo_valor = st.text_input(
                "Editar valor",
                value=formatar_valor_planilha(row["_valor_num"]),
                key=f"novo_valor_{sheet_row}",
                label_visibility="collapsed",
                placeholder="Ex.: 1574,00"
            )
            st.markdown('<div class="edit-hint">Digite o novo valor</div>', unsafe_allow_html=True)

        with info3:
            st.markdown('<div class="item-value-label">Status atual</div>', unsafe_allow_html=True)
            st.markdown(
                f'<span class="{status_class(status_atual)}">{status_atual}</span>',
                unsafe_allow_html=True
            )

        with b0:
            if st.button("Salvar alterações", key=f"salvar_valor_{sheet_row}", use_container_width=True):
                try:
                    atualizar_valor(sheet_row, novo_valor)
                    st.success(f"Alterações da linha {sheet_row} salvas com sucesso.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar valor: {e}")

        with b1:
            if st.button(btn_ok_label, key=f"ok_{sheet_row}", use_container_width=True):
                try:
                    atualizar_status(sheet_row, status_ok_gravar)
                    st.success(f"Status da linha {sheet_row} atualizado para {status_ok_gravar}.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar para {status_ok_gravar}: {e}")

        with b2:
            if st.button(btn_pendente_label, key=f"pend_{sheet_row}", use_container_width=True):
                try:
                    atualizar_status(sheet_row, status_pendente_gravar)
                    st.success(f"Status da linha {sheet_row} atualizado para {status_pendente_gravar}.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar para {status_pendente_gravar}: {e}")

        st.markdown('</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="small-note">Se algo não carregar, confira se a planilha continua compartilhada com o e-mail da service account como Editor.</div>',
    unsafe_allow_html=True
)
