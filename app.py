"""
BYM - Interface Web para Geração de Relatório Gerencial Mensal
Execute com: streamlit run app.py
"""

import json
import copy
import sys
import os
from datetime import datetime, date

import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from gerador_relatorio import GeradorRelatorio

# ─────────────────────────── CONFIGURAÇÃO DA PÁGINA ─────────────────────────

st.set_page_config(
    page_title="BYM - Relatório Gerencial",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    /* ── Google Font ── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* ── Variáveis de cor ── */
    :root {
        --azul-escuro:  #1a2f52;
        --azul-medio:   #2563a8;
        --azul-claro:   #dbeafe;
        --laranja:      #d97706;
        --laranja-dark: #b45309;
        --cinza-bg:     #F4F6F9;
        --cinza-borda:  #e2e8f0;
        --branco:       #FFFFFF;
        --texto-escuro: #0f172a;
        --texto-medio:  #475569;
    }

    /* ── Base ── */
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }

    /* ── Header principal ── */
    .main-title {
        background: linear-gradient(135deg, #1a2f52 0%, #2563a8 60%, #3b7dd8 100%);
        color: white;
        padding: 28px 36px;
        border-radius: 12px;
        margin-bottom: 24px;
        box-shadow: 0 4px 20px rgba(31,56,100,0.25);
        display: flex;
        align-items: center;
        gap: 16px;
    }
    .main-title h2 { margin: 0; font-size: 1.6rem; font-weight: 700; letter-spacing: -0.3px; }
    .main-title p  { margin: 6px 0 0 0; opacity: 0.82; font-size: 0.95rem; font-weight: 400; }
    .main-title .badge {
        background: rgba(255,255,255,0.15);
        border: 1px solid rgba(255,255,255,0.3);
        border-radius: 20px;
        padding: 4px 14px;
        font-size: 0.78rem;
        font-weight: 600;
        white-space: nowrap;
        margin-left: auto;
    }

    /* ── Cabeçalhos de seção ── */
    .section-header {
        display: flex;
        align-items: center;
        gap: 10px;
        background: linear-gradient(90deg, #1a2f52, #2563a8);
        color: white;
        padding: 10px 18px;
        border-radius: 8px;
        margin: 20px 0 12px 0;
        font-weight: 600;
        font-size: 0.92rem;
        letter-spacing: 0.2px;
        box-shadow: 0 2px 8px rgba(31,56,100,0.18);
    }

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background: #F0F4F8;
        padding: 6px;
        border-radius: 10px;
        margin-bottom: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 7px;
        padding: 8px 16px;
        font-size: 0.85rem;
        font-weight: 500;
        color: var(--texto-medio);
        border: none;
        background: transparent;
    }
    .stTabs [aria-selected="true"] {
        background: white !important;
        color: var(--azul-escuro) !important;
        font-weight: 700;
        box-shadow: 0 1px 6px rgba(0,0,0,0.12);
    }
    .stTabs [data-baseweb="tab-panel"] {
        padding-top: 16px;
    }

    /* ── Botões ── */
    .stButton > button {
        background: linear-gradient(135deg, #1a2f52, #2563a8);
        color: white;
        font-size: 0.95rem;
        font-weight: 600;
        border-radius: 8px;
        padding: 11px 28px;
        border: none;
        width: 100%;
        transition: all 0.2s ease;
        box-shadow: 0 3px 10px rgba(26,47,82,0.30);
        letter-spacing: 0.2px;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #142444, #1a4a8a);
        box-shadow: 0 4px 14px rgba(26,47,82,0.40);
        transform: translateY(-1px);
    }
    .stButton > button:active { transform: translateY(0); }

    /* ── Botão de ação principal (Gerar Relatório) ── */
    .btn-gerar .stButton > button {
        background: linear-gradient(135deg, #d97706, #b45309);
        box-shadow: 0 3px 10px rgba(217,119,6,0.35);
    }
    .btn-gerar .stButton > button:hover {
        background: linear-gradient(135deg, #b45309, #92400e);
        box-shadow: 0 4px 14px rgba(217,119,6,0.45);
    }

    /* ── Botão secundário (Salvar JSON) ── */
    .btn-salvar .stButton > button {
        background: linear-gradient(135deg, #2563a8, #1a2f52);
        box-shadow: 0 3px 10px rgba(37,99,168,0.3);
    }
    .btn-salvar .stButton > button:hover {
        background: linear-gradient(135deg, #1a2f52, #0f1e38);
        box-shadow: 0 4px 14px rgba(26,47,82,0.4);
    }

    /* ── Info box ── */
    .info-box {
        background: linear-gradient(135deg, #EEF4FB, #E3EDF8);
        padding: 14px 18px;
        border-radius: 10px;
        border-left: 5px solid #2E75B6;
        margin: 4px 0;
        font-size: 0.88rem;
        line-height: 1.7;
        box-shadow: 0 1px 4px rgba(31,56,100,0.08);
    }
    .info-box b { color: var(--azul-escuro); }

    /* ── Cards de métricas ── */
    [data-testid="metric-container"] {
        background: white;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 14px 18px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    [data-testid="metric-container"] label { color: var(--texto-medio) !important; font-size: 0.78rem !important; }
    [data-testid="metric-container"] [data-testid="stMetricValue"] { font-size: 1.4rem !important; font-weight: 700 !important; color: var(--azul-escuro) !important; }

    /* ── Inputs ── */
    .stTextInput input, .stNumberInput input, .stSelectbox select {
        border-radius: 7px;
        border: 1px solid var(--cinza-borda);
        font-size: 0.9rem;
    }
    .stTextInput input:focus, .stNumberInput input:focus {
        border-color: var(--azul-medio);
        box-shadow: 0 0 0 2px rgba(46,117,182,0.15);
    }

    /* ── Data editor ── */
    [data-testid="stDataFrame"], [data-testid="data-editor"] {
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid var(--cinza-borda);
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a2f52 0%, #142444 100%);
    }

    /* Textos diretos da sidebar (não widgets) */
    [data-testid="stSidebar"] > div > div > div p,
    [data-testid="stSidebar"] > div > div > div span,
    [data-testid="stSidebar"] > div > div > div label,
    [data-testid="stSidebar"] .stMarkdown p,
    [data-testid="stSidebar"] .stMarkdown span,
    [data-testid="stSidebar"] .stMarkdown div {
        color: rgba(255,255,255,0.80) !important;
        font-size: 0.88rem;
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 { color: white !important; font-weight: 700; }
    [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.12) !important; }

    /* File uploader: fundo semitransparente mas texto escuro interno preservado */
    [data-testid="stSidebar"] [data-testid="stFileUploader"] {
        background: rgba(255,255,255,0.92);
        border-radius: 10px;
        padding: 10px;
    }
    [data-testid="stSidebar"] [data-testid="stFileUploader"] * {
        color: #1a2f52 !important;
    }
    [data-testid="stSidebar"] [data-testid="stFileUploader"] small,
    [data-testid="stSidebar"] [data-testid="stFileUploader"] span {
        color: #4a6080 !important;
    }

    /* ── Sidebar step markers ── */
    .sidebar-step {
        display: flex;
        align-items: flex-start;
        gap: 10px;
        margin: 8px 0;
        font-size: 0.88rem;
        color: rgba(255,255,255,0.85) !important;
    }
    .sidebar-step .num {
        background: #d97706;
        color: white !important;
        border-radius: 50%;
        width: 22px; height: 22px;
        min-width: 22px;
        display: flex; align-items: center; justify-content: center;
        font-size: 0.72rem; font-weight: 700;
    }

    /* ── Divider ── */
    hr { border: none; border-top: 1px solid var(--cinza-borda); margin: 20px 0; }

    /* ── Spinner ── */
    .stSpinner > div { border-top-color: var(--laranja) !important; }

    /* ── Alertas ── */
    .stSuccess { border-radius: 8px; }
    .stError   { border-radius: 8px; }
    .stInfo    { border-radius: 8px; font-size: 0.88rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────── CARREGA DADOS BASE ──────────────────────────────

@st.cache_data(ttl=0)
def carregar_dados_base():
    caminho = os.path.join(os.path.dirname(__file__), "dados_relatorio.json")
    with open(caminho, encoding="utf-8") as f:
        return json.load(f)

# ─────────────────────────── HELPERS ─────────────────────────────────────────

def _str_para_date(valor):
    if not valor:
        return None
    if isinstance(valor, (datetime, date)):
        return valor
    try:
        return datetime.strptime(str(valor)[:10], "%Y-%m-%d").date()
    except Exception:
        return None

def _date_para_str(valor):
    if not valor:
        return None
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%Y-%m-%d")
    return str(valor)

def _nome_arquivo(dados):
    proj = dados["projeto"]["nome"].replace(" ", "_").replace("(", "").replace(")", "")
    periodo = f"{dados['relatorio']['mes']}-{dados['relatorio']['ano']}"
    return f"BYM_{proj}_{periodo}.xlsx"

# ─────────────────────────── SIDEBAR ─────────────────────────────────────────

with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding: 20px 0 10px 0;">
        <div style="font-size:3rem; margin-bottom:4px;">🏗️</div>
        <div style="font-size:1.6rem; font-weight:800; color:white; letter-spacing:-0.5px;">BYM</div>
        <div style="font-size:0.78rem; color:rgba(255,255,255,0.6); margin-top:2px; font-weight:400;">
            Relatório Gerencial Mensal
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    st.markdown('<div style="font-size:0.72rem; font-weight:700; color:rgba(255,255,255,0.5); letter-spacing:1.5px; text-transform:uppercase; margin-bottom:10px;">Como usar</div>', unsafe_allow_html=True)
    for num, texto in [("1", "Preencha os dados em cada aba"),
                        ("2", "Clique em <b>Gerar Relatório</b>"),
                        ("3", "Baixe o arquivo <b>.xlsx</b>")]:
        st.markdown(f"""
        <div class="sidebar-step">
            <div class="num">{num}</div>
            <div>{texto}</div>
        </div>""", unsafe_allow_html=True)

    st.divider()
    st.markdown('<div style="font-size:0.72rem; font-weight:700; color:rgba(255,255,255,0.5); letter-spacing:1.5px; text-transform:uppercase; margin-bottom:10px;">Carregar Dados</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Arquivo JSON existente",
        type="json",
        help="Carregue um arquivo JSON salvo anteriormente para editar",
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown('<div style="text-align:center; font-size:0.75rem; color:rgba(255,255,255,0.35);">BYM © 2026 — v1.0</div>', unsafe_allow_html=True)

# ─────────────────────────── ESTADO INICIAL ───────────────────────────────────

if "dados" not in st.session_state:
    st.session_state.dados = carregar_dados_base()

if uploaded:
    try:
        st.session_state.dados = json.load(uploaded)
        st.sidebar.success("JSON carregado com sucesso!")
    except Exception as e:
        st.sidebar.error(f"Erro ao carregar: {e}")

d = st.session_state.dados  # referência direta

# ─────────────────────────── TÍTULO ──────────────────────────────────────────

proj_nome = d['projeto']['nome'] or "—"
periodo_str = f"{d['relatorio']['mes']} / {d['relatorio']['ano']}" if d['relatorio']['mes'] else "—"
st.markdown(f"""
<div class="main-title">
    <div>
        <h2 style="margin:0">🏗️ BYM — Relatório Gerencial Mensal</h2>
        <p style="margin:6px 0 0 0; opacity:0.82">Preencha os dados nas abas abaixo e clique em <strong>Gerar Relatório</strong></p>
    </div>
    <div class="badge">📁 {proj_nome} &nbsp;·&nbsp; 📅 {periodo_str}</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────── ABAS ────────────────────────────────────────────

abas = st.tabs([
    "📋 Projeto",
    "📅 Prazos",
    "📊 Avanço Físico",
    "🚦 Farol de Metas",
    "🎯 Próximo Mês",
    "👷 Mão de Obra",
    "💰 Fluxo de Caixa",
    "💳 Tabela de Aporte",
    "🔨 Contratações",
    "📑 Legalização",
])

# ════════════════════════════════════════════════════════════════════════════
# ABA 1 — DADOS DO PROJETO
# ════════════════════════════════════════════════════════════════════════════
with abas[0]:
    st.markdown('<div class="section-header">1. Identificação do Relatório</div>',
                unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    MESES = ["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
             "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]
    with c1:
        idx_mes = MESES.index(d["relatorio"]["mes"]) if d["relatorio"]["mes"] in MESES else 0
        mes = st.selectbox("Mês de referência", MESES, index=idx_mes)
        d["relatorio"]["mes"] = mes
    with c2:
        ano = st.text_input("Ano (2 dígitos)", value=d["relatorio"]["ano"], max_chars=2)
        d["relatorio"]["ano"] = ano
    with c3:
        dr = _str_para_date(d["relatorio"]["data_relatorio"])
        data_rel = st.date_input("Data do Relatório", value=dr or date.today())
        d["relatorio"]["data_relatorio"] = _date_para_str(data_rel)

    st.markdown('<div class="section-header">2. Dados do Empreendimento</div>',
                unsafe_allow_html=True)
    p = d["projeto"]
    c1, c2 = st.columns(2)
    with c1:
        p["nome"]          = st.text_input("Nome do Empreendimento", p["nome"])
        p["incorporadora"] = st.text_input("Incorporadora", p["incorporadora"])
        p["construtora"]   = st.text_input("Construtora", p["construtora"])
        p["gerenciadora"]  = st.text_input("Gerenciadora", p["gerenciadora"])
        p["padrao"]        = st.text_input("Padrão", p["padrao"])
        p["endereco"]      = st.text_input("Endereço", p["endereco"])
    with c2:
        p["torres"]        = st.number_input("Torres", min_value=0, value=int(p["torres"]))
        p["num_fases"]     = st.number_input("Nº de Fases", min_value=0, value=int(p["num_fases"]))
        p["unidades"]      = st.number_input("Unidades", min_value=0, value=int(p["unidades"]))
        p["subsolos_vagas"]= st.text_input("Subsolos e Vagas", p["subsolos_vagas"])
        p["num_pavimentos"]= st.text_input("Nº de Pavimentos", p["num_pavimentos"])

    st.markdown('<div class="section-header">3. Áreas e Técnico</div>',
                unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        p["area_terreno"]  = st.number_input("Área do Terreno (m²)", value=float(p["area_terreno"]), format="%.2f")
        p["area_terreno_referencia"] = st.text_input("Referência área terreno", p["area_terreno_referencia"])
    with c2:
        p["area_construida"] = st.number_input("Área Construída (m²)", value=float(p["area_construida"]), format="%.2f")
        p["area_construida_referencia"] = st.text_input("Referência área construída", p["area_construida_referencia"])
    with c3:
        p["area_privativa"] = st.number_input("Área Privativa (m²)", value=float(p["area_privativa"]), format="%.2f")
        p["area_privativa_referencia"] = st.text_input("Referência área privativa", p["area_privativa_referencia"])

    c1, c2, c3 = st.columns(3)
    with c1:
        p["tipo_contencao"] = st.text_input("Tipo de Contenção", p["tipo_contencao"])
    with c2:
        p["tipo_fundacao"]  = st.text_input("Tipo de Fundação", p["tipo_fundacao"])
    with c3:
        p["tipo_estrutura"] = st.text_input("Tipo de Estrutura", p["tipo_estrutura"])

# ════════════════════════════════════════════════════════════════════════════
# ABA 2 — CONTROLE DE PRAZO
# ════════════════════════════════════════════════════════════════════════════
with abas[1]:
    pr = d["controle_prazo"]

    st.markdown('<div class="section-header">2.1 Prazos Contratuais x Tendência</div>',
                unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.caption("**Contratual**")
        v = _str_para_date(pr["data_inicio_contratual"])
        pr["data_inicio_contratual"] = _date_para_str(
            st.date_input("Início Contratual", value=v or date.today(), key="ic"))
        v2 = _str_para_date(pr["data_termino_contratual"])
        pr["data_termino_contratual"] = _date_para_str(
            st.date_input("Término Contratual", value=v2 or date.today(), key="tc"))
        pr["qtde_meses_contratual"] = st.number_input(
            "Qtde Meses Contratual", value=float(pr["qtde_meses_contratual"]),
            format="%.1f")
    with c2:
        st.caption("**Tendência**")
        v3 = _str_para_date(pr["data_inicio_tendencia"])
        pr["data_inicio_tendencia"] = _date_para_str(
            st.date_input("Início Tendência", value=v3 or date.today(), key="it"))
        v4 = _str_para_date(pr["data_termino_tendencia"])
        pr["data_termino_tendencia"] = _date_para_str(
            st.date_input("Término Tendência", value=v4 or date.today(), key="tt"))
        pr["qtde_meses_tendencia"] = st.number_input(
            "Qtde Meses Tendência", value=float(pr["qtde_meses_tendencia"]),
            format="%.1f")
    with c3:
        desvio = pr["qtde_meses_tendencia"] - pr["qtde_meses_contratual"]
        pr["desvio_dias"] = st.number_input("Desvio (dias)", value=int(pr.get("desvio_dias", 0)))
        st.metric("Desvio em Meses", f"{desvio:+.1f}")

    st.markdown('<div class="section-header">2.3 Datas Marco Contratual</div>',
                unsafe_allow_html=True)
    marcos = pr.get("datas_marco_contratual", [])
    df_marcos = pd.DataFrame([{
        "Atividade":        m["atividade"],
        "Baseline Meses":   m["baseline_meses"],
        "Baseline Data":    _str_para_date(m["baseline_data"]),
        "Previsto Meses":   m["previsto_meses"],
        "Previsto Data":    _str_para_date(m["previsto_data"]),
        "Desvio (dias)":    m["desvio_dias"],
        "Farol (1=V 2=A 3=R)": m.get("farol", 3),
    } for m in marcos])
    edited_marcos = st.data_editor(
        df_marcos, num_rows="dynamic", use_container_width=True, key="marcos",
        column_config={
            "Baseline Data":  st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Previsto Data":  st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Farol (1=V 2=A 3=R)": st.column_config.SelectboxColumn(
                options=[1, 2, 3]),
        }
    )
    pr["datas_marco_contratual"] = [
        {"atividade": row["Atividade"],
         "baseline_meses": row["Baseline Meses"],
         "baseline_data": _date_para_str(row["Baseline Data"]),
         "previsto_meses": row["Previsto Meses"],
         "previsto_data": _date_para_str(row["Previsto Data"]),
         "desvio_dias": row["Desvio (dias)"],
         "farol": row["Farol (1=V 2=A 3=R)"]}
        for _, row in edited_marcos.iterrows()
        if pd.notna(row.get("Atividade"))
    ]

    st.markdown('<div class="section-header">5. Datas Marco Protótipo</div>',
                unsafe_allow_html=True)
    df_proto = pd.DataFrame([{
        "Atividade":       m["atividade"],
        "Local":           m["local"],
        "Data Contratual": _str_para_date(m["data_contratual"]),
        "Data Prev. Eng":  _str_para_date(m["data_prevista_eng"]),
        "Data Prev. Dir":  _str_para_date(m["data_prevista_dir"]),
        "Desvio (dias)":   m["desvio_dias"],
        "Farol":           m.get("farol"),
    } for m in d.get("datas_marco_prototipo", [])])
    edited_proto = st.data_editor(
        df_proto, num_rows="dynamic", use_container_width=True, key="proto",
        column_config={
            "Data Contratual": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Data Prev. Eng":  st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Data Prev. Dir":  st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Farol": st.column_config.SelectboxColumn(options=[None, 1, 2, 3]),
        }
    )
    d["datas_marco_prototipo"] = [
        {"atividade": row["Atividade"], "local": row["Local"],
         "data_contratual": _date_para_str(row["Data Contratual"]),
         "data_prevista_eng": _date_para_str(row["Data Prev. Eng"]),
         "data_prevista_dir": _date_para_str(row["Data Prev. Dir"]),
         "desvio_dias": row["Desvio (dias)"],
         "farol": row["Farol"]}
        for _, row in edited_proto.iterrows()
        if pd.notna(row.get("Atividade"))
    ]

# ════════════════════════════════════════════════════════════════════════════
# ABA 3 — AVANÇO FÍSICO
# ════════════════════════════════════════════════════════════════════════════
with abas[2]:
    av = d["avanco_fisico"]
    st.markdown('<div class="section-header">2.2 Avanço Físico no Período</div>',
                unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        av["qtde_servicos_previstos_mes"] = st.number_input(
            "Qtde Serviços Previstos", min_value=0,
            value=int(av["qtde_servicos_previstos_mes"]))
        av["qtde_servicos_executados_mes"] = st.number_input(
            "Qtde Serviços Executados", min_value=0,
            value=int(av["qtde_servicos_executados_mes"]))
        if av["qtde_servicos_previstos_mes"] > 0:
            av["percentual_atingido"] = av["qtde_servicos_executados_mes"] / av["qtde_servicos_previstos_mes"]
        st.metric("% Atingido", f"{av['percentual_atingido']:.1%}")
    with c2:
        st.caption("**Mês**")
        av["previsto_curva_contratual_mes"] = st.number_input(
            "Previsto Contratual Mês (%)", value=float(av.get("previsto_curva_contratual_mes", 0)) * 100,
            format="%.2f") / 100
        av["realizado_medido_mes"] = st.number_input(
            "Realizado Medido Mês (%)", value=float(av.get("realizado_medido_mes", 0)) * 100,
            format="%.2f") / 100
        av["meta_bym_mes"] = st.number_input(
            "Meta BYM Mês (%)", value=float(av.get("meta_bym_mes", 0)) * 100,
            format="%.2f") / 100
    with c3:
        st.caption("**Acumulado**")
        av["previsto_curva_contratual_acum"] = st.number_input(
            "Previsto Contratual Acum (%)", value=float(av.get("previsto_curva_contratual_acum", 0)) * 100,
            format="%.2f") / 100
        av["realizado_medido_acum"] = st.number_input(
            "Realizado Medido Acum (%)", value=float(av.get("realizado_medido_acum", 0)) * 100,
            format="%.2f") / 100
        av["meta_bym_acum"] = st.number_input(
            "Meta BYM Acum (%)", value=float(av.get("meta_bym_acum", 0)) * 100,
            format="%.2f") / 100

    # Calcula desvios e IPF automaticamente
    av["desvio_mes"] = av["realizado_medido_mes"] - av["previsto_curva_contratual_mes"]
    av["desvio_acumulado"] = av["realizado_medido_acum"] - av["previsto_curva_contratual_acum"]
    if av["previsto_curva_contratual_acum"] > 0:
        av["ipf_contratual"] = av["realizado_medido_acum"] / av["previsto_curva_contratual_acum"]
    if av["meta_bym_acum"] > 0:
        av["ipf_meta_bym"] = av["realizado_medido_acum"] / av["meta_bym_acum"]

    st.divider()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Desvio Mês", f"{av['desvio_mes']:.2%}",
              delta=f"{av['desvio_mes']:.2%}", delta_color="normal")
    c2.metric("Desvio Acum", f"{av['desvio_acumulado']:.2%}",
              delta=f"{av['desvio_acumulado']:.2%}", delta_color="normal")
    c3.metric("IPF Contratual", f"{av.get('ipf_contratual', 0):.3f}")
    c4.metric("IPF Meta BYM", f"{av.get('ipf_meta_bym', 0):.3f}")

# ════════════════════════════════════════════════════════════════════════════
# ABA 4 — FAROL DE METAS
# ════════════════════════════════════════════════════════════════════════════
with abas[3]:
    st.markdown('<div class="section-header">8. Farol da Lista de Metas do Período</div>',
                unsafe_allow_html=True)
    st.info("Farol: 1 = Vermelho (não executado) | 2 = Amarelo (parcial) | 3 = Verde (executado)")

    rows = []
    for grupo in d["farol_metas"]:
        for item in grupo["itens"]:
            rows.append({
                "Torre":       grupo.get("torre", "TORRE"),
                "Categoria":   grupo["categoria"],
                "Nome":        item["nome"],
                "% Concluído": item["percentual"],
                "Duração":     item["duracao"],
                "Início":      _str_para_date(item["inicio"]),
                "Término":     _str_para_date(item["termino"]),
                "Término Real": _str_para_date(item.get("termino_real")),
                "Farol":       item.get("farol", 3),
                "Observações": item.get("observacoes", ""),
            })

    df_farol = pd.DataFrame(rows)
    edited_farol = st.data_editor(
        df_farol, num_rows="dynamic", use_container_width=True, key="farol",
        column_config={
            "Início":       st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Término":      st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Término Real": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Farol":        st.column_config.SelectboxColumn(options=[1, 2, 3]),
            "% Concluído":  st.column_config.NumberColumn(min_value=0, max_value=100),
        }
    )

    # Reconstrói estrutura agrupada
    grupos_dict = {}
    for _, row in edited_farol.iterrows():
        if not pd.notna(row.get("Nome")):
            continue
        chave = (str(row.get("Torre", "TORRE")), str(row.get("Categoria", "")))
        if chave not in grupos_dict:
            grupos_dict[chave] = []
        grupos_dict[chave].append({
            "nome":        str(row["Nome"]),
            "percentual":  int(row.get("% Concluído", 0) or 0),
            "duracao":     str(row.get("Duração", "")),
            "inicio":      _date_para_str(row.get("Início")),
            "termino":     _date_para_str(row.get("Término")),
            "termino_real": _date_para_str(row.get("Término Real")),
            "farol":       int(row.get("Farol", 3) or 3),
            "observacoes": str(row.get("Observações", "") or ""),
        })
    d["farol_metas"] = [
        {"torre": k[0], "categoria": k[1], "itens": v}
        for k, v in grupos_dict.items()
    ]

# ════════════════════════════════════════════════════════════════════════════
# ABA 5 — METAS PRÓXIMO MÊS
# ════════════════════════════════════════════════════════════════════════════
with abas[4]:
    st.markdown('<div class="section-header">9. Metas Para o Próximo Período</div>',
                unsafe_allow_html=True)

    rows_pm = []
    for grupo in d["metas_proximo_mes"]:
        for item in grupo["itens"]:
            rows_pm.append({
                "Torre":            grupo.get("torre", "TORRE"),
                "Categoria":        grupo["categoria"],
                "Nome":             item["nome"],
                "% Concluído":      item["percentual"],
                "Duração":          item["duracao"],
                "Início":           _str_para_date(item["inicio"]),
                "Término":          _str_para_date(item["termino"]),
                "Baseline Início":  _str_para_date(item.get("baseline_inicio")),
                "Baseline Término": _str_para_date(item.get("baseline_termino")),
            })

    df_pm = pd.DataFrame(rows_pm)
    edited_pm = st.data_editor(
        df_pm, num_rows="dynamic", use_container_width=True, key="metas_pm",
        column_config={
            "Início":           st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Término":          st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Baseline Início":  st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Baseline Término": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "% Concluído":      st.column_config.NumberColumn(min_value=0, max_value=100),
        }
    )

    grupos_pm = {}
    for _, row in edited_pm.iterrows():
        if not pd.notna(row.get("Nome")):
            continue
        chave = (str(row.get("Torre", "TORRE")), str(row.get("Categoria", "")))
        if chave not in grupos_pm:
            grupos_pm[chave] = []
        grupos_pm[chave].append({
            "nome":             str(row["Nome"]),
            "percentual":       int(row.get("% Concluído", 0) or 0),
            "duracao":          str(row.get("Duração", "")),
            "inicio":           _date_para_str(row.get("Início")),
            "termino":          _date_para_str(row.get("Término")),
            "baseline_inicio":  _date_para_str(row.get("Baseline Início")),
            "baseline_termino": _date_para_str(row.get("Baseline Término")),
        })
    d["metas_proximo_mes"] = [
        {"torre": k[0], "categoria": k[1], "itens": v}
        for k, v in grupos_pm.items()
    ]

# ════════════════════════════════════════════════════════════════════════════
# ABA 6 — MÃO DE OBRA
# ════════════════════════════════════════════════════════════════════════════
with abas[5]:
    hist = d["histograma_mao_obra"]
    st.markdown('<div class="section-header">6. Histograma de Mão de Obra</div>',
                unsafe_allow_html=True)

    v_ref = _str_para_date(hist["mes_referencia"])
    ref = st.date_input("Mês de referência", value=v_ref or date.today(), key="hist_ref")
    hist["mes_referencia"] = _date_para_str(ref)

    df_hist = pd.DataFrame([{
        "Empresa":    e["nome"],
        "Serviço":    e["servico"],
        "Mês -2":     e.get("mes_anterior2"),
        "Mês -1":     e.get("mes_anterior1"),
        "Mês Atual":  e.get("mes_atual"),
        "Semana 1":   e.get("semana1"),
        "Semana 2":   e.get("semana2"),
        "Semana 3":   e.get("semana3"),
        "Semana 4":   e.get("semana4"),
        "Total":      e.get("total"),
    } for e in hist["empresas"]])
    edited_hist = st.data_editor(
        df_hist, num_rows="dynamic", use_container_width=True, key="hist")
    hist["empresas"] = [
        {"nome": row["Empresa"], "servico": row["Serviço"],
         "mes_anterior2": row.get("Mês -2"), "mes_anterior1": row.get("Mês -1"),
         "mes_atual": row.get("Mês Atual"),
         "semana1": row.get("Semana 1"), "semana2": row.get("Semana 2"),
         "semana3": row.get("Semana 3"), "semana4": row.get("Semana 4"),
         "total": row.get("Total")}
        for _, row in edited_hist.iterrows()
        if pd.notna(row.get("Empresa"))
    ]

# ════════════════════════════════════════════════════════════════════════════
# ABA 7 — FLUXO DE CAIXA
# ════════════════════════════════════════════════════════════════════════════
with abas[6]:
    fc = d.get("fluxo_caixa", {})
    st.markdown('<div class="section-header">3.2 Fluxo de Caixa</div>',
                unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        fc["valor_total_contrato"] = st.number_input(
            "Valor Total do Contrato (R$)",
            value=float(fc.get("valor_total_contrato", 0)), format="%.2f")
    with c2:
        fc["incc_base"] = st.number_input(
            "INCC Base", value=float(fc.get("incc_base", 0)), format="%.3f")
    with c3:
        fc["incc_base_referencia"] = st.text_input(
            "Referência INCC", value=fc.get("incc_base_referencia", ""))

    df_fc = pd.DataFrame([{
        "#":                 m["mes"],
        "Data":              _str_para_date(m["data"]),
        "Previsto Valor":    m.get("previsto_valor"),
        "Previsto INCC":     m.get("previsto_incc"),
        "Previsto %":        m.get("previsto_pct"),
        "Realizado Valor":   m.get("realizado_valor"),
        "Realizado INCC":    m.get("realizado_incc"),
        "Realizado %":       m.get("realizado_pct"),
    } for m in fc.get("meses", [])])
    edited_fc = st.data_editor(
        df_fc, num_rows="dynamic", use_container_width=True, key="fc",
        column_config={
            "Data": st.column_config.DateColumn(format="MM/YYYY"),
            "Previsto %":  st.column_config.NumberColumn(format="%.4f"),
            "Realizado %": st.column_config.NumberColumn(format="%.4f"),
        }
    )
    novos_meses_fc = []
    for i, (_, row) in enumerate(edited_fc.iterrows()):
        if not pd.notna(row.get("Data")):
            continue
        rv = row.get("Realizado Valor")
        rp = row.get("Realizado %")
        pv = row.get("Previsto Valor") or 0
        pp = row.get("Previsto %") or 0
        desvio = (rp - pp) if (rv is not None and pd.notna(rv) and rp is not None and pd.notna(rp)) else None
        novos_meses_fc.append({
            "mes": int(row.get("#") or i + 1),
            "data": _date_para_str(row["Data"]),
            "previsto_valor": pv,
            "previsto_incc": row.get("Previsto INCC"),
            "previsto_pct": pp,
            "realizado_valor": rv if pd.notna(rv) else None,
            "realizado_incc": row.get("Realizado INCC") if pd.notna(row.get("Realizado INCC")) else None,
            "realizado_pct": rp if pd.notna(rp) else None,
            "desvio_pct": desvio,
        })
    fc["meses"] = novos_meses_fc
    d["fluxo_caixa"] = fc

# ════════════════════════════════════════════════════════════════════════════
# ABA 8 — TABELA DE APORTE
# ════════════════════════════════════════════════════════════════════════════
with abas[7]:
    st.markdown('<div class="section-header">13. Tabela de Aportes — Previsto x Realizado</div>',
                unsafe_allow_html=True)

    df_ap = pd.DataFrame([{
        "#":             a["mes"],
        "Data":          _str_para_date(a["data"]),
        "Prev % Mês":    a["previsto_pct_mes"],
        "Prev % Acum":   a["previsto_pct_acum"],
        "Data Aporte":   _str_para_date(a["data_aporte"]),
        "Real Valor":    a["realizado_valor"],
        "Real INCC":     a["realizado_incc"],
        "Real %":        a["realizado_pct"],
        "Acum Valor":    a["acum_valor"],
        "Acum INCC":     a["acum_incc"],
        "Acum %":        a["acum_pct"],
        "INCC Utilizado":a["incc_utilizado"],
        "Desvio %":      a["desvio_pct"],
    } for a in d["tabela_aporte"]])
    edited_ap = st.data_editor(
        df_ap, num_rows="dynamic", use_container_width=True, key="aporte",
        column_config={
            "Data":       st.column_config.DateColumn(format="MM/YYYY"),
            "Data Aporte":st.column_config.DateColumn(format="DD/MM/YYYY"),
        }
    )
    d["tabela_aporte"] = [
        {"mes": int(row.get("#") or 0),
         "data": _date_para_str(row["Data"]),
         "previsto_pct_mes": float(row.get("Prev % Mês") or 0),
         "previsto_pct_acum": float(row.get("Prev % Acum") or 0),
         "data_aporte": _date_para_str(row.get("Data Aporte")),
         "realizado_valor": float(row.get("Real Valor") or 0),
         "realizado_incc": float(row.get("Real INCC") or 0),
         "realizado_pct": float(row.get("Real %") or 0),
         "acum_valor": float(row.get("Acum Valor") or 0),
         "acum_incc": float(row.get("Acum INCC") or 0),
         "acum_pct": float(row.get("Acum %") or 0),
         "incc_utilizado": float(row.get("INCC Utilizado") or 0),
         "desvio_pct": float(row.get("Desvio %") or 0)}
        for _, row in edited_ap.iterrows()
        if pd.notna(row.get("Data"))
    ]

# ════════════════════════════════════════════════════════════════════════════
# ABA 9 — CONTRATAÇÕES
# ════════════════════════════════════════════════════════════════════════════
with abas[8]:
    st.markdown('<div class="section-header">14. Gerenciamento de Contratações</div>',
                unsafe_allow_html=True)

    df_con = pd.DataFrame([{
        "ID":              c["id"],
        "Serviço":         c["servico"],
        "Data Prevista":   _str_para_date(c.get("data_prevista")),
        "Data Contratação":_str_para_date(c.get("data_contratacao")),
        "Farol":           c.get("farol", 3),
        "Fornecedor":      c.get("fornecedor", ""),
    } for c in d.get("gerenciamento_contratacoes", [])])
    edited_con = st.data_editor(
        df_con, num_rows="dynamic", use_container_width=True, key="contrat",
        column_config={
            "Data Prevista":    st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Data Contratação": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Farol":            st.column_config.SelectboxColumn(options=[1, 2, 3]),
        }
    )
    d["gerenciamento_contratacoes"] = [
        {"id": int(row.get("ID") or i + 1),
         "servico": str(row["Serviço"]),
         "data_prevista": _date_para_str(row.get("Data Prevista")),
         "data_contratacao": _date_para_str(row.get("Data Contratação")),
         "farol": int(row.get("Farol") or 3),
         "fornecedor": str(row.get("Fornecedor") or "")}
        for i, (_, row) in enumerate(edited_con.iterrows())
        if pd.notna(row.get("Serviço"))
    ]

    st.markdown('<div class="section-header">15. Controle de Mapas de Contratações</div>',
                unsafe_allow_html=True)

    df_map = pd.DataFrame([{
        "ID":              m["id"],
        "Mês":             m.get("mes", ""),
        "Serviço":         m["servico"],
        "Fornecedor":      m["fornecedor"],
        "Valor Contratado":m["valor_contratado"],
        "Orçamento Atualizado": m["orcamento_atualizado"],
        "Desvio":          m.get("desvio", 0),
        "Índice":          m.get("indice", 1.0),
    } for m in d.get("controle_mapas_contratacoes", [])])
    edited_map = st.data_editor(
        df_map, num_rows="dynamic", use_container_width=True, key="mapas",
        column_config={
            "Valor Contratado":      st.column_config.NumberColumn(format="R$ %.2f"),
            "Orçamento Atualizado":  st.column_config.NumberColumn(format="R$ %.2f"),
        }
    )
    d["controle_mapas_contratacoes"] = [
        {"id": int(row.get("ID") or i + 1),
         "mes": str(row.get("Mês") or ""),
         "servico": str(row["Serviço"]),
         "fornecedor": str(row.get("Fornecedor") or ""),
         "valor_contratado": float(row.get("Valor Contratado") or 0),
         "orcamento_atualizado": float(row.get("Orçamento Atualizado") or 0),
         "desvio": float(row.get("Desvio") or 0),
         "indice": float(row.get("Índice") or 1.0)}
        for i, (_, row) in enumerate(edited_map.iterrows())
        if pd.notna(row.get("Serviço"))
    ]

# ════════════════════════════════════════════════════════════════════════════
# ABA 10 — LEGALIZAÇÃO
# ════════════════════════════════════════════════════════════════════════════
with abas[9]:
    leg = d["legalizacao"]
    st.markdown('<div class="section-header">10. Legalização e Habite-se</div>',
                unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        v = _str_para_date(leg["inicio_obra"])
        leg["inicio_obra"] = _date_para_str(
            st.date_input("Início da Obra", value=v or date.today(), key="leg_ini"))
    with c2:
        v2 = _str_para_date(leg["mes_habitese"])
        leg["mes_habitese"] = _date_para_str(
            st.date_input("Mês Habite-se", value=v2 or date.today(), key="leg_hab"))
    with c3:
        v3 = _str_para_date(leg["mes_termino"])
        leg["mes_termino"] = _date_para_str(
            st.date_input("Mês Término", value=v3 or date.today(), key="leg_ter"))

    st.markdown("**Documentos Iniciais**")
    df_leg = pd.DataFrame([{
        "Documento":       doc["documento"],
        "Disponibilizado": _str_para_date(doc["disponibilizado"]),
        "Validade":        doc.get("validade", "N/A"),
        "Prazo":           doc.get("prazo", "I"),
        "Status":          doc["status"],
    } for doc in leg["documentos_iniciais"]])
    edited_leg = st.data_editor(
        df_leg, num_rows="dynamic", use_container_width=True, key="leg",
        column_config={
            "Disponibilizado": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Status": st.column_config.SelectboxColumn(options=["OK", "Pendente", "Vencido"]),
        }
    )
    leg["documentos_iniciais"] = [
        {"documento": row["Documento"],
         "disponibilizado": _date_para_str(row.get("Disponibilizado")),
         "validade": str(row.get("Validade") or "N/A"),
         "prazo": str(row.get("Prazo") or "I"),
         "status": str(row.get("Status") or "Pendente")}
        for _, row in edited_leg.iterrows()
        if pd.notna(row.get("Documento"))
    ]

# ════════════════════════════════════════════════════════════════════════════
# BOTÃO GERAR RELATÓRIO
# ════════════════════════════════════════════════════════════════════════════
st.markdown("<hr>", unsafe_allow_html=True)
col_btn, col_save, col_info = st.columns([2, 1, 3])

with col_save:
    st.markdown('<div class="btn-salvar">', unsafe_allow_html=True)
    salvar_clicado = st.button("💾 Salvar JSON", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    if salvar_clicado:
        json_str = json.dumps(d, ensure_ascii=False, indent=2, default=str)
        st.download_button(
            label="⬇️ Baixar dados_relatorio.json",
            data=json_str.encode("utf-8"),
            file_name="dados_relatorio.json",
            mime="application/json",
            use_container_width=True,
        )

with col_btn:
    st.markdown('<div class="btn-gerar">', unsafe_allow_html=True)
    gerar = st.button("📄 Gerar Relatório Excel", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with col_info:
    nome_proj   = d['projeto']['nome']      or '<span style="opacity:.5">não informado</span>'
    periodo_inf = f"{d['relatorio']['mes']} / {d['relatorio']['ano']}" if d['relatorio']['mes'] else '<span style="opacity:.5">—</span>'
    gerenc      = d['projeto']['gerenciadora'] or '<span style="opacity:.5">não informado</span>'
    constrt     = d['projeto']['construtora']  or '<span style="opacity:.5">não informado</span>'
    st.markdown(f"""
    <div class="info-box">
        <b>Projeto:</b> {nome_proj}<br>
        <b>Período:</b> {periodo_inf}<br>
        <b>Construtora:</b> {constrt}<br>
        <b>Gerenciadora:</b> {gerenc}
    </div>""", unsafe_allow_html=True)

if gerar:
    with st.spinner("⏳ Gerando relatório, aguarde..."):
        try:
            log_msgs = []
            gerador = GeradorRelatorio(copy.deepcopy(d))
            xlsx_bytes = gerador.gerar_bytes(log_fn=log_msgs.append)

            nome_arquivo = _nome_arquivo(d)
            tamanho_kb = len(xlsx_bytes) // 1024

            st.success(f"✅ Relatório gerado com sucesso! Tamanho: **{tamanho_kb} KB**")
            st.download_button(
                label=f"⬇️ Baixar {nome_arquivo}",
                data=xlsx_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"❌ Erro ao gerar relatório: {e}")
            st.exception(e)
