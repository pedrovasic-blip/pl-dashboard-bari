import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(
    page_title="Dashboard P&L 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_PADRAO = "2026_03_PL_com_BASE_DASH_v2.xlsx"
ABA_BASE = "BASE_DASH"

CUSTOM_CSS = """
<style>
    .stApp {
        background: #080f1f;
        color: #e5ecff;
    }
    [data-testid="stSidebar"] {
        background: #0b1224;
        border-right: 1px solid #1e2a44;
    }
    [data-testid="stHeader"] {
        background: rgba(8, 15, 31, 0.95);
    }
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    .title {
        font-size: 2.2rem;
        font-weight: 800;
        margin-bottom: 0.15rem;
    }
    .subtitle {
        color: #8ea0c9;
        font-size: 0.95rem;
        margin-bottom: 1.4rem;
    }
    .kpi-card {
        background: #111a2e;
        border: 1px solid #1e2a44;
        border-radius: 14px;
        padding: 18px 18px;
        min-height: 112px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.18);
    }
    .kpi-label {
        color: #8ea0c9;
        font-size: 0.78rem;
        margin-bottom: 8px;
    }
    .kpi-value {
        color: #ffffff;
        font-size: 1.55rem;
        font-weight: 800;
        line-height: 1.2;
    }
    .kpi-help {
        color: #5f719a;
        font-size: 0.72rem;
        margin-top: 7px;
    }
    div[data-testid="stMetricValue"] {
        color: #ffffff;
    }
    div[data-testid="stMetricLabel"] {
        color: #8ea0c9;
    }
    .section-title {
        font-size: 1.25rem;
        font-weight: 700;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
    }
    .info-box {
        background: #111a2e;
        border: 1px solid #1e2a44;
        border-radius: 14px;
        padding: 14px 16px;
        color: #8ea0c9;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        border-bottom: 1px solid #1e2a44;
    }
    .stTabs [data-baseweb="tab"] {
        color: #8ea0c9;
        background: transparent;
        border-radius: 10px 10px 0 0;
    }
    .stTabs [aria-selected="true"] {
        color: #ffffff;
        border-bottom: 2px solid #1f77ff;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def carregar_base(arquivo):
    df = pd.read_excel(arquivo, sheet_name=ABA_BASE)

    # Remove colunas vazias/auxiliares geradas pelo Excel.
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df = df.drop(columns=[c for c in df.columns if c.lower().startswith("observ")], errors="ignore")

    # Padronizações básicas.
    for col in ["Visao", "Linha_PnL", "Produto", "Metrica", "Aba_Origem", "Celula_Origem"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    if "Valor" in df.columns:
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)

    # Tenta converter datas, inclusive serial Excel.
    for col in ["Data_Inicio", "Data_Fim", "Data_Base"]:
        if col in df.columns:
            s = df[col]
            if pd.api.types.is_numeric_dtype(s):
                df[col] = pd.to_datetime(s, unit="D", origin="1899-12-30", errors="coerce")
            else:
                df[col] = pd.to_datetime(s, errors="coerce", dayfirst=True)

    if "Periodo" in df.columns:
        df["Periodo"] = df["Periodo"].astype(str).replace("nan", "")

    return df


def fmt_moeda(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0
    sinal = "-" if valor < 0 else ""
    valor = abs(valor)

    if valor >= 1_000_000_000:
        return f"{sinal}R$ {valor/1_000_000_000:,.2f} bi".replace(",", "X").replace(".", ",").replace("X", ".")
    if valor >= 1_000_000:
        return f"{sinal}R$ {valor/1_000_000:,.2f} mi".replace(",", "X").replace(".", ",").replace("X", ".")
    if valor >= 1_000:
        return f"{sinal}R$ {valor/1_000:,.1f} mil".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sinal}R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_num(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0
    return f"{valor:,.0f}".replace(",", ".")


def card(label, value, help_text=""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-help">{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def escolher_linha(df, termos_preferidos):
    linhas = sorted(df["Linha_PnL"].dropna().unique()) if "Linha_PnL" in df.columns else []
    linhas_lower = {linha.lower(): linha for linha in linhas}
    for termo in termos_preferidos:
        for linha_lower, linha_original in linhas_lower.items():
            if termo.lower() in linha_lower:
                return linha_original
    return linhas[0] if linhas else None


# ==========================
# Entrada de dados
# ==========================

st.sidebar.title("Filtros")

arquivo_local = Path(ARQUIVO_PADRAO)
upload = st.sidebar.file_uploader("Atualizar base manualmente", type=["xlsx"])

try:
    if upload is not None:
        df = carregar_base(upload)
    elif arquivo_local.exists():
        df = carregar_base(arquivo_local)
    else:
        st.error(
            f"Arquivo '{ARQUIVO_PADRAO}' não encontrado. "
            "Suba o Excel no mesmo repositório do app ou use o upload lateral."
        )
        st.stop()
except Exception as e:
    st.error(f"Erro ao carregar a base: {e}")
    st.stop()

colunas_necessarias = {"Visao", "Linha_PnL", "Produto", "Metrica", "Valor"}
faltantes = colunas_necessarias - set(df.columns)
if faltantes:
    st.error(f"A aba {ABA_BASE} não contém as colunas esperadas: {', '.join(sorted(faltantes))}")
    st.stop()

visoes = sorted(df["Visao"].dropna().unique())
produtos = sorted(df["Produto"].dropna().unique())
metricas = sorted(df["Metrica"].dropna().unique())
periodos = sorted(df["Periodo"].dropna().unique()) if "Periodo" in df.columns else []

visao_sel = st.sidebar.selectbox("Visão", visoes, index=0 if "Mensal" not in visoes else visoes.index("Mensal"))
produto_sel = st.sidebar.multiselect("Produto", produtos, default=produtos)
metrica_sel = st.sidebar.multiselect("Métrica", metricas, default=[m for m in metricas if m in ["Realizado", "Orçado"]] or metricas)

if periodos:
    periodo_sel = st.sidebar.multiselect("Período", periodos, default=periodos)
else:
    periodo_sel = []

df_f = df[
    (df["Visao"] == visao_sel)
    & (df["Produto"].isin(produto_sel))
    & (df["Metrica"].isin(metrica_sel))
].copy()

if periodos:
    df_f = df_f[df_f["Periodo"].isin(periodo_sel)].copy()

# ==========================
# Header
# ==========================

st.markdown('<div class="title">Dashboard P&L 2026</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Base gerencial preparada para acompanhamento de realizado, orçado e variações.</div>',
    unsafe_allow_html=True,
)

# ==========================
# KPIs
# ==========================

df_real = df_f[df_f["Metrica"].str.lower().eq("realizado")]
df_orc = df_f[df_f["Metrica"].str.lower().isin(["orçado", "orcado"])]

linha_receita = escolher_linha(df, ["receita", "produto", "margem financeira"])
linha_resultado = escolher_linha(df, ["resultado", "lucro", "ebitda"])
linha_despesa = escolher_linha(df, ["despesa", "opex", "custo"])

def valor_linha(base, linha):
    if linha is None:
        return 0
    return base[base["Linha_PnL"] == linha]["Valor"].sum()

receita_real = valor_linha(df_real, linha_receita)
resultado_real = valor_linha(df_real, linha_resultado)
despesa_real = valor_linha(df_real, linha_despesa)
total_real = df_real["Valor"].sum()
total_orc = df_orc["Valor"].sum()
var_total = total_real - total_orc if len(df_orc) else 0

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    card("Receita / Produto", fmt_moeda(receita_real), linha_receita or "Realizado")
with c2:
    card("Resultado", fmt_moeda(resultado_real), linha_resultado or "Realizado")
with c3:
    card("Despesas / Custos", fmt_moeda(despesa_real), linha_despesa or "Realizado")
with c4:
    card("Total Realizado", fmt_moeda(total_real), visao_sel)
with c5:
    card("Var. vs Orçado", fmt_moeda(var_total), "Realizado - Orçado")

# ==========================
# Abas
# ==========================

tab1, tab2, tab3, tab4 = st.tabs(["Resumo Executivo", "P&L por Período", "Análise Detalhada", "Base"])

with tab1:
    st.markdown('<div class="section-title">Resultado por linha do P&L</div>', unsafe_allow_html=True)

    resumo = (
        df_f.groupby(["Linha_PnL", "Metrica"], as_index=False)["Valor"]
        .sum()
        .pivot(index="Linha_PnL", columns="Metrica", values="Valor")
        .fillna(0)
        .reset_index()
    )

    cols_valor = [c for c in resumo.columns if c != "Linha_PnL"]
    for c in cols_valor:
        resumo[c] = resumo[c].astype(float)

    if {"Realizado", "Orçado"}.issubset(set(resumo.columns)):
        resumo["Variação R$"] = resumo["Realizado"] - resumo["Orçado"]
        resumo["Variação %"] = resumo.apply(
            lambda r: (r["Variação R$"] / abs(r["Orçado"])) if r["Orçado"] else 0,
            axis=1,
        )

    st.dataframe(resumo, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">Distribuição por produto</div>', unsafe_allow_html=True)
    prod = df_real.groupby("Produto", as_index=False)["Valor"].sum()
    if not prod.empty:
        fig = px.bar(prod, x="Produto", y="Valor", text_auto=".2s")
        fig.update_layout(
            template="plotly_dark",
            paper_bgcolor="#080f1f",
            plot_bgcolor="#080f1f",
            margin=dict(l=10, r=10, t=20, b=10),
            height=360,
        )
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.markdown('<div class="section-title">Evolução por período</div>', unsafe_allow_html=True)

    if "Periodo" in df_f.columns and df_f["Periodo"].astype(str).str.len().gt(0).any():
        evo = df_f.groupby(["Periodo", "Metrica"], as_index=False)["Valor"].sum()
        fig = px.line(evo, x="Periodo", y="Valor", color="Metrica", markers=True)
        fig.update_layout(
            template="plotly_dark",
            paper_bgcolor="#080f1f",
            plot_bgcolor="#080f1f",
            margin=dict(l=10, r=10, t=20, b=10),
            height=420,
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("A coluna Período não possui valores suficientes para montar a evolução.")

with tab3:
    st.markdown('<div class="section-title">Detalhamento da base filtrada</div>', unsafe_allow_html=True)
    st.dataframe(df_f, use_container_width=True, hide_index=True)

    csv = df_f.to_csv(index=False, sep=";").encode("utf-8-sig")
    st.download_button(
        "Baixar base filtrada em CSV",
        data=csv,
        file_name="base_pnl_filtrada.csv",
        mime="text/csv",
    )

with tab4:
    st.markdown('<div class="section-title">Informações da base</div>', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="info-box">
            Linhas carregadas: <b>{len(df):,}</b><br>
            Linhas após filtros: <b>{len(df_f):,}</b><br>
            Aba usada: <b>{ABA_BASE}</b><br>
            Arquivo padrão: <b>{ARQUIVO_PADRAO}</b>
        </div>
        """.replace(",", "."),
        unsafe_allow_html=True,
    )
    st.dataframe(df.head(200), use_container_width=True, hide_index=True)
