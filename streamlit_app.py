import re
import unicodedata
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Dashboard P&L 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_PADRAO = "2026_03_PL_com_BASE_DASH_v2.xlsx"
ABA_RESULTADO = "RESULTADO"
ABA_BASE = "BASE_DASH"
DATA_MINIMA_DASH = pd.Timestamp(2026, 1, 1)

CSS = """
<style>
    .stApp { background: #080f1f; color: #e5ecff; }
    [data-testid="stSidebar"] { background: #0b1224; border-right: 1px solid #1e2a44; }
    [data-testid="stHeader"] { background: rgba(8, 15, 31, .95); }
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
    .dash-title { font-size: 2.25rem; font-weight: 850; color: #ffffff; letter-spacing: .2px; margin-bottom: .2rem; }
    .dash-subtitle { color: #9fb2df; font-size: .95rem; margin-bottom: 1.3rem; }
    .section-title { color: #ffffff; font-size: 1.25rem; font-weight: 750; margin-top: 1.1rem; margin-bottom: .6rem; }
    .kpi-card {
        background: #111a2e;
        border: 1px solid #243150;
        border-radius: 16px;
        padding: 18px 18px;
        min-height: 118px;
        box-shadow: 0 10px 26px rgba(0,0,0,.20);
        text-align: center;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    .card-row-spacer {
        height: 14px;
    }
    table.pnl-matrix {
        width: 100%;
        border-collapse: collapse;
        background: #080f1f;
        color: #e5ecff;
        font-size: .88rem;
    }
    table.pnl-matrix thead th {
        background: #111a2e;
        color: #ffffff;
        font-weight: 850;
        text-align: center;
        padding: 11px 10px;
        border-right: 1px solid rgba(255,255,255,.55);
        border-bottom: 1px solid rgba(255,255,255,.72);
        white-space: nowrap;
    }
    table.pnl-matrix tbody td {
        background: #080f1f;
        color: #e5ecff;
        text-align: center;
        vertical-align: middle;
        padding: 10px 10px;
        border-right: 1px solid rgba(255,255,255,.38);
        border-bottom: 1px solid rgba(255,255,255,.28);
        white-space: nowrap;
        font-weight: 400;
    }
    table.pnl-matrix tbody td:first-child {
        text-align: left;
        color: #ffffff;
        font-weight: 850;
        min-width: 250px;
    }
    table.pnl-matrix tbody tr.main-line td {
        background: #162338;
        color: #ffffff;
        font-weight: 400;
    }
    table.pnl-matrix tbody tr.main-line td:first-child {
        font-weight: 900;
    }
    table.pnl-matrix tbody tr.result-line td {
        background: #1d2d48;
        color: #ffffff;
        font-weight: 950;
        font-size: .95rem;
    }
    table.pnl-matrix td.neg-value {
        color: #ef4444;
        font-weight: 400;
    }
    table.pnl-matrix tbody tr.result-line td.neg-value {
        font-weight: 950;
    }
    table.pnl-matrix td.delta-positive,
    table.pnl-matrix tbody tr.main-line td.delta-positive {
        color: #22c55e !important;
        font-weight: 400 !important;
    }
    table.pnl-matrix td.delta-negative,
    table.pnl-matrix tbody tr.main-line td.delta-negative {
        color: #ef4444 !important;
        font-weight: 400 !important;
    }
    table.pnl-matrix tbody tr.result-line td.delta-positive {
        color: #22c55e !important;
        font-weight: 950 !important;
    }
    table.pnl-matrix tbody tr.result-line td.delta-negative {
        color: #ef4444 !important;
        font-weight: 950 !important;
    }
    table.pnl-matrix th.product-header {
        background: #101a2d;
        font-size: .95rem;
        letter-spacing: .2px;
    }
    table.pnl-matrix th.metric-header {
        background: #162338;
        font-size: .84rem;
    }
    .kpi-label { color: #9fb2df; font-size: .78rem; margin-bottom: 10px; }
    .kpi-value { color: #ffffff; font-size: 1.65rem; font-weight: 850; line-height: 1.15; }
    .kpi-help { color: #60759f; font-size: .72rem; margin-top: 9px; }
    .side-card {
        background: #111a2e;
        border: 1px solid #243150;
        border-radius: 16px;
        padding: 22px 20px;
        min-height: 245px;
        box-shadow: 0 10px 26px rgba(0,0,0,.20);
        text-align: center;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        margin-top: 10px;
    }
    .side-card-label {
        color: #9fb2df;
        font-size: .86rem;
        font-weight: 700;
        margin-bottom: 12px;
    }
    .side-card-value {
        color: #ffffff;
        font-size: 2.05rem;
        font-weight: 900;
        line-height: 1.1;
    }
    .side-card-delta {
        font-size: 1.25rem;
        font-weight: 900;
        margin-top: 12px;
    }
    .side-card-help {
        color: #60759f;
        font-size: .78rem;
        margin-top: 12px;
        line-height: 1.25;
    }
    .kpi-delta {
        font-size: .88rem;
        font-weight: 800;
        margin-top: 9px;
        line-height: 1.15;
    }
    .delta-positive { color: #22c55e; }
    .delta-negative { color: #ef4444; }
    .delta-neutral { color: #9fb2df; }
    .note-box {
        background: #111a2e;
        border: 1px solid #243150;
        border-radius: 14px;
        padding: 13px 16px;
        color: #9fb2df;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; border-bottom: 1px solid #243150; }
    .stTabs [data-baseweb="tab"] { color: #9fb2df; background: transparent; }
    .stTabs [aria-selected="true"] { color: #ffffff; border-bottom: 2px solid #24a8ff; }

    .table-wrap {
        width: 100%;
        overflow-x: auto;
        border: 1px solid rgba(255,255,255,.58);
        border-radius: 14px;
        background: #080f1f;
        box-shadow: 0 10px 26px rgba(0,0,0,.20);
    }
    table.dash-table {
        width: 100%;
        border-collapse: collapse;
        background: #080f1f;
        color: #e5ecff;
        font-size: .92rem;
    }
    table.dash-table thead th {
        background: #111a2e;
        color: #ffffff;
        font-weight: 850;
        font-size: 1.05rem;
        text-align: center;
        padding: 15px 14px;
        border-right: 1px solid rgba(255,255,255,.52);
        border-bottom: 1px solid rgba(255,255,255,.70);
        white-space: nowrap;
    }
    table.dash-table thead th:first-child {
        text-align: center;
        min-width: 260px;
    }
    table.dash-table tbody td {
        background: #080f1f;
        color: #e5ecff;
        text-align: center;
        vertical-align: middle;
        padding: 12px 14px;
        border-right: 1px solid rgba(255,255,255,.42);
        border-bottom: 1px solid rgba(255,255,255,.32);
        white-space: nowrap;
    }
    table.dash-table tbody td:first-child {
        color: #ffffff;
        font-weight: 800;
        text-align: left;
    }
    table.dash-table tbody tr:last-child td {
        border-bottom: none;
    }
    table.dash-table th:last-child,
    table.dash-table td:last-child {
        border-right: none;
    }
    table.dash-table tbody tr:hover td {
        background: #111a2e;
    }
    table.dash-table tbody tr.total-row td {
        font-size: 1.04rem;
        font-weight: 900;
        background: #111a2e;
        color: #ffffff;
    }
    table.dash-table tbody tr.total-row td:first-child {
        text-align: center;
    }
    table.dash-table td.neg-value {
        color: #ef4444;
        font-weight: 900;
    }
    table.dash-table td.delta-positive {
        color: #22c55e;
        font-weight: 900;
    }
    table.dash-table td.delta-negative {
        color: #ef4444;
        font-weight: 900;
    }
    table.dash-table td.delta-neutral {
        color: #9fb2df;
        font-weight: 850;
    }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def formatar_moeda(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    sinal = "-" if valor < 0 else ""
    valor_abs = abs(valor)
    if valor_abs >= 1_000_000_000:
        texto = f"{sinal}R$ {valor_abs / 1_000_000_000:,.2f} bi"
    elif valor_abs >= 1_000_000:
        texto = f"{sinal}R$ {valor_abs / 1_000_000:,.2f} mi"
    elif valor_abs >= 1_000:
        texto = f"{sinal}R$ {valor_abs / 1_000:,.1f} mil"
    else:
        texto = f"{sinal}R$ {valor_abs:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_numero(valor):
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
    except Exception:
        return str(valor)
    return f"{valor:,.0f}".replace(",", ".")


def formatar_percentual(valor):
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
    except Exception:
        return str(valor)

    sinal = "+" if valor > 0 else ""
    texto = f"{sinal}{valor * 100:,.1f}%"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def tabela_html(df, df_valores=None, coluna_delta="Δ mês anterior"):
    html = ['<div class="table-wrap"><table class="dash-table">']
    html.append("<thead><tr>")
    for col in df.columns:
        html.append(f"<th>{col}</th>")
    html.append("</tr></thead><tbody>")

    for idx, row in df.iterrows():
        classe_linha = ' class="total-row"' if str(row.iloc[0]).strip().lower() == "resultado total" else ""
        html.append(f"<tr{classe_linha}>")

        for col in df.columns:
            classes = []

            if col == coluna_delta and df_valores is not None:
                valor_delta = df_valores.loc[idx, col]
                if pd.notna(valor_delta):
                    if valor_delta > 0:
                        classes.append("delta-positive")
                    elif valor_delta < 0:
                        classes.append("delta-negative")
                    else:
                        classes.append("delta-neutral")
            elif col != "Empresa" and df_valores is not None and col in df_valores.columns:
                valor = df_valores.loc[idx, col]
                if pd.notna(valor) and valor < 0:
                    classes.append("neg-value")

            classe_td = f' class="{" ".join(classes)}"' if classes else ""
            html.append(f"<td{classe_td}>{row[col]}</td>")

        html.append("</tr>")

    html.append("</tbody></table></div>")
    return "".join(html)


def converter_periodo(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, pd.Timestamp):
        return valor.to_period("M").to_timestamp()

    if hasattr(valor, "year") and hasattr(valor, "month"):
        try:
            return pd.Timestamp(valor.year, valor.month, 1)
        except Exception:
            pass

    texto = str(valor).strip().lower()
    if not texto or texto == "nan":
        return None

    meses = {
        "jan": 1, "janeiro": 1,
        "fev": 2, "fevereiro": 2,
        "mar": 3, "marco": 3, "março": 3,
        "abr": 4, "abril": 4,
        "mai": 5, "maio": 5,
        "jun": 6, "junho": 6,
        "jul": 7, "julho": 7,
        "ago": 8, "agosto": 8,
        "set": 9, "sep": 9, "setembro": 9,
        "out": 10, "oct": 10, "outubro": 10,
        "nov": 11, "novembro": 11,
        "dez": 12, "dec": 12, "dezembro": 12,
    }

    texto_sem_acento = normalizar_texto(texto)
    partes = texto_sem_acento.split()
    mes = None
    ano = None

    for parte in partes:
        if parte in meses:
            mes = meses[parte]
        elif re.fullmatch(r"\d{4}", parte):
            ano = int(parte)
        elif re.fullmatch(r"\d{2}", parte):
            ano = 2000 + int(parte)

    if mes and ano:
        return pd.Timestamp(ano, mes, 1)

    tentativa = pd.to_datetime(texto, errors="coerce", dayfirst=True)
    if pd.notna(tentativa):
        return tentativa.to_period("M").to_timestamp()

    return None


def nome_periodo(data):
    if pd.isna(data):
        return ""
    meses = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
    data = pd.Timestamp(data)
    return f"{meses[data.month - 1]}/{data.year}"


def formatar_variacao(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0

    sinal = "+" if valor > 0 else ""
    texto = f"{sinal}{valor * 100:,.1f}%"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def classe_variacao(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0

    if valor > 0:
        return "delta-positive"
    if valor < 0:
        return "delta-negative"
    return "delta-neutral"


def card(titulo, valor, ajuda="", variacao=None):
    delta_html = ""
    if variacao is not None:
        delta_html = f'<div class="kpi-delta {classe_variacao(variacao)}">{formatar_variacao(variacao)}</div>'

    ajuda_html = f'<div class="kpi-help">{ajuda}</div>' if ajuda else ""

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{titulo}</div>
            <div class="kpi-value">{formatar_moeda(valor)}</div>
            {delta_html}
            {ajuda_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def carregar_resultado(arquivo):
    bruto = pd.read_excel(arquivo, sheet_name=ABA_RESULTADO, header=None, engine="openpyxl")
    bruto = bruto.dropna(how="all")

    linha_mes = None
    col_rotulo = None

    for idx in bruto.index:
        for col in bruto.columns:
            if normalizar_texto(bruto.loc[idx, col]) == "mes":
                linha_mes = idx
                col_rotulo = col
                break
        if linha_mes is not None:
            break

    if linha_mes is None or col_rotulo is None:
        raise ValueError("Não encontrei a célula com 'Mês' na aba RESULTADO.")

    colunas_periodo = []
    for col in bruto.columns:
        if col <= col_rotulo:
            continue
        periodo = converter_periodo(bruto.loc[linha_mes, col])
        if periodo is not None:
            colunas_periodo.append((col, periodo))

    if not colunas_periodo:
        raise ValueError("Encontrei a célula 'Mês', mas não encontrei meses válidos na mesma linha.")

    registros = []
    ordem_linha = 0

    for idx in bruto.index:
        if idx <= linha_mes:
            continue

        linha_nome = bruto.loc[idx, col_rotulo]
        if pd.isna(linha_nome) or str(linha_nome).strip() == "":
            continue

        linha_tem_valor = False
        for col, periodo in colunas_periodo:
            valor = pd.to_numeric(bruto.loc[idx, col], errors="coerce")
            if pd.notna(valor):
                linha_tem_valor = True
                registros.append(
                    {
                        "Linha": str(linha_nome).strip(),
                        "Linha_Normalizada": normalizar_texto(linha_nome),
                        "Data": periodo,
                        "Período": nome_periodo(periodo),
                        "Valor": float(valor),
                        "Ordem_Linha": ordem_linha,
                    }
                )

        if linha_tem_valor:
            ordem_linha += 1

    df = pd.DataFrame(registros)

    if df.empty:
        raise ValueError("A aba RESULTADO foi encontrada, mas nenhum valor numérico foi lido.")

    df = df[df["Data"] >= DATA_MINIMA_DASH].copy()

    if df.empty:
        raise ValueError("A aba RESULTADO não possui dados a partir de janeiro/2026.")

    return df


@st.cache_data(show_spinner=False)
def carregar_base_dash(arquivo):
    df = pd.read_excel(arquivo, sheet_name=ABA_BASE, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in ["Visao", "Linha_PnL", "Produto", "Metrica", "Periodo"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    if "Valor" in df.columns:
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
    return df



def obter_periodo_pnl_mensal(arquivo):
    try:
        bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal", header=None, engine="openpyxl")
        for _, row in bruto.iterrows():
            for valor in row.tolist():
                data = converter_periodo(valor)
                if data is not None and pd.Timestamp(data).year >= 2020:
                    return nome_periodo(data), pd.Timestamp(data)
    except Exception:
        pass

    return "Período atual", None


@st.cache_data(show_spinner=False)
def carregar_pnl_mensal(arquivo):
    bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal", header=None, engine="openpyxl")
    bruto = bruto.dropna(how="all")

    linha_produto = None
    produtos_encontrados = {}

    for idx in bruto.index:
        for col in bruto.columns:
            texto = normalizar_texto(bruto.loc[idx, col])
            if texto in ["consignado", "imobiliario", "total"]:
                nome = {
                    "consignado": "Consignado",
                    "imobiliario": "Imobiliário",
                    "total": "Total",
                }[texto]
                produtos_encontrados[nome] = col

        if {"Consignado", "Imobiliário", "Total"}.issubset(set(produtos_encontrados.keys())):
            linha_produto = idx
            break

    if linha_produto is None:
        raise ValueError("Não encontrei os blocos Consignado, Imobiliário e Total na aba P&L Mensal.")

    linha_metrica = linha_produto + 1
    col_rotulo = min(produtos_encontrados.values()) - 1

    produtos_ordenados = sorted(produtos_encontrados.items(), key=lambda x: x[1])
    blocos = []

    for i, (produto, col_inicio) in enumerate(produtos_ordenados):
        col_fim = produtos_ordenados[i + 1][1] if i + 1 < len(produtos_ordenados) else max(bruto.columns) + 1

        for col in range(col_inicio, col_fim):
            metrica_norm = normalizar_texto(bruto.loc[linha_metrica, col]) if col in bruto.columns else ""

            if metrica_norm == "realizado":
                metrica = "Realizado"
            elif metrica_norm == "orcado":
                metrica = "Orçado"
            elif metrica_norm in ["", "nan"]:
                continue
            elif "r" in metrica_norm and ("delta" in metrica_norm or metrica_norm == "r"):
                metrica = "Δ R$"
            elif "%" in str(bruto.loc[linha_metrica, col]) or "delta" in metrica_norm or metrica_norm in ["", ""]:
                metrica = "Δ %"
            else:
                metrica = str(bruto.loc[linha_metrica, col]).strip()

            blocos.append({"Produto": produto, "Coluna": col, "Métrica": metrica})

    registros = []
    ordem = 0

    for idx in bruto.index:
        if idx <= linha_metrica:
            continue

        linha_nome = bruto.loc[idx, col_rotulo] if col_rotulo in bruto.columns else None
        if pd.isna(linha_nome) or str(linha_nome).strip() == "":
            continue

        linha_tem_valor = False

        for bloco in blocos:
            col = bloco["Coluna"]
            if col not in bruto.columns:
                continue

            valor = pd.to_numeric(bruto.loc[idx, col], errors="coerce")
            if pd.notna(valor):
                linha_tem_valor = True
                registros.append(
                    {
                        "Produto": bloco["Produto"],
                        "Linha": str(linha_nome).strip(),
                        "Linha_Normalizada": normalizar_texto(linha_nome),
                        "Métrica": bloco["Métrica"],
                        "Valor": float(valor),
                        "Ordem_Linha": ordem,
                    }
                )

        if linha_tem_valor:
            ordem += 1

    df = pd.DataFrame(registros)

    if df.empty:
        raise ValueError("A aba P&L Mensal foi encontrada, mas nenhum valor numérico foi lido.")

    return df


def obter_linhas_principais_pnl(df_pnl):
    linhas_desejadas = [
        "RECEITAS",
        "Operações de Crédito",
        "DESPESAS DE ORIGINAÇÃO",
        "MARGEM INTERMEDIAÇÃO",
        "MG INTERMEDIAÇÃO LIQ",
        "MG CONTRIBUIÇÃO DIRETA",
        "RESULTADO ANTES IMPOSTO",
        "RESULTADO CONTÁBIL",
    ]

    disponiveis = df_pnl[["Linha", "Linha_Normalizada", "Ordem_Linha"]].drop_duplicates().sort_values("Ordem_Linha")

    selecionadas = []
    for linha in linhas_desejadas:
        alvo = normalizar_texto(linha)
        match = disponiveis[disponiveis["Linha_Normalizada"].eq(alvo)]

        if match.empty:
            match = disponiveis[disponiveis["Linha_Normalizada"].str.contains(alvo, regex=False, na=False)]

        if not match.empty:
            selecionadas.append(match.iloc[0]["Linha"])

    return selecionadas


def valor_pnl(df, produto, linha, metrica):
    base = df[
        (df["Produto"] == produto)
        & (df["Linha"] == linha)
        & (df["Métrica"] == metrica)
    ]

    if base.empty:
        return 0

    return base["Valor"].sum()


def card_pnl(titulo, valor, variacao=None):
    delta_html = ""
    if variacao is not None and pd.notna(variacao):
        delta_html = f'<div class="kpi-delta {classe_variacao(variacao)}">{formatar_variacao(variacao)}</div>'

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{titulo}</div>
            <div class="kpi-value">{formatar_moeda(valor)}</div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def montar_tabela_pnl_principal(df_pnl, linhas_principais):
    base = df_pnl[df_pnl["Linha"].isin(linhas_principais)].copy()

    tabela = (
        base.pivot_table(
            index=["Produto", "Linha"],
            columns="Métrica",
            values="Valor",
            aggfunc="sum",
        )
        .reset_index()
    )

    for col in ["Realizado", "Orçado", "Δ %", "Δ R$"]:
        if col not in tabela.columns:
            tabela[col] = pd.NA

    ordem_linhas = {linha: i for i, linha in enumerate(linhas_principais)}
    ordem_produtos = {"Consignado": 1, "Imobiliário": 2, "Total": 3}
    tabela["ordem_linha"] = tabela["Linha"].map(ordem_linhas)
    tabela["ordem_produto"] = tabela["Produto"].map(ordem_produtos)
    tabela = tabela.sort_values(["ordem_produto", "ordem_linha"]).drop(columns=["ordem_produto", "ordem_linha"])

    return tabela[["Produto", "Linha", "Realizado", "Orçado", "Δ %", "Δ R$"]]


def tabela_html_pnl(df, df_valores=None):
    html = ['<div class="table-wrap"><table class="dash-table">']
    html.append("<thead><tr>")
    for col in df.columns:
        html.append(f"<th>{col}</th>")
    html.append("</tr></thead><tbody>")

    for idx, row in df.iterrows():
        is_total = str(row.get("Linha", "")).strip().lower() in ["resultado contábil", "resultado contabil"]
        classe_linha = ' class="total-row"' if is_total else ""
        html.append(f"<tr{classe_linha}>")

        for col in df.columns:
            classes = []
            if df_valores is not None and col in df_valores.columns and col not in ["Produto", "Linha"]:
                valor = df_valores.loc[idx, col]
                if pd.notna(valor) and isinstance(valor, (int, float)) and valor < 0:
                    classes.append("neg-value")

            classe_td = f' class="{" ".join(classes)}"' if classes else ""
            html.append(f"<td{classe_td}>{row[col]}</td>")

        html.append("</tr>")

    html.append("</tbody></table></div>")
    return "".join(html)


def montar_matriz_pnl_excel(df_pnl, linhas_principais):
    produtos = ["Consignado", "Imobiliário", "Total"]
    metricas_por_produto = {
        "Consignado": ["Realizado", "Orçado", "Δ %"],
        "Imobiliário": ["Realizado", "Orçado", "Δ %"],
        "Total": ["Realizado", "Orçado", "Δ %", "Δ R$"],
    }

    linhas = []

    for linha in linhas_principais:
        row = {"Linha": linha}

        for produto in produtos:
            realizado = valor_pnl(df_pnl, produto, linha, "Realizado")
            orcado = valor_pnl(df_pnl, produto, linha, "Orçado")

            delta_rs = realizado - orcado
            delta_pct = pd.NA if orcado == 0 else delta_rs / abs(orcado)

            row[(produto, "Realizado")] = realizado
            row[(produto, "Orçado")] = orcado
            row[(produto, "Δ %")] = delta_pct

            # Regra de cor:
            # Para linhas de despesa/custo negativas, compara o tamanho do gasto em módulo.
            # Para as demais linhas, compara Realizado > Orçado.
            if realizado < 0 or orcado < 0:
                delta_bad = abs(realizado) > abs(orcado)
            else:
                delta_bad = realizado > orcado

            row[(produto, "_delta_bad")] = delta_bad

            if produto == "Total":
                row[(produto, "Δ R$")] = delta_rs

        linhas.append(row)

    return pd.DataFrame(linhas), produtos, metricas_por_produto


def tabela_html_pnl_matriz(df_matrix, produtos, metricas_por_produto):
    linhas_destaque = {
        normalizar_texto("RECEITAS"),
        normalizar_texto("Operações de Crédito"),
        normalizar_texto("DESPESAS DE ORIGINAÇÃO"),
        normalizar_texto("MARGEM INTERMEDIAÇÃO"),
        normalizar_texto("MG INTERMEDIAÇÃO LIQ"),
        normalizar_texto("MG CONTRIBUIÇÃO DIRETA"),
        normalizar_texto("RESULTADO ANTES IMPOSTO"),
        normalizar_texto("RESULTADO CONTÁBIL"),
    }

    html = ['<div class="table-wrap"><table class="pnl-matrix">']

    html.append("<thead>")
    html.append("<tr>")
    html.append('<th rowspan="2">Linha P&L</th>')
    for produto in produtos:
        html.append(f'<th class="product-header" colspan="{len(metricas_por_produto[produto])}">{produto.upper()}</th>')
    html.append("</tr>")

    html.append("<tr>")
    for produto in produtos:
        for metrica in metricas_por_produto[produto]:
            html.append(f'<th class="metric-header">{metrica}</th>')
    html.append("</tr>")
    html.append("</thead><tbody>")

    for _, row in df_matrix.iterrows():
        linha = row["Linha"]
        linha_norm = normalizar_texto(linha)

        if linha_norm in ["resultado contabil", "resultado contábil"]:
            classe = "result-line"
        elif linha_norm in linhas_destaque:
            classe = "main-line"
        else:
            classe = ""

        tr_class = f' class="{classe}"' if classe else ""
        html.append(f"<tr{tr_class}>")
        html.append(f"<td>{linha}</td>")

        for produto in produtos:
            for metrica in metricas_por_produto[produto]:
                valor = row[(produto, metrica)]
                classes = []

                if metrica == "Δ %":
                    texto = formatar_percentual(valor)
                    if pd.notna(valor):
                        classes.append("delta-negative" if row[(produto, "_delta_bad")] else "delta-positive")
                elif metrica == "Δ R$":
                    texto = formatar_numero(valor)
                    if pd.notna(valor):
                        classes.append("delta-negative" if row[(produto, "_delta_bad")] else "delta-positive")
                else:
                    texto = formatar_numero(valor)
                    if pd.notna(valor) and valor < 0:
                        classes.append("neg-value")

                classe_td = f' class="{" ".join(classes)}"' if classes else ""
                html.append(f"<td{classe_td}>{texto}</td>")

        html.append("</tr>")

    html.append("</tbody></table></div>")
    return "".join(html)


def achar_linha_exata_ou_contendo(df, termos):
    linhas = df[["Linha", "Linha_Normalizada", "Ordem_Linha"]].drop_duplicates().sort_values("Ordem_Linha")
    for termo in termos:
        termo_norm = normalizar_texto(termo)
        exato = linhas[linhas["Linha_Normalizada"].eq(termo_norm)]
        if not exato.empty:
            return exato.iloc[0]["Linha"]
    for termo in termos:
        termo_norm = normalizar_texto(termo)
        encontrado = linhas[linhas["Linha_Normalizada"].str.contains(termo_norm, regex=False, na=False)]
        if not encontrado.empty:
            return encontrado.iloc[0]["Linha"]
    return None


def montar_resultados_principais(df):
    specs = [
        ("Resultado Conglomerado Financeiro", ["resultado congl financeiro", "resultado conglomerado financeiro"]),
        ("Resultado Coligadas", ["resultado coligadas"]),
        ("Resultado Conglomerado + Coligadas", ["resultado congl coligadas", "resultado conglomerado coligadas"]),
        ("Resultado Total", ["res total", "resultado total"]),
    ]
    mapeamento = []
    for titulo, termos in specs:
        linha = achar_linha_exata_ou_contendo(df, termos)
        if linha:
            mapeamento.append({"Indicador": titulo, "Linha": linha})

    if not mapeamento:
        return pd.DataFrame(columns=["Indicador", "Linha", "Data", "Período", "Valor"])

    mapa = pd.DataFrame(mapeamento)
    return df.merge(mapa, on="Linha", how="inner")



def periodo_anterior(periodos_df, periodo_atual):
    linha_atual = periodos_df[periodos_df["Período"] == periodo_atual]
    if linha_atual.empty:
        return None

    data_atual = linha_atual["Data"].iloc[0]
    anteriores = periodos_df[periodos_df["Data"] < data_atual].sort_values("Data")

    if anteriores.empty:
        return None

    return anteriores.iloc[-1]["Período"]


def variacao_mes_anterior(df_principais, indicador, periodo_atual, periodo_ant):
    if periodo_ant is None:
        return None

    valor_atual = df_principais[
        (df_principais["Indicador"] == indicador)
        & (df_principais["Período"] == periodo_atual)
    ]["Valor"].sum()

    valor_ant = df_principais[
        (df_principais["Indicador"] == indicador)
        & (df_principais["Período"] == periodo_ant)
    ]["Valor"].sum()

    if valor_ant == 0:
        return None

    return (valor_atual - valor_ant) / abs(valor_ant)


def resultado_total_acumulado_ano(df_principais, periodo_atual):
    linha_atual = df_principais[
        (df_principais["Indicador"] == "Resultado Total")
        & (df_principais["Período"] == periodo_atual)
    ]

    if linha_atual.empty:
        return None, None, None, None

    data_atual = linha_atual["Data"].iloc[0]
    ano_atual = pd.Timestamp(data_atual).year
    data_inicio = pd.Timestamp(ano_atual, 1, 1)

    base_ano = df_principais[
        (df_principais["Indicador"] == "Resultado Total")
        & (df_principais["Data"] >= data_inicio)
        & (df_principais["Data"] <= data_atual)
    ].copy()

    if base_ano.empty:
        return None, None, None, None

    valor_acumulado = base_ano["Valor"].sum()

    data_mes_anterior = pd.Timestamp(data_atual) - pd.DateOffset(months=1)
    base_ate_mes_anterior = base_ano[base_ano["Data"] <= data_mes_anterior]

    valor_acumulado_anterior = base_ate_mes_anterior["Valor"].sum() if not base_ate_mes_anterior.empty else None

    if valor_acumulado_anterior is None or valor_acumulado_anterior == 0:
        variacao = None
    else:
        variacao = (valor_acumulado - valor_acumulado_anterior) / abs(valor_acumulado_anterior)

    return valor_acumulado, variacao, valor_acumulado_anterior, data_inicio


def card_resultado_total_acumulado(valor_acumulado, variacao, valor_acumulado_anterior, periodo_atual):
    if valor_acumulado is None:
        valor_html = "N/D"
        delta_html = '<div class="side-card-delta delta-neutral">N/D</div>'
        ajuda = "Resultado Total não encontrado"
    else:
        valor_html = formatar_moeda(valor_acumulado)

        if variacao is None:
            delta_html = '<div class="side-card-delta delta-neutral">N/D</div>'
        else:
            delta_html = f'<div class="side-card-delta {classe_variacao(variacao)}">{formatar_variacao(variacao)}</div>'

        if valor_acumulado_anterior is None:
            ajuda = f"Acumulado de jan/2026 até {periodo_atual}"
        else:
            ajuda = f"Acumulado de jan/2026 até {periodo_atual}"

    st.markdown(
        f"""
        <div class="side-card">
            <div class="side-card-label">Resultado Total acumulado em 2026</div>
            <div class="side-card-value">{valor_html}</div>
            {delta_html}
            <div class="side-card-help">{ajuda}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def adicionar_coluna_variacao_tabela(tabela, periodos_df, periodo_atual):
    coluna_delta = "Δ mês anterior"
    tabela = tabela.copy()

    periodo_ant = periodo_anterior(periodos_df, periodo_atual)

    if periodo_ant is None or periodo_ant not in tabela.columns or periodo_atual not in tabela.columns:
        tabela[coluna_delta] = pd.NA
        return tabela, coluna_delta

    atual = pd.to_numeric(tabela[periodo_atual], errors="coerce")
    anterior = pd.to_numeric(tabela[periodo_ant], errors="coerce")

    tabela[coluna_delta] = (atual - anterior) / anterior.abs()
    tabela.loc[anterior.eq(0) | anterior.isna(), coluna_delta] = pd.NA

    return tabela, coluna_delta


def montar_tabela_empresas_e_total(df):
    excluir_exatos = {
        "banco",
        "equiv patr",
        "jcp dividendos",
        "br cards",
        "resultado mep",
        "resultado congl financeiro",
        "resultado conglomerado financeiro",
        "resultado coligadas",
        "resultado congl coligadas",
        "resultado conglomerado coligadas",
    }

    excluir_contem = [
        "resultado congl financeiro",
        "resultado conglomerado financeiro",
        "resultado coligadas",
        "resultado congl coligadas",
        "resultado conglomerado coligadas",
    ]

    renomear = {
        "resultado banco": "Banco",
        "resulta br cards": "BR Cards",
        "resultado br cards": "BR Cards",
        "res total": "Resultado Total",
        "resultado total": "Resultado Total",
    }

    linha_total = achar_linha_exata_ou_contendo(df, ["res total", "resultado total"])

    linhas = df[["Linha", "Linha_Normalizada", "Ordem_Linha"]].drop_duplicates().sort_values("Ordem_Linha").copy()

    def manter(row):
        nome = row["Linha_Normalizada"]
        if linha_total is not None and row["Linha"] == linha_total:
            return True
        if nome in excluir_exatos:
            return False
        if any(term in nome for term in excluir_contem):
            return False
        return True

    linhas_filtradas = linhas[linhas.apply(manter, axis=1)]["Linha"].tolist()
    df_tabela = df[df["Linha"].isin(linhas_filtradas)].copy()

    datas_ordem = df_tabela[["Período", "Data"]].drop_duplicates().sort_values("Data")
    colunas = datas_ordem["Período"].tolist()

    tabela = (
        df_tabela.pivot_table(index="Linha", columns="Período", values="Valor", aggfunc="sum")
        .reindex(index=linhas_filtradas)
        .reindex(columns=colunas)
        .reset_index()
    )

    def nome_exibicao(linha):
        nome_norm = normalizar_texto(linha)
        return renomear.get(nome_norm, linha)

    tabela["Linha"] = tabela["Linha"].map(nome_exibicao)
    return tabela


st.sidebar.title("Filtros")
arquivo_local = Path(ARQUIVO_PADRAO)
upload = st.sidebar.file_uploader("Atualizar base manualmente", type=["xlsx"])

if upload is not None:
    arquivo = upload
elif arquivo_local.exists():
    arquivo = arquivo_local
else:
    st.error(f"Arquivo '{ARQUIVO_PADRAO}' não encontrado no repositório.")
    st.stop()

try:
    df_resultado = carregar_resultado(arquivo)
except Exception as erro:
    st.error(f"Erro ao carregar a aba RESULTADO: {erro}")
    st.stop()

periodos_disponiveis = (
    df_resultado[["Data", "Período"]]
    .drop_duplicates()
    .sort_values("Data")
    .reset_index(drop=True)
)

periodo_padrao = len(periodos_disponiveis) - 1

st.sidebar.markdown(
    """
    <div class="note-box">
        Exibindo somente dados a partir de <b>jan/2026</b>.
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="dash-title">Dashboard P&L 2026</div>', unsafe_allow_html=True)
st.markdown('<div class="dash-subtitle">Resultados consolidados, evolução histórica e abertura por empresa.</div>', unsafe_allow_html=True)

tab_resultados, tab_pnl_mensal, tab_pnl_acum, tab_base = st.tabs(
    ["Resultados", "P&L Mensal", "P&L Acumulado", "Base"]
)

with tab_resultados:
    st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)
    col_filtro_mes, col_filtro_vazio = st.columns([1, 3])
    with col_filtro_mes:
        periodo_sel = st.selectbox(
            "Mês de referência",
            periodos_disponiveis["Período"].tolist(),
            index=periodo_padrao,
        )

    df_principais = montar_resultados_principais(df_resultado)
    df_cards = df_principais[df_principais["Período"] == periodo_sel].copy()
    periodo_ant = periodo_anterior(periodos_disponiveis, periodo_sel)

    st.markdown('<div class="section-title">Principais resultados</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    indicadores = [
        "Resultado Conglomerado Financeiro",
        "Resultado Coligadas",
        "Resultado Conglomerado + Coligadas",
        "Resultado Total",
    ]

    for coluna, indicador in zip([c1, c2, c3, c4], indicadores):
        with coluna:
            linha = df_cards[df_cards["Indicador"] == indicador]
            valor = linha["Valor"].sum() if not linha.empty else 0
            origem = linha["Linha"].iloc[0] if not linha.empty else "Linha não encontrada"
            variacao = variacao_mes_anterior(df_principais, indicador, periodo_sel, periodo_ant)
            card(indicador, valor, variacao=variacao)

    st.markdown('<div class="section-title">Evolução histórica dos resultados</div>', unsafe_allow_html=True)

    if df_principais.empty:
        st.warning("Não encontrei as linhas principais na aba RESULTADO. Verifique os nomes das linhas na planilha.")
    else:
        col_grafico, col_card_variacao = st.columns([3.4, 1])

        with col_grafico:
            fig = px.line(
                df_principais.sort_values("Data"),
                x="Data",
                y="Valor",
                color="Indicador",
                markers=True,
                line_shape="spline",
                labels={"Data": "Mês", "Valor": "Resultado", "Indicador": "Resultado"},
            )
            tick_datas = periodos_disponiveis["Data"].tolist()
            tick_textos = periodos_disponiveis["Período"].tolist()
            fig.update_layout(
                template="plotly_dark",
                paper_bgcolor="#080f1f",
                plot_bgcolor="#080f1f",
                height=430,
                margin=dict(l=10, r=10, t=10, b=10),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
            )
            fig.update_xaxes(
                tickmode="array",
                tickvals=tick_datas,
                ticktext=tick_textos,
                showgrid=False,
                zeroline=False,
            )
            fig.update_yaxes(
                tickprefix="R$ ",
                separatethousands=True,
                showgrid=False,
                zeroline=False,
            )
            st.plotly_chart(fig, use_container_width=True)

        with col_card_variacao:
            valor_acumulado, variacao_acumulado, valor_acumulado_anterior, data_inicio = resultado_total_acumulado_ano(
                df_principais, periodo_sel
            )
            card_resultado_total_acumulado(valor_acumulado, variacao_acumulado, valor_acumulado_anterior, periodo_sel)

    st.markdown('<div class="section-title">Resultado aberto por empresa</div>', unsafe_allow_html=True)

    tabela = montar_tabela_empresas_e_total(df_resultado)
    tabela, coluna_delta = adicionar_coluna_variacao_tabela(tabela, periodos_disponiveis, periodo_sel)

    tabela_valores = tabela.copy()
    tabela_formatada = tabela.copy()

    for col in tabela_formatada.columns:
        if col == coluna_delta:
            tabela_formatada[col] = tabela_formatada[col].map(formatar_percentual)
        elif col != "Linha":
            tabela_formatada[col] = tabela_formatada[col].map(formatar_numero)

    tabela_formatada = tabela_formatada.rename(columns={"Linha": "Empresa"})
    tabela_valores = tabela_valores.rename(columns={"Linha": "Empresa"})

    st.markdown(tabela_html(tabela_formatada, tabela_valores, coluna_delta=coluna_delta), unsafe_allow_html=True)

with tab_pnl_mensal:
    try:
        df_pnl = carregar_pnl_mensal(arquivo)
        linhas_principais = obter_linhas_principais_pnl(df_pnl)

        periodo_pnl, data_pnl = obter_periodo_pnl_mensal(arquivo)

        st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)
        col_data, col_produto, col_espaco = st.columns([1, 1, 2])
        with col_data:
            data_sel_pnl = st.selectbox(
                "Data base",
                [periodo_pnl],
                index=0,
                key="data_pnl_mensal",
            )
        with col_produto:
            produto_sel_pnl = st.selectbox(
                "Produto",
                ["Consignado", "Imobiliário", "Total"],
                index=2,
                key="produto_pnl_mensal",
            )

        st.markdown('<div class="section-title">Principais linhas do P&L Mensal</div>', unsafe_allow_html=True)

        for inicio in range(0, len(linhas_principais), 4):
            if inicio > 0:
                st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)

            cols_cards = st.columns(4)
            for col_card, linha in zip(cols_cards, linhas_principais[inicio:inicio + 4]):
                realizado = valor_pnl(df_pnl, produto_sel_pnl, linha, "Realizado")
                variacao = valor_pnl(df_pnl, produto_sel_pnl, linha, "Δ %")
                with col_card:
                    card_pnl(linha, realizado, variacao=variacao)

        st.markdown('<div class="section-title">Realizado x Orçado por linha principal</div>', unsafe_allow_html=True)

        base_grafico = df_pnl[
            (df_pnl["Produto"] == produto_sel_pnl)
            & (df_pnl["Linha"].isin(linhas_principais))
            & (df_pnl["Métrica"].isin(["Realizado", "Orçado"]))
        ].copy()

        ordem_linhas = {linha: i for i, linha in enumerate(linhas_principais)}
        base_grafico["Ordem"] = base_grafico["Linha"].map(ordem_linhas)
        base_grafico = base_grafico.sort_values("Ordem", ascending=False)

        fig_comp = px.bar(
            base_grafico,
            x="Valor",
            y="Linha",
            color="Métrica",
            orientation="h",
            barmode="group",
            labels={"Valor": "Valor", "Linha": "", "Métrica": ""},
        )
        fig_comp.update_layout(
            template="plotly_dark",
            paper_bgcolor="#080f1f",
            plot_bgcolor="#080f1f",
            height=470,
            margin=dict(l=10, r=10, t=10, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        )
        fig_comp.update_xaxes(showgrid=False, zeroline=False, tickprefix="R$ ", separatethousands=True)
        fig_comp.update_yaxes(showgrid=False, zeroline=False)
        st.plotly_chart(fig_comp, use_container_width=True)

        st.markdown('<div class="section-title">Resultado Contábil por produto</div>', unsafe_allow_html=True)
        linha_resultado_contabil = next(
            (linha for linha in linhas_principais if normalizar_texto(linha) in ["resultado contabil", "resultado contábil"]),
            linhas_principais[-1] if linhas_principais else None,
        )

        base_produtos = df_pnl[
            (df_pnl["Linha"] == linha_resultado_contabil)
            & (df_pnl["Produto"].isin(["Consignado", "Imobiliário", "Total"]))
            & (df_pnl["Métrica"] == "Realizado")
        ].copy()

        fig_prod = px.bar(
            base_produtos,
            x="Produto",
            y="Valor",
            text=base_produtos["Valor"].map(lambda v: formatar_moeda(v).replace("R$ ", "")),
            labels={"Valor": "Realizado", "Produto": ""},
        )
        fig_prod.update_traces(
            textposition="inside",
            textfont=dict(size=18, family="Arial Black"),
            insidetextanchor="middle",
        )
        fig_prod.update_layout(
            template="plotly_dark",
            paper_bgcolor="#080f1f",
            plot_bgcolor="#080f1f",
            height=390,
            margin=dict(l=10, r=10, t=10, b=10),
            showlegend=False,
        )
        fig_prod.update_xaxes(showgrid=False, zeroline=False)
        fig_prod.update_yaxes(showgrid=False, zeroline=False, tickprefix="R$ ", separatethousands=True)
        st.plotly_chart(fig_prod, use_container_width=True)

        st.markdown('<div class="section-title">Resumo das linhas principais por produto</div>', unsafe_allow_html=True)
        matriz_pnl, produtos_matriz, metricas_matriz = montar_matriz_pnl_excel(df_pnl, linhas_principais)
        st.markdown(
            tabela_html_pnl_matriz(matriz_pnl, produtos_matriz, metricas_matriz),
            unsafe_allow_html=True,
        )

    except Exception as erro:
        st.info(f"Não consegui carregar a aba P&L Mensal: {erro}")

with tab_pnl_acum:
    try:
        df_base = carregar_base_dash(arquivo)
        df_acum = df_base[df_base["Visao"].str.lower().eq("acumulado")].copy()
        st.markdown('<div class="section-title">P&L Acumulado</div>', unsafe_allow_html=True)
        st.dataframe(df_acum, use_container_width=True, hide_index=True)
    except Exception as erro:
        st.info(f"Não consegui carregar a aba BASE_DASH para P&L Acumulado: {erro}")

with tab_base:
    st.markdown('<div class="section-title">Base lida da aba RESULTADO</div>', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="note-box">
            Linhas lidas: <b>{len(df_resultado):,}</b><br>
            Períodos encontrados: <b>{", ".join(periodos_disponiveis["Período"].tolist())}</b><br>
            Filtro aplicado: <b>jan/2026 em diante</b><br>
            Aba principal: <b>{ABA_RESULTADO}</b>
        </div>
        """.replace(",", "."),
        unsafe_allow_html=True,
    )
    st.dataframe(df_resultado.drop(columns=["Linha_Normalizada"], errors="ignore"), use_container_width=True, hide_index=True)
