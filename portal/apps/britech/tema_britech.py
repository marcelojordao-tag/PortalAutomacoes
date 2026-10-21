"""Tema visual do app de Conferência de Rentabilidade (identidade TAG Investimentos).

Paleta e componentes seguem o mesmo padrão da Mesa de Boletas: vinho profundo + dourado discreto,
fonte Inter, cartões brancos com borda colorida, KPIs com faixa no topo e tabelas com cabeçalho vinho.

Uso no app:
    from tema import COR, aplicar_tema, topo, esc, kpis_html, badge, aviso_html, tabela_html, passos_html, estilizar_altair
    aplicar_tema()   # logo depois do st.set_page_config
    topo("Título", "Subtítulo")
"""
from __future__ import annotations

import base64
import html
from pathlib import Path

import pandas as pd
import streamlit as st

LOGO = Path(__file__).resolve().with_name("logo_tag.png")  # opcional: se existir, aparece no topo

# Paleta da marca: vinho profundo + branco, com dourado discreto de apoio
COR = {
    "vinho": "#6A0F2B",
    "vinho_escuro": "#4A0A1E",
    "vinho_medio": "#8A1B3D",
    "vinho_claro": "#F4E8EC",
    "fundo": "#F7F4F5",
    "card": "#FFFFFF",
    "texto": "#26151B",
    "texto_suave": "#7A6770",
    "borda": "#E6DCE0",
    "dourado": "#B8924A",
    "ok": "#2E8B57",
    "div": "#C0392B",
    "alerta": "#C98A1B",
}

_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, .stApp, [class*="css"] { font-family: 'Inter', system-ui, sans-serif; }
#MainMenu, footer, header[data-testid="stHeader"], [data-testid="stToolbar"] { visibility: hidden; height: 0; }
.stApp { background: __FUNDO__; color: __TEXTO__; }
.block-container { padding-top: 1rem; padding-bottom: 3rem; max-width: 1320px; }

/* ---------- Topo ---------- */
.topo {
    display: flex; align-items: center; gap: 1.4rem;
    background: linear-gradient(120deg, __VINHO_ESCURO__ 0%, __VINHO__ 60%, __VINHO_MEDIO__ 100%);
    color: #fff; padding: 1rem 1.6rem; border-radius: 16px; margin-bottom: 1.2rem;
    box-shadow: 0 6px 20px rgba(74, 10, 30, .28);
    border-bottom: 3px solid __DOURADO__;
}
.topo img { height: 76px; width: 76px; border-radius: 12px; }
.topo .marca { font-size: 1.8rem; font-weight: 700; letter-spacing: -.01em; margin: 0; color: #fff; line-height: 1.1; }
.topo .sub { margin: .2rem 0 0 0; font-size: .92rem; color: #E9CFD7; }
.topo .selo {
    margin-left: auto; font-size: .72rem; font-weight: 600; letter-spacing: .12em; text-transform: uppercase;
    color: __DOURADO__; border: 1px solid __DOURADO__; border-radius: 999px; padding: .3rem .9rem; white-space: nowrap;
}

/* ---------- Textos de seção ---------- */
.secao { color: __VINHO__; font-size: 1.05rem; font-weight: 700; margin: .2rem 0 .1rem 0; }
.secao-sub { color: __TEXTO_SUAVE__; font-size: .85rem; margin-bottom: .8rem; }
.rotulo { color: __VINHO__; font-size: .74rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; margin: .4rem 0 .25rem 0; }
.fonte-dados { color: __TEXTO_SUAVE__; font-size: .85rem; padding-top: .55rem; }

/* ---------- Botões ---------- */
.stButton > button, .stDownloadButton > button {
    background: __VINHO__; color: #fff; border: 1px solid __VINHO__;
    border-radius: 10px; padding: .5rem 1.2rem; font-weight: 600; transition: all .15s ease;
}
.stButton > button p, .stDownloadButton > button p { color: #fff; font-weight: 600; }
.stButton > button:hover, .stDownloadButton > button:hover {
    background: __VINHO_MEDIO__; border-color: __VINHO_MEDIO__; color: #fff;
    box-shadow: 0 4px 12px rgba(106, 15, 43, .25);
}
.stButton > button:focus:not(:active), .stDownloadButton > button:focus:not(:active) { border-color: __DOURADO__; color: #fff; }
.stButton > button:disabled { background: #D9CDD2; border-color: #D9CDD2; color: #fff; }

/* ---------- Inputs ---------- */
.stTextInput input, .stTextArea textarea, [data-baseweb="select"] > div, .stNumberInput input {
    border-radius: 10px !important; border-color: __BORDA__ !important; background: #fff !important;
}

/* ---------- Filtros em pílula (st.pills) ---------- */
button[aria-pressed] {
    border-radius: 999px; padding: .3rem 1.05rem; font-weight: 600; font-size: .88rem; transition: all .15s ease;
    box-shadow: none;
}
/* !important: o Streamlit usa uma regra [data-selected] mais específica para o estado ativo */
button[aria-pressed="false"] { background: #fff !important; border: 1px solid __BORDA__ !important; color: __VINHO__ !important; }
button[aria-pressed="false"] p { color: __VINHO__ !important; font-weight: 600; }
button[aria-pressed="false"]:hover { border-color: __VINHO__ !important; background: __VINHO_CLARO__ !important; }
button[aria-pressed="true"] { background: __VINHO__ !important; border: 1px solid __VINHO__ !important; color: #fff !important; }
button[aria-pressed="true"] p { color: #fff !important; font-weight: 600; }

/* ---------- KPIs ---------- */
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: .9rem; margin: 0 0 1.1rem 0; }
.kpi {
    background: #fff; border: 1px solid __BORDA__; border-top: 4px solid __VINHO__;
    border-radius: 12px; padding: .95rem 1.25rem; box-shadow: 0 2px 8px rgba(74, 10, 30, .07);
}
.kpi.ok { border-top-color: __OK__; } .kpi.div { border-top-color: __DIV__; } .kpi.alerta { border-top-color: __ALERTA__; }
.kpi-label { color: __TEXTO_SUAVE__; font-size: .74rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
.kpi-valor { color: __TEXTO__; font-size: 2.1rem; font-weight: 700; line-height: 1.2; font-variant-numeric: tabular-nums; }
.kpi-valor.menor { font-size: 1.45rem; padding: .35rem 0; }
.kpi.ok .kpi-valor { color: __OK__; } .kpi.div .kpi-valor { color: __DIV__; } .kpi.alerta .kpi-valor { color: __ALERTA__; }
.kpi-sub { color: __TEXTO_SUAVE__; font-size: .8rem; }

/* ---------- Avisos ---------- */
.aviso {
    border: 1px solid __BORDA__; border-left: 6px solid __VINHO__; background: #fff; border-radius: 12px;
    padding: .8rem 1.1rem; margin: 0 0 1rem 0; color: __TEXTO__; font-size: .92rem; line-height: 1.45;
    box-shadow: 0 2px 8px rgba(74, 10, 30, .06);
}
.aviso b { color: __VINHO__; }
.aviso.ok { border-left-color: __OK__; background: #F2FAF5; } .aviso.ok b { color: __OK__; }
.aviso.alerta { border-left-color: __ALERTA__; background: #FFFAF0; } .aviso.alerta b { color: #8A5A00; }
.aviso.div { border-left-color: __DIV__; background: #FEF5F4; } .aviso.div b { color: __DIV__; }

/* ---------- Selos ---------- */
.badge { display: inline-block; padding: .18rem .7rem; border-radius: 999px; font-size: .72rem; font-weight: 700; letter-spacing: .04em; white-space: nowrap; }
.badge.ok { background: #DDF1E5; color: #1D6B40; }
.badge.div { background: #F9DAD6; color: #9B2C22; }
.badge.neutro { background: #EEE6E9; color: #6B5560; }
.badge.ok-suave { background: #EAF6EE; color: #2E7D52; }
.badge.div-suave { background: #FBE7E4; color: #B5392E; }
.badge.warn { background: #FBEBC8; color: #8A5A00; }

/* ---------- Cartões de carteira ---------- */
[class*="st-key-cartCard_"] {
    background: #fff; border: 1px solid __BORDA__; border-left: 6px solid __VINHO__;
    border-radius: 12px; padding: 1rem 1.1rem; margin-bottom: .4rem; gap: .5rem;
    box-shadow: 0 2px 8px rgba(74, 10, 30, .08);
}
[class*="st-key-cartCard_ok_"] { border-left-color: __OK__; }
[class*="st-key-cartCard_explicada_"] { border-left-color: __ALERTA__; }
[class*="st-key-cartCard_div_"] { border-left-color: __DIV__; }
[class*="st-key-cartCard_"] .stButton > button { width: 100%; }

/* Cartões do monitor de cotas (mesmo desenho dos cartões de carteira, com cor por status) */
[class*="st-key-monCard_"] {
    background: #fff; border: 1px solid __BORDA__; border-left: 6px solid __BORDA__;
    border-radius: 12px; padding: 1rem 1.1rem; margin-bottom: .4rem; gap: .5rem;
    box-shadow: 0 2px 8px rgba(74, 10, 30, .08);
}
[class*="st-key-monCard_acima_"] { border-left-color: __OK__; }
[class*="st-key-monCard_abaixo_"] { border-left-color: __DIV__; }
[class*="st-key-monCard_dentro_"] { border-left-color: __VINHO__; }
[class*="st-key-monCard_sem_"] { border-left-color: #BBA9B0; }
[class*="st-key-monCard_"] .stButton > button { width: 100%; }
.kv b.pos { color: __OK__; }
.card-topo { display: flex; justify-content: space-between; align-items: center; margin-bottom: .35rem; }
.card-id { color: __TEXTO_SUAVE__; font-size: .75rem; font-weight: 700; }
.card-nome { color: __TEXTO__; font-weight: 700; font-size: .98rem; line-height: 1.25; overflow: hidden; text-overflow: ellipsis;
             display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; min-height: 2.5em; margin-bottom: .55rem; }
.kv { display: flex; justify-content: space-between; gap: .75rem; padding: .25rem 0; border-bottom: 1px dashed __BORDA__; font-size: .88rem; }
.kv span { color: __TEXTO_SUAVE__; }
.kv b { font-variant-numeric: tabular-nums; text-align: right; color: __TEXTO__; }
.kv b.neg { color: __DIV__; }
.resumo { font-size: .84rem; color: __TEXTO_SUAVE__; margin: .6rem 0 .75rem 0; min-height: 2.6em; line-height: 1.35; }

/* ---------- Pop-up ---------- */
[data-testid="stDialog"] div[role="dialog"] { background: __FUNDO__; border-radius: 16px; }
[data-testid="stDialog"] h2 { color: __VINHO__; font-weight: 700; }
.dlg-nome { color: __VINHO__; font-size: 1.25rem; font-weight: 700; line-height: 1.25; margin: 0; }
.dlg-meta { color: __TEXTO_SUAVE__; font-size: .85rem; margin: .15rem 0 .6rem 0; }

/* Abas */
[role="tablist"] { gap: 6px !important; border: none !important; background: transparent !important; padding: 0 !important; }
[role="tab"] {
    height: 42px !important; padding: 0 1.2rem !important; margin: 0 !important;
    background: __VINHO_ESCURO__ !important; border: none !important;
    border-radius: 12px 12px 0 0 !important; opacity: 1 !important; transition: background .15s ease; cursor: pointer;
}
[role="tab"] p, [role="tab"] * { color: #E9CFD7 !important; font-weight: 600 !important; font-size: .88rem !important; }
[role="tab"]:hover { background: __VINHO_MEDIO__ !important; }
[role="tab"][aria-selected="true"] { background: __CARD__ !important; box-shadow: inset 0 3px 0 __DOURADO__; }
[role="tab"][aria-selected="true"] p, [role="tab"][aria-selected="true"] * { color: __VINHO__ !important; }
.react-aria-SelectionIndicator, [data-baseweb="tab-highlight"], [data-baseweb="tab-border"] { display: none !important; }
[role="tabpanel"] {
    background: __CARD__ !important; border: 1px solid __BORDA__; border-radius: 0 14px 14px 14px !important;
    padding: 1.2rem 1.4rem 1.5rem 1.4rem !important; box-shadow: 0 4px 14px rgba(74, 10, 30, .06);
}

/* Expansor e popover */
[data-testid="stExpander"] { background: #fff; border: 1px solid __BORDA__; border-radius: 12px; margin-bottom: 1rem; }
[data-testid="stExpander"] details { border: none; }
[data-testid="stExpander"] summary p { color: __VINHO__; font-weight: 600; }

/* ---------- Passo a passo ---------- */
.passos { width: 100%; border-collapse: separate; border-spacing: 0; font-size: .88rem; color: __TEXTO__; }
.passos td { padding: .5rem .7rem; border-bottom: 1px solid #F0E8EB; vertical-align: middle; }
.passos .grupo td { background: __VINHO__; color: #fff; font-weight: 700; font-size: .78rem; letter-spacing: .06em;
                    text-transform: uppercase; border-bottom: 2px solid __DOURADO__; }
.passos .n { width: 2.2rem; color: __TEXTO_SUAVE__; font-weight: 700; text-align: right; }
.passos .val { text-align: right; font-weight: 700; white-space: nowrap; font-variant-numeric: tabular-nums; width: 11rem; }
.passos .expl { color: __TEXTO_SUAVE__; font-size: .82rem; }
.passos tr.destaque td { background: __VINHO_CLARO__; }
.passos tr.destaque .rot, .passos tr.destaque .val { color: __VINHO__; font-weight: 700; }

/* ---------- Tabela ---------- */
.tabela-wrap { background: #fff; border: 1px solid __BORDA__; border-radius: 12px; overflow: auto; box-shadow: 0 2px 8px rgba(74, 10, 30, .07); margin: .6rem 0 1rem 0; }
.tabela { width: 100%; border-collapse: separate; border-spacing: 0; font-size: .86rem; color: __TEXTO__; }
.tabela thead th {
    position: sticky; top: 0; z-index: 2; text-align: left; white-space: nowrap;
    background: __VINHO__; color: #fff; padding: .75rem .95rem; font-size: .76rem; font-weight: 600;
    letter-spacing: .05em; text-transform: uppercase; border-bottom: 2px solid __DOURADO__;
}
.tabela td { padding: .55rem .95rem; border-bottom: 1px solid #F0E8EB; white-space: nowrap; color: __TEXTO__; }
.tabela tbody tr:nth-child(even) td { background: #FBF8F9; }
.tabela tbody tr:hover td { background: __VINHO_CLARO__; }
.tabela .num, .tabela th.num { text-align: right; font-variant-numeric: tabular-nums; }
.tabela.quebra td { white-space: normal; line-height: 1.4; vertical-align: top; }
.tabela.quebra td:first-child { white-space: nowrap; width: 11rem; }
.tabela tr.linha-total td { font-weight: 700; background: __VINHO_CLARO__; }
.tabela tr.linha-geral td { font-weight: 700; background: __VINHO__; color: #fff; }
.tabela tr.linha-div td:first-child { box-shadow: inset 4px 0 0 __DIV__; }
.tabela .neg { color: __DIV__; }
.tabela .pos { color: __OK__; }
.tabela .pos, .tabela .neg { font-weight: 600; }
.tabela tr.linha-geral .neg { color: #FFB4AB; }
.barra { display: inline-block; height: 8px; border-radius: 4px; margin-left: .55rem; vertical-align: middle; min-width: 2px; }
.barra.pos { background: __OK__; } .barra.neg { background: __DIV__; }
.trunc { display: inline-block; max-width: 250px; overflow: hidden; text-overflow: ellipsis; vertical-align: bottom; }
.rank-titulo { font-weight: 700; font-size: .95rem; margin: .2rem 0 .1rem 0; }
.rank-titulo.pos { color: __OK__; } .rank-titulo.neg { color: __DIV__; }
.nota-tabela { color: __TEXTO_SUAVE__; font-size: .8rem; margin: -.4rem 0 .8rem 0; }
</style>
"""


def _pintar(css: str) -> str:
    for chave, valor in COR.items():
        css = css.replace(f"__{chave.upper()}__", valor)
    return css


def aplicar_tema() -> None:
    st.markdown(_pintar(_CSS), unsafe_allow_html=True)


def esc(txt) -> str:
    """Escapa o texto para HTML (e o $, que o Streamlit trata como fórmula)."""
    return html.escape(str(txt).strip()).replace("$", "&#36;")


def _logo_base64() -> str:
    try:
        return base64.b64encode(LOGO.read_bytes()).decode()
    except OSError:
        return ""


def topo(titulo: str, subtitulo: str = "") -> None:
    b64 = _logo_base64()
    img = f'<img src="data:image/png;base64,{b64}" alt="TAG Investimentos">' if b64 else ""
    st.markdown(
        f'<div class="topo">{img}<div><p class="marca">{esc(titulo)}</p><p class="sub">{esc(subtitulo)}</p></div>'
        f'<span class="selo">TAG Investimentos</span></div>',
        unsafe_allow_html=True,
    )


def badge(cls: str, texto: str) -> str:
    return f'<span class="badge {cls}">{esc(texto)}</span>'


def kpis_html(cartoes: list) -> str:
    """cartoes: lista de (classe, rótulo, valor, subtítulo[, menor]). 'menor' reduz a fonte para valores longos (R$)."""
    itens = []
    for c in cartoes:
        cls, rot, valor, sub_ = c[:4]
        menor = " menor" if len(c) > 4 and c[4] else ""
        itens.append(f'<div class="kpi {cls}"><div class="kpi-label">{esc(rot)}</div>'
                     f'<div class="kpi-valor{menor}">{esc(valor)}</div><div class="kpi-sub">{esc(sub_)}</div></div>')
    return '<div class="kpis">' + "".join(itens) + "</div>"


def aviso_html(cls: str, texto: str, titulo: str = "") -> str:
    cab = f"<b>{esc(titulo)}</b> " if titulo else ""
    return f'<div class="aviso {cls}">{cab}{esc(texto)}</div>'


def tabela_html(df: pd.DataFrame, formatos: dict = None, numericas=(), destaque=None, altura: int = 420,
                limite: int = 1500, ocultar=(), dicas: dict = None, quebrar: bool = False) -> str:
    """Tabela no padrão da marca (cabeçalho vinho, linhas zebradas).

    formatos   - {coluna: função(valor) -> HTML seguro}; sem função, o valor é escapado como texto
    numericas  - colunas alinhadas à direita
    destaque   - função(linha) -> classe CSS da linha ('linha-total', 'linha-geral', 'linha-div') ou ''
    ocultar    - colunas que não aparecem, mas continuam disponíveis para a função destaque
    dicas      - {coluna: texto} mostrado ao passar o mouse sobre o cabeçalho
    quebrar    - permite quebra de linha nas células (para textos longos, como legendas)
    """
    formatos = formatos or {}
    dicas = dicas or {}
    colunas = [c for c in df.columns if c not in ocultar]
    vis = df.head(limite)
    cab = "".join(f'<th class="{"num" if c in numericas else ""}"{f" title=" + chr(34) + esc(dicas[c]) + chr(34) if c in dicas else ""}>'
                  f'{esc(c)}</th>' for c in colunas)
    linhas = []
    for _, lin in vis.iterrows():
        classe = destaque(lin) if destaque else ""
        tds = []
        for c in colunas:
            v = lin[c]
            vazio = v is None or (not isinstance(v, str) and pd.isna(v))
            txt = "" if vazio else (formatos[c](v) if c in formatos else esc(v))
            tds.append(f'<td class="{"num" if c in numericas else ""}">{txt}</td>')
        linhas.append(f'<tr class="{classe}">{"".join(tds)}</tr>')
    nota = ""
    if len(df) > limite:
        nota = f'<div class="nota-tabela">Mostrando as primeiras {limite} de {len(df)} linhas. O Excel traz todas.</div>'
    return (f'<div class="tabela-wrap" style="max-height:{altura}px"><table class="tabela{" quebra" if quebrar else ""}">'
            f'<thead><tr>{cab}</tr></thead><tbody>{"".join(linhas)}</tbody></table></div>{nota}')


def passos_html(passos: list, formatar) -> str:
    """Passo a passo numerado: grupo, rótulo, valor e explicação. formatar(valor, tipo) -> texto."""
    linhas, n, grupo_atual = [], 0, None
    for p in passos:
        if p["grupo"] != grupo_atual:
            grupo_atual = p["grupo"]
            linhas.append(f'<tr class="grupo"><td colspan="4">{esc(grupo_atual)}</td></tr>')
        n += 1
        classe = "destaque" if p["destaque"] else ""
        linhas.append(f'<tr class="{classe}"><td class="n">{n}</td><td class="rot">{esc(p["rotulo"])}</td>'
                      f'<td class="val">{esc(formatar(p["valor"], p["tipo"]))}</td>'
                      f'<td class="expl">{esc(p["explicacao"])}</td></tr>')
    return f'<div class="tabela-wrap"><table class="passos"><tbody>{"".join(linhas)}</tbody></table></div>'


def estilizar_altair(chart):
    """Aplica o visual da marca a um gráfico Altair (use no gráfico final, já com as camadas)."""
    return (chart.configure_view(strokeWidth=0)
            .configure_axis(labelColor=COR["texto_suave"], titleColor=COR["texto_suave"], gridColor="#F0E8EB",
                            domainColor=COR["borda"], tickColor=COR["borda"], labelFont="Inter", titleFont="Inter")
            .configure(background="transparent"))
