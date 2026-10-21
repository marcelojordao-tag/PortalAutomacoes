"""Tema visual da Mesa de Boletas (identidade TAG Investimentos).

Uso no app principal:
    from tema import aplicar_tema, topo, CSS_ERROS, COR
    aplicar_tema()   # logo depois do st.set_page_config
    topo()           # no lugar do bloco <div class="topo">
    # e remova o CSS inline antigo e o CSS_ERROS antigo.
"""
import base64
from pathlib import Path

import streamlit as st

LOGO = Path(__file__).resolve().with_name("logo_tag.png")

# Paleta extraída do logo: vinho profundo + branco, com dourado discreto de apoio
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

CORES_PRIORIDADE = {"Baixa": "#2E8B57", "Média": "#C98A1B", "Alta": "#D2691E", "Crítica": "#8E1B2E"}

_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, .stApp, [class*="css"] { font-family: 'Inter', system-ui, sans-serif; }
#MainMenu, footer, header[data-testid="stHeader"] { visibility: hidden; height: 0; }
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
.topo img { height: 52px; width: auto; margin-left: auto; }
.topo .marca { font-family: 'Inter', system-ui, sans-serif; font-size: 1.8rem; font-weight: 700; letter-spacing: -.01em; margin: 0; color: #fff; line-height: 1.1; }
.topo .sub { margin: .2rem 0 0 0; font-size: .92rem; color: #E9CFD7; }
.topo .selo {
    margin-left: auto; font-size: .72rem; font-weight: 600; letter-spacing: .12em; text-transform: uppercase;
    color: __DOURADO__; border: 1px solid __DOURADO__; border-radius: 999px; padding: .3rem .9rem;
}

/* ---------- Abas (Streamlit 1.5x+: react-aria, sem baseweb) ---------- */
[role="tablist"] { gap: 6px !important; border: none !important; background: transparent !important; padding: 0 !important; box-shadow: none !important; }
[role="tab"] {
    height: 46px !important; padding: 0 1.6rem !important; margin: 0 !important;
    background: __VINHO_ESCURO__ !important; border: none !important; border-radius: 12px 12px 0 0 !important;
    opacity: 1 !important; transition: background .15s ease; cursor: pointer; display: flex !important; align-items: center;
}
[role="tab"] p, [role="tab"] * { color: #E9CFD7 !important; font-weight: 600 !important; font-size: .92rem !important; }
[role="tab"]:hover { background: __VINHO_MEDIO__ !important; }
[role="tab"][aria-selected="true"] { background: __CARD__ !important; box-shadow: inset 0 3px 0 __DOURADO__ !important; }
[role="tab"][aria-selected="true"] p, [role="tab"][aria-selected="true"] * { color: __VINHO__ !important; }
.react-aria-SelectionIndicator, [data-baseweb="tab-highlight"], [data-baseweb="tab-border"] { display: none !important; }
[role="tabpanel"] {
    background: __CARD__ !important; border: 1px solid __BORDA__; border-radius: 0 14px 14px 14px !important;
    padding: 1.4rem 1.6rem 1.8rem 1.6rem !important; box-shadow: 0 4px 14px rgba(74, 10, 30, .06);
}
/* ---------- Cartões e textos ---------- */
.card { background: __CARD__; border: 1px solid __BORDA__; border-radius: 12px; padding: 1.4rem 1.6rem; margin-top: 1rem; }
.card h3 { margin: 0 0 .3rem 0; color: __VINHO__; font-size: 1.1rem; }
.card p { margin: 0; color: __TEXTO_SUAVE__; }
.secao { color: __VINHO__; font-size: 1.05rem; font-weight: 700; margin: .2rem 0 .1rem 0; }
.secao-sub { color: __TEXTO_SUAVE__; font-size: .85rem; margin-bottom: .8rem; }

/* ---------- Botões ---------- */
.stButton > button, .stDownloadButton > button {
    background: __VINHO__; color: #fff; border: 1px solid __VINHO__;
    border-radius: 10px; padding: .55rem 1.4rem; font-weight: 600; transition: all .15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    background: __VINHO_MEDIO__; border-color: __VINHO_MEDIO__; color: #fff;
    box-shadow: 0 4px 12px rgba(106, 15, 43, .25);
}
.stButton > button:focus:not(:active) { border-color: __DOURADO__; color: #fff; }
.stButton > button:disabled { background: #D9CDD2; border-color: #D9CDD2; color: #fff; }

/* ---------- Upload e inputs ---------- */
[data-testid="stFileUploaderDropzone"] {
    background: __VINHO_CLARO__; border: 1.5px dashed #C9A3B1; border-radius: 12px; color: __TEXTO__;
}
[data-testid="stFileUploaderDropzone"] * { color: __TEXTO__; }
[data-testid="stFileUploaderDropzone"] button { color: #fff; }
.stTextInput input, .stTextArea textarea, [data-baseweb="select"] > div, .stNumberInput input {
    border-radius: 10px !important; border-color: __BORDA__ !important; background: #fff !important;
}

/* ---------- KPIs ---------- */
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: .9rem; margin: 0 0 1.1rem 0; }
.kpi {
    background: #fff; border: 1px solid __BORDA__; border-top: 4px solid __VINHO__;
    border-radius: 12px; padding: .95rem 1.25rem; box-shadow: 0 2px 8px rgba(74, 10, 30, .07);
}
.kpi.ok { border-top-color: __OK__; } .kpi.div { border-top-color: __DIV__; } .kpi.vazio { border-top-color: __ALERTA__; }
.kpi-label { color: __TEXTO_SUAVE__; font-size: .74rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
.kpi-valor { color: __TEXTO__; font-size: 2.1rem; font-weight: 700; line-height: 1.2; }
.kpi.ok .kpi-valor { color: __OK__; } .kpi.div .kpi-valor { color: __DIV__; }
.kpi-sub { color: __TEXTO_SUAVE__; font-size: .8rem; }

/* ---------- Etapas e progresso ---------- */
.etapas { display: flex; flex-wrap: wrap; gap: .6rem; margin: .2rem 0 1rem 0; }
.etapa {
    display: flex; align-items: center; gap: .5rem; padding: .4rem .95rem .4rem .5rem; border-radius: 999px;
    background: #fff; border: 1px solid __BORDA__; color: __TEXTO_SUAVE__; font-size: .85rem; font-weight: 600;
}
.etapa .n {
    width: 22px; height: 22px; border-radius: 50%; background: #E3D8DC; color: #fff;
    display: grid; place-items: center; font-size: .75rem; font-weight: 700;
}
.etapa.feita { border-color: __OK__; color: __OK__; }
.etapa.feita .n { background: __OK__; }
.etapa.atual { border-color: __VINHO__; color: __VINHO__; box-shadow: 0 2px 8px rgba(106, 15, 43, .15); }
.etapa.atual .n { background: __VINHO__; }
.progresso-txt { color: __TEXTO_SUAVE__; font-size: .82rem; font-weight: 600; margin: -.2rem 0 .3rem 0; }
.progresso { background: #EEE6E9; border-radius: 999px; height: 10px; overflow: hidden; margin: 0 0 1.1rem 0; }
.progresso > div { height: 100%; background: __OK__; border-radius: 999px; }

/* ---------- Resultado ---------- */
.res-titulo { display: flex; align-items: center; gap: .8rem; margin: .6rem 0 .9rem 0; }
.res-titulo h3 { margin: 0; color: __VINHO__; font-size: 1.3rem; font-weight: 700; }
.chip-modo { background: __VINHO_CLARO__; color: __VINHO__; border-radius: 999px; padding: .2rem .9rem; font-size: .78rem; font-weight: 600; }
.rotulo { color: __VINHO__; font-size: .74rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; margin: .4rem 0 .25rem 0; }

/* Radios em pílula */
div[data-testid="stRadio"] div[role="radiogroup"] { gap: .45rem; flex-wrap: wrap; }
div[data-testid="stRadio"] label[data-baseweb="radio"] {
    background: #fff; border: 1px solid __BORDA__; border-radius: 999px;
    padding: .35rem 1.1rem; margin: 0; cursor: pointer; transition: all .15s ease;
}
div[data-testid="stRadio"] label[data-baseweb="radio"] > div:first-child { display: none; }
div[data-testid="stRadio"] label[data-baseweb="radio"] p { color: __VINHO__; font-weight: 600; font-size: .88rem; }
div[data-testid="stRadio"] label[data-baseweb="radio"]:hover { border-color: __VINHO__; }
div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) { background: __VINHO__; border-color: __VINHO__; }
div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) p { color: #fff; }

/* Expansor */
[data-testid="stExpander"] { background: #fff; border: 1px solid __BORDA__; border-radius: 12px; margin-bottom: 1rem; }
[data-testid="stExpander"] details { border: none; }
[data-testid="stExpander"] summary p { color: __VINHO__; font-weight: 600; }

/* ---------- Tabela ---------- */
.tabela-wrap { background: #fff; border: 1px solid __BORDA__; border-radius: 12px; max-height: 540px; overflow: auto; box-shadow: 0 2px 8px rgba(74, 10, 30, .07); margin: .6rem 0 1rem 0; }
.tabela { width: 100%; border-collapse: separate; border-spacing: 0; font-size: .86rem; color: __TEXTO__; }
.tabela thead th {
    position: sticky; top: 0; z-index: 2; text-align: left; white-space: nowrap;
    background: __VINHO__; color: #fff; padding: .75rem .95rem; font-size: .76rem; font-weight: 600;
    letter-spacing: .05em; text-transform: uppercase; border-bottom: 2px solid __DOURADO__;
}
.tabela td { padding: .55rem .95rem; border-bottom: 1px solid #F0E8EB; white-space: nowrap; color: __TEXTO__; }
.tabela tbody tr:nth-child(even) td { background: #FBF8F9; }
.tabela tbody tr:hover td { background: __VINHO_CLARO__; }
.tabela .num { text-align: right; font-variant-numeric: tabular-nums; }
.tabela tr.linha-div td:first-child { box-shadow: inset 4px 0 0 __DIV__; }
.tabela .vazio { color: #BBA9B0; }
.tabela .motivo { color: __TEXTO_SUAVE__; font-size: .82rem; }
.badge { display: inline-block; padding: .18rem .7rem; border-radius: 999px; font-size: .72rem; font-weight: 700; letter-spacing: .04em; }
.badge.ok { background: #DDF1E5; color: #1D6B40; }
.badge.div { background: #F9DAD6; color: #9B2C22; }
.badge.neutro { background: #EEE6E9; color: #6B5560; }
.badge.ok-suave { background: #EAF6EE; color: #2E7D52; }
.badge.div-suave { background: #FBE7E4; color: #B5392E; }
.badge.warn { background: #FBEBC8; color: #8A5A00; }
.nota-tabela { color: __TEXTO_SUAVE__; font-size: .8rem; margin: -.4rem 0 .8rem 0; }

/* ---------- Compatibilidade Streamlit 1.5x+ e modo escuro do Windows ---------- */
/* O app é sempre claro; sem isso, os widgets seguem o tema escuro do sistema. */
html, .stApp, [data-testid="stDialog"] { color-scheme: light; }
[data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"] p,
[data-testid="stCheckbox"] label, [data-testid="stCheckbox"] label p { color: __TEXTO__; }
[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * { color: __TEXTO_SUAVE__; }

/* Botões (inclui os que têm tooltip, que ficam dentro de um wrapper) */
button[data-testid="stBaseButton-secondary"], button[data-testid="stBaseButton-primary"] {
    background: __VINHO__ !important; border: 1px solid __VINHO__ !important; border-radius: 10px !important;
    color: #fff !important; font-weight: 600 !important; transition: all .15s ease;
}
button[data-testid="stBaseButton-secondary"] *, button[data-testid="stBaseButton-primary"] * { color: #fff !important; }
button[data-testid="stBaseButton-secondary"]:hover, button[data-testid="stBaseButton-primary"]:hover {
    background: __VINHO_MEDIO__ !important; border-color: __VINHO_MEDIO__ !important;
    box-shadow: 0 4px 12px rgba(106, 15, 43, .25);
}
button[data-testid="stBaseButton-secondary"]:disabled, button[data-testid="stBaseButton-primary"]:disabled {
    background: #E3D8DC !important; border-color: #E3D8DC !important; box-shadow: none; cursor: not-allowed;
}
button[data-testid="stBaseButton-secondary"]:disabled * { color: #FFFFFF !important; }

/* Campos de seleção */
[data-testid="stSelectbox"] [role="group"] {
    background: #fff !important; border: 1px solid __BORDA__ !important; border-radius: 10px !important;
}
[data-testid="stSelectbox"] input { color: __TEXTO__ !important; -webkit-text-fill-color: __TEXTO__ !important; }
[data-testid="stSelectbox"] button svg { color: __VINHO__ !important; fill: __VINHO__ !important; }
[role="listbox"] { background: #fff !important; border: 1px solid __BORDA__; border-radius: 10px; }
[role="listbox"] [role="option"] { background: #fff !important; color: __TEXTO__ !important; }
[role="listbox"] [role="option"]:hover, [role="listbox"] [role="option"][aria-selected="true"],
[role="listbox"] [role="option"][data-focused="true"] { background: __VINHO_CLARO__ !important; }
[role="listbox"] [role="option"] * { color: __TEXTO__ !important; }

/* Seleção múltipla (ex.: "Este problema impacta") */
[data-testid="stMultiSelect"] [role="group"] {
    background: #fff !important; border: 1px solid __BORDA__ !important; border-radius: 10px !important;
}
[data-testid="stMultiSelect"] [role="group"] span {
    background: __VINHO_CLARO__ !important; color: __VINHO__ !important; border-radius: 6px; font-weight: 600;
}
[data-testid="stMultiSelect"] button { background: transparent !important; color: __VINHO__ !important; }
[data-testid="stMultiSelect"] button svg { color: __VINHO__ !important; }
input::placeholder, textarea::placeholder {
    color: __TEXTO_SUAVE__ !important; -webkit-text-fill-color: __TEXTO_SUAVE__ !important; opacity: 1 !important;
}

/* Campos de texto e número */
input, textarea { color: __TEXTO__ !important; -webkit-text-fill-color: __TEXTO__ !important; }
[data-testid="stTextInput"] input, [data-testid="stNumberInput"] input, textarea {
    background: #fff !important; border-radius: 10px;
}
[data-testid="stTextInput"] input { border: 1px solid __BORDA__ !important; }
[data-testid="stNumberInput"] button { background: #fff !important; }
[data-testid="stNumberInput"] button * { color: __VINHO__ !important; }

/* Radios em pílula */
div[data-testid="stRadioGroup"] { gap: .45rem; flex-wrap: wrap; }
label[data-testid="stRadioOption"] {
    background: #fff; border: 1px solid __BORDA__; border-radius: 999px;
    padding: .35rem 1.1rem; margin: 0; cursor: pointer; transition: all .15s ease;
}
label[data-testid="stRadioOption"] > div > div:first-child { display: none; }
label[data-testid="stRadioOption"] p { color: __VINHO__ !important; font-weight: 600; font-size: .88rem; }
label[data-testid="stRadioOption"]:hover { border-color: __VINHO__; }
label[data-testid="stRadioOption"][data-selected="true"] { background: __VINHO__; border-color: __VINHO__; }
label[data-testid="stRadioOption"][data-selected="true"] p { color: #fff !important; }

/* Janelas (dialogs) e alertas */
[data-testid="stDialog"] [role="dialog"] { background: #fff !important; color: __TEXTO__; border-radius: 16px; }
[data-testid="stDialog"] h2, [data-testid="stDialog"] h2 * { color: __VINHO__ !important; }
[data-testid="stExpander"] summary * { color: __VINHO__ !important; }
[data-testid="stExpander"] summary, [data-testid="stExpander"] details[open] > summary {
    background: #fff !important; border-radius: 12px;
}
[data-testid="stExpander"] summary:hover { background: __VINHO_CLARO__ !important; }
[data-testid="stExpander"] [data-testid="stExpanderDetails"] { background: #fff; border-radius: 0 0 12px 12px; }

/* Caixa de seleção marcada (o padrão do Streamlit é vermelho) */
[data-testid="stCheckbox"][data-selected="true"] label > div:first-of-type,
[data-testid="stCheckbox"] label[data-selected="true"] > div:first-of-type {
    background: __VINHO__ !important; border-color: __VINHO__ !important;
}

/* Arquivo enviado (chip que aparece depois do upload) */
[data-testid="stFileChip"] {
    background: #fff !important; border: 1px solid __BORDA__ !important; border-radius: 10px !important;
    box-shadow: 0 1px 4px rgba(74, 10, 30, .08);
}
[data-testid="stFileChip"] > div:first-child { background: __VINHO_CLARO__ !important; border-radius: 8px; }
[data-testid="stFileChip"] svg { color: __VINHO__ !important; fill: __VINHO__ !important; }
[data-testid="stFileChipName"] { color: __TEXTO__ !important; font-weight: 600; }
[data-testid="stFileChip"] div { color: __TEXTO_SUAVE__; }
[data-testid="stFileChip"] [data-testid="stFileChipName"] { color: __TEXTO__ !important; }
[data-testid="stFileChipDeleteBtn"] button, [data-testid="stFileChips"] button[data-testid="stBaseButton-minimal"] {
    background: transparent !important; border: none !important; box-shadow: none !important;
}
[data-testid="stFileChipDeleteBtn"] svg { color: __TEXTO_SUAVE__ !important; fill: __TEXTO_SUAVE__ !important; }
[data-testid="stFileChipDeleteBtn"] button:hover svg { color: __DIV__ !important; fill: __DIV__ !important; }
button[data-testid="stBaseButton-borderlessIcon"] { background: transparent !important; border: none !important; box-shadow: none !important; }
button[data-testid="stBaseButton-borderlessIcon"], button[data-testid="stBaseButton-borderlessIcon"] svg {
    color: __VINHO__ !important;
}
</style>
"""

CSS_ERROS = """
<style>
.erro-col-topo {
    display: flex; justify-content: space-between; align-items: center;
    background: __VINHO__; color: #fff; padding: .6rem .9rem; border-radius: 10px;
    font-weight: 700; font-size: .85rem; margin-bottom: .6rem; border-bottom: 2px solid __DOURADO__;
}
.erro-col-qtd { background: #fff; color: __VINHO__; border-radius: 999px; padding: .05rem .65rem; font-size: .78rem; }
.erro-vazio { border: 2px dashed #D8C3CB; color: __TEXTO_SUAVE__; border-radius: 10px; padding: 1rem; text-align: center; font-size: .85rem; }
[class*="st-key-cardErro_"] {
    background: #fff; border: 1px solid __BORDA__; border-left: 6px solid __VINHO__;
    border-radius: 12px; padding: .8rem .9rem; margin-bottom: .4rem; gap: .5rem;
    box-shadow: 0 2px 8px rgba(74, 10, 30, .08);
}
[class*="st-key-cardErro_baixa_"]   { border-left-color: #2E8B57; }
[class*="st-key-cardErro_media_"]   { border-left-color: #C98A1B; }
[class*="st-key-cardErro_alta_"]    { border-left-color: #D2691E; }
[class*="st-key-cardErro_critica_"] { border-left-color: #8E1B2E; }
[class*="st-key-cardErro_"] button[data-testid^="stBaseButton"] { padding: .2rem .4rem; min-height: 2rem; font-size: .8rem; width: 100%; }
[class*="st-key-cardErro_"] [data-testid="stTooltipHoverTarget"] { width: 100%; }
.erro-card-topo { display: flex; justify-content: space-between; align-items: center; margin-bottom: .3rem; }
.erro-id { color: __TEXTO_SUAVE__; font-size: .75rem; font-weight: 700; }
.erro-titulo { color: __TEXTO__; font-weight: 700; font-size: .95rem; line-height: 1.25; }
.erro-meta { color: __TEXTO_SUAVE__; font-size: .78rem; margin-top: .35rem; margin-bottom: .85rem; }
.erro-cat { display: flex; flex-wrap: wrap; gap: .35rem; margin-top: .4rem; }
.badge.cat { background: __VINHO_CLARO__; color: __VINHO__; }
.badge.causa { background: #EEE6E9; color: #6B5560; }
.badge.prio-baixa   { background: #DDF1E5; color: #1D6B40; }
.badge.prio-media   { background: #FBEBC8; color: #8A5A00; }
.badge.prio-alta    { background: #F9DFC9; color: #9A4A0B; }
.badge.prio-critica { background: #F4D3D8; color: #8E1B2E; }
</style>
"""


def _pintar(css: str) -> str:
    for chave, valor in COR.items():
        css = css.replace(f"__{chave.upper()}__", valor)
    return css


CSS_ERROS = _pintar(CSS_ERROS)


_CSS_PORTAL = """
<style>
/* ---------- Menu lateral (vinho, texto claro) ---------- */
section[data-testid="stSidebar"], [data-testid="stSidebar"] > div {
    background: linear-gradient(180deg, __VINHO_ESCURO__ 0%, __VINHO__ 55%, __VINHO_MEDIO__ 100%) !important;
}
section[data-testid="stSidebar"] { border-right: 3px solid __DOURADO__; box-shadow: 4px 0 18px rgba(74, 10, 30, .25); }
[data-testid="stSidebar"] * { color: #E9CFD7; }
[data-testid="stSidebar"] { font-family: 'Inter', system-ui, sans-serif; }
/* ícones do Streamlit são fonte de ligaturas: não podem herdar a Inter */
[data-testid="stIconMaterial"], [data-testid="stSidebar"] [class*="material" i] { font-family: "Material Symbols Rounded", "Material Symbols Outlined" !important; }
[data-testid="stSidebarHeader"] {
    background: transparent !important; padding: 1.4rem 1.3rem 1rem; height: auto; margin-bottom: .4rem;
    border-bottom: 1px solid rgba(184, 146, 74, .55);
}
[data-testid="stSidebarHeader"] img { height: 44px; width: auto; max-width: 100%; }
/* Botão de recolher/abrir o menu: o tema esconde o header, então o botão precisa ser forçado a aparecer */
[data-testid="stSidebarCollapseButton"], [data-testid="stSidebarCollapseButton"] button {
    visibility: visible !important; opacity: 1 !important;
}
[data-testid="stSidebarCollapseButton"] button, [data-testid="stSidebarCollapseButton"] button:hover {
    background: rgba(255, 255, 255, .14); border-radius: 10px; color: #fff;
}
[data-testid="stExpandSidebarButton"], [data-testid="stSidebarCollapsedControl"] {
    visibility: visible !important; opacity: 1 !important; position: fixed; top: .9rem; left: .9rem; z-index: 999999;
    display: flex !important;
}
[data-testid="stExpandSidebarButton"], [data-testid="stExpandSidebarButton"] button, [data-testid="stSidebarCollapsedControl"] button {
    background: __VINHO__ !important; color: #fff !important; border-radius: 10px; border-bottom: 3px solid __DOURADO__;
    box-shadow: 0 4px 14px rgba(74, 10, 30, .35);
}
[data-testid="stExpandSidebarButton"] *, [data-testid="stSidebarCollapsedControl"] * { color: #fff !important; fill: #fff !important; }
[data-testid="stExpandSidebarButton"]:hover, [data-testid="stExpandSidebarButton"] button:hover { background: __VINHO_MEDIO__ !important; }
/* Ícone próprio (independe da fonte de ícones do Streamlit, que pode falhar e mostrar o texto da ligatura) */
[data-testid="stSidebarCollapseButton"] button *, [data-testid="stExpandSidebarButton"] *, [data-testid="stSidebarCollapsedControl"] button * {
    font-size: 0 !important; width: 0; height: 0; overflow: hidden; display: block; opacity: 0;
}
[data-testid="stSidebarCollapseButton"] button, [data-testid="stExpandSidebarButton"], [data-testid="stSidebarCollapsedControl"] button {
    width: 34px; height: 34px; min-width: 34px; padding: 0; display: flex; align-items: center; justify-content: center;
    cursor: pointer;
}
[data-testid="stSidebarCollapseButton"] button { position: relative; }
[data-testid="stExpandSidebarButton"], [data-testid="stSidebarCollapsedControl"] {
    position: fixed !important; top: 1rem !important; left: 1rem !important; margin: 0 !important; z-index: 999999;
}
[data-testid="stSidebarCollapseButton"] button::after, [data-testid="stExpandSidebarButton"]::after, [data-testid="stSidebarCollapsedControl"] button::after {
    font-family: 'Inter', system-ui, sans-serif; font-size: 1.15rem; font-weight: 700; line-height: 1; color: #fff;
}
[data-testid="stSidebarCollapseButton"] button::after { content: "\\00AB"; }
[data-testid="stExpandSidebarButton"]::after, [data-testid="stSidebarCollapsedControl"] button::after { content: "\\00BB"; }
[data-testid="stSidebarHeader"] button svg, [data-testid="stSidebarCollapseButton"] svg { color: #fff; fill: #fff; }
[data-testid="stSidebarNavSeparator"], [data-testid="stSidebar"] hr { display: none; }
[data-testid="stSidebarNav"] ul { padding: 0 .8rem; gap: .2rem; }
[data-testid="stSidebarNav"] [data-testid="stNavSectionHeader"], [data-testid="stSidebarNav"] header {
    color: __DOURADO__ !important; font-size: .68rem; font-weight: 700; letter-spacing: .14em; text-transform: uppercase;
    padding: 1.1rem 1rem .35rem;
}
[data-testid="stSidebarNav"] a {
    border-radius: 10px; font-weight: 500; padding: .6rem 1rem; border-left: 3px solid transparent;
    transition: background .15s ease, color .15s ease, padding-left .15s ease;
}
[data-testid="stSidebarNav"] a:hover { background: rgba(255, 255, 255, .10); padding-left: 1.25rem; }
[data-testid="stSidebarNav"] a:hover * { color: #fff; }
[data-testid="stSidebarNav"] a[aria-current="page"] {
    background: rgba(255, 255, 255, .16); border-left: 3px solid __DOURADO__; font-weight: 700; box-shadow: 0 2px 10px rgba(0, 0, 0, .18);
}
[data-testid="stSidebarNav"] a[aria-current="page"] * { color: #fff; font-weight: 700; }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { padding: 0 1.3rem; opacity: .75; font-size: .75rem; letter-spacing: .04em; }

/* ---------- Home: cartões das ferramentas ---------- */
[class*="st-key-cartao_"] {
    position: relative; background: __CARD__; border: 1px solid __BORDA__; border-top: 4px solid __VINHO__;
    border-radius: 12px; padding: 1.3rem 1.5rem 1.1rem; box-shadow: 0 2px 8px rgba(74, 10, 30, .07);
    transition: box-shadow .15s ease, border-color .15s ease; gap: .6rem; min-height: 11.5rem;
}
[class*="st-key-cartao_"]:hover { box-shadow: 0 6px 20px rgba(74, 10, 30, .18); border-color: __DOURADO__; border-top-color: __DOURADO__; }
[class*="st-key-cartao_"] .cartao-sigla {
    width: 46px; height: 46px; border-radius: 12px; display: flex; align-items: center; justify-content: center;
    background: linear-gradient(135deg, __VINHO_ESCURO__, __VINHO_MEDIO__); color: #fff; font-weight: 700;
    font-size: 1rem; letter-spacing: .04em; border-bottom: 3px solid __DOURADO__; margin-bottom: .8rem;
}
[class*="st-key-cartao_"] .cartao-titulo { color: __VINHO__; font-size: 1.15rem; font-weight: 700; margin: 0 0 .3rem; }
[class*="st-key-cartao_"] .cartao-desc { color: __TEXTO_SUAVE__; font-size: .92rem; margin: .5rem 0 .3rem !important; }
[class*="st-key-cartao_"] [data-testid="stPageLink"] { margin-top: .5rem; }
[class*="st-key-cartao_"] a[data-testid="stPageLink-NavLink"] {
    background: __VINHO__; color: #fff; border-radius: 10px; font-weight: 600; width: fit-content; padding: .45rem 1.1rem;
}
[class*="st-key-cartao_"] a[data-testid="stPageLink-NavLink"]:hover { background: __VINHO_MEDIO__; }
[class*="st-key-cartao_"] a[data-testid="stPageLink-NavLink"] p { color: #fff; font-weight: 600; }
[class*="st-key-cartao_"] a[data-testid="stPageLink-NavLink"]::after { content: ""; position: absolute; inset: 0; }
</style>
"""


def aplicar_tema() -> None:
    st.markdown(_pintar(_CSS), unsafe_allow_html=True)
    st.markdown(_pintar(_CSS_PORTAL), unsafe_allow_html=True)


def _logo_base64() -> str:
    try:
        return base64.b64encode(LOGO.read_bytes()).decode()
    except OSError:
        return ""


def topo(titulo: str = "Mesa de Boletas", subtitulo: str = "Conferência, liquidação e controle de erros em um só lugar") -> None:
    b64 = _logo_base64()
    img = f'<img src="data:image/png;base64,{b64}" alt="TAG Investimentos">' if b64 else ""
    st.markdown(
        f'<div class="topo"><div><p class="marca">{titulo}</p><p class="sub">{subtitulo}</p></div>'
        f'{img or "<span class=selo>TAG Investimentos</span>"}</div>',
        unsafe_allow_html=True,
    )


def estilizar_figura(fig, titulo: str = "") -> None:
    """Aplica o visual da marca a um gráfico Plotly."""
    fig.update_layout(
        title=dict(text=titulo or fig.layout.title.text, font=dict(size=16, color=COR["vinho"]), x=0.01),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=COR["texto"]),
        margin=dict(t=70, b=20, l=20, r=20),
    )
    fig.update_xaxes(showgrid=False, linecolor=COR["borda"])
    fig.update_yaxes(gridcolor="#F0E8EB", linecolor=COR["borda"])
