import json
from pathlib import Path

import pandas as pd
import streamlit as st

from tema import aplicar_tema, topo

RAIZ = Path(__file__).resolve().parent
LOGO = RAIZ / "logo_tag.png"
PLANILHA_EXTRATOS = RAIZ / "apps" / "extratos" / "Template_Fundos_Ativos.xlsx"
ARQUIVO_ERROS = RAIZ / "apps" / "boletador" / "controle_erros.json"

st.set_page_config(
    page_title="Portal | TAG Investimentos",
    page_icon=str(LOGO),
    layout="wide",
    initial_sidebar_state="expanded",
)


def selo_extratos() -> str:
    try:
        n = len(pd.read_excel(PLANILHA_EXTRATOS, dtype=str).dropna(how="all"))
    except (OSError, ValueError):
        return '<span class="badge warn">PLANILHA NÃO ENCONTRADA</span>'
    return f'<span class="badge ok">{n} LINHAS NA PLANILHA</span>'


def selo_boletas() -> str:
    try:
        cards = json.loads(ARQUIVO_ERROS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return '<span class="badge neutro">SEM DADOS</span>'
    abertos = [c for c in cards if c.get("coluna") != "Solucionados"]
    criticos = [c for c in abertos if str(c.get("prioridade", "")).startswith("Cr")]
    if criticos:
        return f'<span class="badge div">{len(criticos)} CRÍTICO(S) EM ABERTO</span>'
    if abertos:
        return f'<span class="badge warn">{len(abertos)} PROBLEMA(S) EM ABERTO</span>'
    return '<span class="badge ok">SEM PENDÊNCIAS</span>'


pg_home = st.Page(lambda: pagina_home(), title="Home", default=True)
pg_extratos = st.Page("apps/extratos/app.py", title="Solicitação de extratos", url_path="extratos")
pg_boletador = st.Page("apps/boletador/app.py", title="Mesa de Boletas", url_path="boletas")
pg_britech = st.Page("apps/britech/app.py", title="Rentabilidade e Cotas", url_path="rentabilidade")
pg_bradesco = st.Page("apps/bradesco_miami/app.py", title="Transactions Offshore", url_path="bradesco-miami")

# Para incluir uma nova ferramenta: crie a st.Page acima e acrescente uma linha aqui e em st.navigation.
FERRAMENTAS = [
    {
        "chave": "extratos", "sigla": "SE", "titulo": "Solicitação de extratos", "pagina": pg_extratos,
        "descricao": "Envio de e-mails, acompanhamento e conciliação dos extratos de cotistas.", "selo": selo_extratos,
    },
    {
        "chave": "boletas", "sigla": "MB", "titulo": "Mesa de Boletas", "pagina": pg_boletador,
        "descricao": "Conciliação de boletas, controle de erros e importação da base.", "selo": selo_boletas,
    },
    {
        "chave": "rentabilidade", "sigla": "RC", "titulo": "Rentabilidade e Cotas", "pagina": pg_britech,
        "descricao": "Conferência do PL com a cota e monitor da rentabilidade diária (API Britech).",
        "selo": lambda: '<span class="badge neutro">API BRITECH</span>',
    },
    {
        "chave": "bradesco", "sigla": "BM", "titulo": "Transactions Offshore", "pagina": pg_bradesco,
        "descricao": "Tratamento dos arquivos de transações (Bradesco Miami e BTG US): filtros, janela D-2 e cliente.",
        "selo": lambda: '<span class="badge neutro">EXCEL</span>',
    },
]
COLUNAS = 3


def pagina_home():
    topo("Portal TAG", "Todas as ferramentas em um só lugar")
    st.markdown(
        '<p class="secao">Ferramentas</p><p class="secao-sub">Escolha uma ferramenta para começar</p>',
        unsafe_allow_html=True,
    )
    for i in range(0, len(FERRAMENTAS), COLUNAS):
        colunas = st.columns(COLUNAS)
        for col, f in zip(colunas, FERRAMENTAS[i : i + COLUNAS]):
            with col, st.container(key=f"cartao_{f['chave']}"):
                st.markdown(
                    f'<div class="cartao-sigla">{f["sigla"]}</div>'
                    f'<p class="cartao-titulo">{f["titulo"]}</p>{f["selo"]()}'
                    f'<p class="cartao-desc">{f["descricao"]}</p>',
                    unsafe_allow_html=True,
                )
                st.page_link(f["pagina"], label="Abrir")


aplicar_tema()  # uma vez, vale para todas as páginas
st.logo(str(LOGO), size="large")
pg = st.navigation({"Início": [pg_home], "Ferramentas": [f["pagina"] for f in FERRAMENTAS]})
st.sidebar.caption("Portal TAG Investimentos")
pg.run()
