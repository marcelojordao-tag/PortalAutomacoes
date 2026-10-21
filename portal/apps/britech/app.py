#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rentabilidade e Cotas (Streamlit)

Duas abas sobre as mesmas carteiras:
  - Conferência de Rentabilidade: o PL final bate com a rentabilidade da cota? Quando não bate, de onde vem a
    diferença (aplicações e resgates de cotas). O passo a passo abre em um pop-up.
  - Monitor de Cotas: o retorno diário da cota está Acima, Abaixo ou Dentro do normal da carteira (banda de média
    ± desvios com peso exponencial, em 4 janelas)? Gráfico por carteira em pop-up e Excel com gráficos.

Como usar:
    streamlit run app.py
    Em "Buscar carteiras na API", escolha os IDs e o período e clique em "Buscar na API": o app baixa os dados
    (mesma busca do teste.py), incluindo o histórico de cotas do monitor (por padrão, 2 anos). Também dá para rodar
    `python teste.py` e só abrir o app.

A busca na API fica no teste.py, a conferência no converter.py, o monitor no monitor.py (a mesma lógica gera o Excel e
este app). O visual (cores, cartões, tabelas) fica no tema.py, no padrão da Mesa de Boletas.
"""
from __future__ import annotations

import io
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

# O portal roda este arquivo via st.navigation: garante a pasta do app (converter, monitor, teste) e a raiz (tema) no sys.path.
for _pasta in (Path(__file__).resolve().parent, Path(__file__).resolve().parents[2]):
    if str(_pasta) not in sys.path:
        sys.path.insert(0, str(_pasta))
import converter as cv  # noqa: E402
import monitor as mon  # noqa: E402  (monitor de cotas: Acima/Abaixo/Dentro da banda)
import teste as tb  # noqa: E402  (busca na API Britech; o mesmo código do teste.py)
from tema import topo  # noqa: E402  (cabeçalho padrão do portal)
from tema_britech import (COR, aplicar_tema, aviso_html, badge, esc, estilizar_altair, kpis_html,  # noqa: E402
                          passos_html, tabela_html)

PASTA = Path(__file__).resolve().parent
ARQUIVO_PADRAO = PASTA / cv.ARQUIVO_PADRAO
CHAVE_ARQUIVO, CHAVE_EXTRATOS, CHAVE_HISTORICO = "arquivo_api", "pasta_extratos", "arquivo_historico"

# set_page_config fica no portal (home.py)
aplicar_tema()

# Como cada status aparece: cor do KPI/aviso, cor do selo, ícone (o ícone garante que a cor não é a única pista) e
# o nome usado no CSS do cartão.
STATUS_UI = {
    cv.STATUS_OK: {"cls": "ok", "selo": "ok", "icone": "✓", "slug": "ok"},
    cv.STATUS_EXPLICADA: {"cls": "alerta", "selo": "warn", "icone": "⇄", "slug": "explicada"},
    cv.STATUS_DIVERGENTE: {"cls": "div", "selo": "div", "icone": "✕", "slug": "div"},
}
SEVERIDADE = {cv.STATUS_DIVERGENTE: 0, cv.STATUS_EXPLICADA: 1, cv.STATUS_OK: 2}


def selo(status: str) -> str:
    ui = STATUS_UI[status]
    return badge(ui["selo"], f"{ui['icone']} {status}")


# ----------------------------------------------------------------------------
# Dados
# ----------------------------------------------------------------------------
def versao_codigo() -> float:
    """Data do converter.py: se a lógica mudar, o cache é refeito."""
    return Path(cv.__file__).stat().st_mtime


@st.cache_data(show_spinner="Analisando as carteiras…")
def carregar(caminho: str, mtime: float, extratos: str, codigo: float):
    """mtime e codigo entram na chave do cache: arquivo novo ou lógica nova recarregam os dados sozinhos."""
    return cv.analisar_todas(caminho, pasta_extratos=extratos)


@st.cache_data(show_spinner=False)
def excel_bytes(caminho: str, mtime: float, extratos: str, codigo: float) -> bytes:
    analises, _ = carregar(caminho, mtime, extratos, codigo)
    buf = io.BytesIO()
    cv.salvar_excel_multi(analises, buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Formatação
# ----------------------------------------------------------------------------
def fmt_num(v, casas=2) -> str:
    if v is None or pd.isna(v):
        return "—"
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_valor(v, tipo) -> str:
    if v is None:
        return "—"
    return {"brl": cv.fmt_brl, "pct": cv.fmt_pct, "qtd": lambda x: fmt_num(x, 4),
            "cota": lambda x: fmt_num(x, 8), "txt": str}[tipo](v)


def txt_brl(v) -> str:
    return esc(cv.fmt_brl(v))


def txt_num(casas=2):
    return lambda v: esc(fmt_num(v, casas))


DICAS_COMPARATIVA = {
    "Carteira": "ID da carteira no Britech.",
    "Nome": "Nome do cliente ou fundo, lido do cabeçalho do extrato de conta corrente.",
    "Status": (f"Sem divergência: a diferença é de até R$ {cv.TOLERANCIA_REAIS:.2f} (arredondamento). "
               f"Explicada por cotistas: há aplicação ou resgate de cotas e o resíduo é de até "
               f"{cv.TOLERANCIA_EXPLICADA_PCT * 100:.2f}% do PL inicial. Divergente: o resto, ou seja, sobra diferença sem explicação."),
    "PL inicial": "Valor da carteira no dia-base (soma do valor final dos ativos nesse dia).",
    "PL final real": "Valor da carteira no último dia do período (soma do valor final dos ativos nesse dia).",
    "Esperado pela cota": "PL inicial × (1 + rentabilidade da cota). É o PL que a carteira teria se só houvesse rentabilidade, "
                          "sem aplicação nem resgate de cotas.",
    "Diferença": "PL final real − Esperado pela cota. Negativo: saiu dinheiro da carteira por fora da rentabilidade. "
                 "Positivo: entrou dinheiro.",
    "Efeito dos cotistas": "Cotas aplicadas ou resgatadas no período × cota do último dia, em R$. Negativo: resgates. "
                           "Fica vazio se o histórico de cotas não foi baixado.",
    "Resíduo": "Diferença − Efeito dos cotistas: o que sobra sem explicação. Perto de zero: a diferença vem dos cotistas.",
}


def resumo_curto(a: dict) -> str:
    r, cot = a["resumo"], a["cot"]
    dif = r["dif_pl"]
    if a["status"] == cv.STATUS_OK:
        return f"A cota explica o PL (diferença de {cv.fmt_brl(dif)})."
    verbo = "saíram" if dif < 0 else "entraram"
    if a["status"] == cv.STATUS_EXPLICADA:
        return f"{cv.fmt_brl(abs(dif))} {verbo} por movimentação de cotas ({cot['n_operacoes']} operações)."
    return f"Diferença de {cv.fmt_brl(dif)} sem explicação nos dados disponíveis."


def html(texto: str) -> None:
    st.markdown(texto, unsafe_allow_html=True)


def grafico_evolucao(h: pd.DataFrame, coluna: str, titulo: str, datas_ops: pd.DataFrame) -> alt.Chart:
    """Linha diária com eixo ajustado (sem forçar zero) e marcas nos dias de aplicação/resgate."""
    dados = h[["Data", coluna]].rename(columns={coluna: titulo})
    linha = alt.Chart(dados).mark_line(point=alt.OverlayMarkDef(filled=True, size=45, color=COR["vinho"]),
                                       color=COR["vinho"], strokeWidth=2.5).encode(
        x=alt.X("Data:T", title=None, axis=alt.Axis(format="%d/%m", labelAngle=0)),
        y=alt.Y(f"{titulo}:Q", title=None, scale=alt.Scale(zero=False), axis=alt.Axis(format=",.0f")),
        tooltip=[alt.Tooltip("Data:T", title="Data", format="%d/%m/%Y"), alt.Tooltip(f"{titulo}:Q", title=titulo, format=",.2f")])
    regras = alt.Chart(datas_ops).mark_rule(strokeDash=[4, 4], color=COR["dourado"], strokeWidth=2).encode(x="Data:T")
    return estilizar_altair((linha + regras).properties(height=210))


def ranking_contribuicao(a: dict, quantos: int = 5) -> None:
    """Maiores altas e baixas da carteira. Contribuição = rendimento do ativo ÷ PL inicial total (em p.p.)."""
    cdf = a["contribuicoes"]
    html('<div class="rotulo">Quem mais contribuiu para a rentabilidade</div>')
    if cdf.empty:
        html(aviso_html("alerta", "Nenhum ativo com rendimento no período."))
        return

    ordem = st.segmented_control("Ordenar por", ["Contribuição (bps)", "% do ativo"], default="Contribuição (bps)",
                                 key=f"ord_{a['id']}", label_visibility="collapsed") or "Contribuição (bps)"
    col = "% Rent" if ordem == "% do ativo" else "Contribuição"
    maximo = float(cdf["Contribuição"].abs().max())
    pl_total = a["resumo"]["pl_inicial"]

    # Deixa explícito o que cada opção significa
    if col == "Contribuição":
        html(aviso_html("", "Ranking pelo quanto cada ativo puxou a rentabilidade da carteira: rendimento do ativo ÷ PL inicial "
                            "da carteira. Considera o tamanho da posição, e a soma de todos os ativos fecha com a rentabilidade "
                            "dos ativos.", "Ordenado por contribuição (bps):"))
    else:
        html(aviso_html("alerta", "Ranking pela rentabilidade do próprio ativo: rendimento ÷ PL inicial do ativo. NÃO considera o "
                                  "tamanho da posição: um ativo pequeno pode aparecer no topo e ter pesado pouco na carteira "
                                  "(veja a coluna 'Peso na carteira'). Para o peso real no resultado, use a contribuição em bps.",
                        "Ordenado por % do ativo:"))

    def contrib(v):
        cls = "pos" if v > 0 else "neg"
        largura = min(100.0, abs(v) / maximo * 100) * 0.6 if maximo else 0
        return (f'<span class="{cls}">{esc(f"{v * 10000:+.2f}".replace(".", ","))} bps</span>'
                f'<span class="barra {cls}" style="width:{largura:.0f}px"></span>')

    def pct_ativo(v):
        return f'<span class="{"pos" if v > 0 else "neg"}">{esc(cv.fmt_pct(v, 2))}</span>'

    def tabela(df):
        vis = df[["Ativo", "Contribuição", "Rendimento", "% Rent"]].copy()
        vis["Peso na carteira"] = df["PL Inicial"] / pl_total if pl_total else 0.0
        vis = vis.rename(columns={"Contribuição": "Contribuição (bps)", "Rendimento": "Rendimento (R$)", "% Rent": "% do ativo"})
        vis = vis[["Ativo", "Peso na carteira", "% do ativo", "Contribuição (bps)", "Rendimento (R$)"]]
        fm = {"Ativo": lambda v: f'<span class="trunc" title="{esc(v)}">{esc(v)}</span>',
              "Peso na carteira": lambda v: esc(cv.fmt_pct(v, 2)), "% do ativo": pct_ativo,
              "Contribuição (bps)": contrib, "Rendimento (R$)": txt_brl}
        return tabela_html(vis, numericas={"Peso na carteira", "% do ativo", "Contribuição (bps)", "Rendimento (R$)"},
                           formatos=fm, altura=320)

    altas = cdf[cdf[col] > 0].sort_values(col, ascending=False).head(quantos)
    baixas = cdf[cdf[col] < 0].sort_values(col).head(quantos)
    # Uma tabela sobre a outra: lado a lado ficam apertadas dentro do pop-up
    html('<div class="rank-titulo pos">▲ Maiores altas</div>')
    html(tabela(altas) if len(altas) else '<div class="nota-tabela">Nenhum ativo com alta.</div>')
    html('<div class="rank-titulo neg">▼ Maiores baixas</div>')
    html(tabela(baixas) if len(baixas) else '<div class="nota-tabela">Nenhum ativo com baixa.</div>')
    soma_bps = f'{a["resumo"]["simples"] * 10000:.2f}'.replace(".", ",")
    html(f'<div class="nota-tabela">Contribuição = rendimento do ativo ÷ PL inicial da carteira ({cv.fmt_brl(a["resumo"]["pl_inicial"])}), '
         f'em bps (1 bps = 0,01%). A soma de todos os ativos é {soma_bps} bps ({esc(cv.fmt_pct(a["resumo"]["simples"]))}), '
         f'a rentabilidade dos ativos. Um ativo pequeno pode ter % alto e contribuir pouco.</div>')

    with st.expander("Da soma dos ativos até a cota: caixa, despesas e outros"):
        etapas = pd.DataFrame(cv.degraus_contribuicao(a), columns=["Etapa", "Tipo", "R$", "frac"])
        etapas["bps"] = etapas["frac"] * 10000
        etapas = etapas[["Etapa", "R$", "bps"]].rename(columns={"R$": "Valor (R$)", "bps": "Contribuição (bps)"})
        sinal = lambda v: f'<span class="{"pos" if v > 0 else "neg" if v < 0 else ""}">{esc(f"{v:+.2f}".replace(".", ","))}</span>'
        html(tabela_html(etapas, numericas={"Valor (R$)", "Contribuição (bps)"},
                         formatos={"Etapa": lambda v: esc(v), "Valor (R$)": txt_brl, "Contribuição (bps)": sinal}, altura=360))
        html('<div class="nota-tabela">A soma das contribuições é a rentabilidade simples dos ativos. A diferença de ponderação '
             'surge quando o PL varia no período (resgates/aplicações): o rendimento de cada dia pesa sobre o PL daquele dia, '
             'não sobre o PL inicial. O valor em R$ dessa linha é só bps × PL inicial, não um valor de caixa. Despesas, provisões e outros levam o resultado até a cota. '
             'Esta seção também sai na aba Contribuicoes do Excel.</div>')


# ----------------------------------------------------------------------------
# Pop-up com o passo a passo
# ----------------------------------------------------------------------------
@st.dialog("Passo a passo da conferência", width="large")
def detalhe(a: dict):
    r, cot, ui = a["resumo"], a["cot"], STATUS_UI[a["status"]]
    ini, fim = a["ativos"].attrs["inicio"], a["ativos"].attrs["fim"]
    html(f'<div class="dlg-nome">{esc(a["nome"])}</div>'
         f'<div class="dlg-meta">Carteira {esc(a["id"])} · período {ini:%d/%m/%Y} a {fim:%d/%m/%Y}</div>'
         f'{selo(a["status"])}<div style="height:.7rem"></div>')
    html(aviso_html(ui["cls"], a["mensagem"]))

    explicado = a["status"] != cv.STATUS_DIVERGENTE
    html(kpis_html([
        ("div" if abs(r["dif_pl"]) > cv.TOLERANCIA_REAIS else "ok", "Diferença (real − esperado)", cv.fmt_brl(r["dif_pl"]),
         "PL real menos PL esperado pela cota", True),
        ("alerta", "Explicada pelos cotistas", cv.fmt_brl(r["efeito_cotas"]) if r["efeito_cotas"] is not None else "—",
         "aplicações e resgates de cotas", True),
        ("ok" if explicado else "div", "Resíduo (sem explicação)",
         cv.fmt_brl(r["residuo_cotistas"]) if r["residuo_cotistas"] is not None else "—", "o que sobra depois dos cotistas", True),
    ]))

    t_passos, t_cotistas, t_evolucao, t_ponte, t_ativos = st.tabs(
        ["Passo a passo", "Movimentação de cotistas", "Evolução diária", "Conferência e ponte", "Ativos"])

    with t_passos:
        html('<div class="secao-sub">Pergunta: se a carteira só tivesse rentabilidade (sem aplicar nem resgatar cotas), '
             'o PL final seria igual ao PL final real?</div>')
        html(passos_html(a["passos"], fmt_valor))

    with t_cotistas:
        if not cot["disponivel"]:
            html(aviso_html("alerta", "Histórico de cotas não encontrado. Rode o teste.py novamente."))
        elif cot["n_operacoes"] == 0:
            html(aviso_html("ok", f"Nenhuma aplicação nem resgate de cotas no período. A quantidade de cotas ficou em "
                                  f"{fmt_num(cot['qtd_fim'], 4)}."))
        else:
            html(kpis_html([
                ("ok", "Aplicações", cv.fmt_brl(cot["aplicacoes"]), "valor bruto no período", True),
                ("div", "Resgates", cv.fmt_brl(cot["resgates"]), "valor bruto no período", True),
                ("total", "Cotas no início", fmt_num(cot["qtd_ini"], 2), "dia-base", True),
                ("total", "Cotas no fim", fmt_num(cot["qtd_fim"], 2), f"variação de {fmt_num(cot['qtd_fim'] - cot['qtd_ini'], 2)}", True),
            ]))
            if cot["qtd_confere"] is False:
                html(aviso_html("alerta", "A variação da quantidade de cotas não bate com a soma das operações."))
            ops = cot["operacoes"].copy()
            ops["Efeito"] = ops["Sinal"].map({1: "Entra", -1: "Sai", 0: "Não classificado"})
            html(tabela_html(
                ops.drop(columns=["Sinal"]), numericas={"Quantidade", "Valor bruto", "Valor líquido", "IR"}, altura=360,
                formatos={"Data": lambda v: esc(f"{v:%d/%m/%Y}"), "Quantidade": txt_num(4), "Valor bruto": txt_num(2),
                          "Valor líquido": txt_num(2), "IR": txt_num(2),
                          "Efeito": lambda v: badge({"Entra": "ok-suave", "Sai": "div-suave"}.get(v, "neutro"), v)}))
            html('<div class="nota-tabela">Fonte: API Fundo/OperacaoCotistaAnalitico_PorCarteiraPeriodo.</div>')

    with t_evolucao:
        h = cot["historico"]
        if h is None or h.empty:
            html(aviso_html("alerta", "Sem histórico diário de cotas."))
        else:
            datas_ops = pd.DataFrame({"Data": pd.to_datetime(cot["operacoes"]["Data"]).drop_duplicates()})
            html('<div class="rotulo">PL de fechamento (R$)</div>')
            st.altair_chart(grafico_evolucao(h, "PLFechamento", "PL (R$)", datas_ops), theme=None, width="stretch")
            html('<div class="rotulo">Quantidade de cotas</div>')
            st.altair_chart(grafico_evolucao(h, "QuantidadeFechamento", "Cotas", datas_ops), theme=None, width="stretch")
            html('<div class="nota-tabela">Linhas douradas tracejadas: dias com aplicação ou resgate de cotas. Uma queda na '
                 'quantidade de cotas indica resgate. Fonte: API Fundo/BuscaHistoricoCotaDia.</div>')

    with t_ponte:
        html('<div class="rotulo">Cota × ativos</div>')
        html(kpis_html([
            ("total", "Rentabilidade dos ativos", cv.fmt_pct(r["simples"]), "rendimento ÷ PL inicial"),
            ("total", "Rentabilidade da cota", cv.fmt_pct(r["carteira"]), "RentabilidadeCotaBruta"),
            ("alerta", "Diferença em R$", cv.fmt_brl(r["diferenca_reais"]), "(cota − ativos) × PL inicial", True),
        ]))
        if a["ponte_cotistas"]:
            p = a["ponte_cotistas"]
            html('<div class="rotulo">Ponte do PL (com movimentação de cotistas)</div>')
            linhas = [("PL inicial", p["pl_inicial"]), ("(+) Rendimento dos ativos", p["rendimento"]),
                      ("(+/−) Movimentação de cotistas", p["fluxo_valor"]),
                      ("(+/−) Outros efeitos (tributos, taxas, provisões, datas)", p["outros"]), ("= PL final", p["pl_final"])]
            html(tabela_html(pd.DataFrame(linhas, columns=["Item", "R$"]), numericas={"R$"}, altura=300,
                             formatos={"R$": txt_brl},
                             destaque=lambda lin: "linha-geral" if lin["Item"].startswith("=") else ""))
        elif a["ponte"]:
            p = a["ponte"]
            html('<div class="rotulo">De onde vem a diferença (extrato de conta corrente e provisões)</div>')
            linhas = [("Despesas e taxas do extrato", sum(e["valor"] for e in p["externos"])),
                      ("Variação das provisões", sum(q["delta"] for q in p["provisoes"])),
                      ("Rendas da planilha − cupom recebido em caixa", -(p["rendas"] - p["cupom_caixa"])),
                      ("Total explicado", p["explicado"]), ("Resíduo (deve ser ~0)", p["residuo"])]
            html(tabela_html(pd.DataFrame(linhas, columns=["Item", "R$"]), numericas={"R$"}, altura=300,
                             formatos={"R$": txt_brl},
                             destaque=lambda lin: "linha-total" if lin["Item"].startswith(("Total", "Resíduo")) else ""))
            with st.expander("Ver lançamentos do extrato"):
                ext = pd.DataFrame([{"Data": e["data"], "Descrição": e["descricao"], "Valor (R$)": e["valor"]} for e in p["externos"]])
                html(tabela_html(ext, numericas={"Valor (R$)"}, altura=300,
                                 formatos={"Data": lambda v: esc(f"{v:%d/%m/%Y}"), "Valor (R$)": txt_brl}))
        else:
            html(aviso_html("alerta", "Ponte não gerada para esta carteira."))

    with t_ativos:
        ranking_contribuicao(a)
        html('<div class="rotulo">Todos os ativos</div>')
        lay = a["layout"]
        formatos = {c: txt_num(2) for c in cv.VALORES}
        formatos["% Rent"] = lambda v: esc(cv.fmt_pct(v, 2))

        def realce(lin):
            if lin["Nível"] == "Total Geral":
                return "linha-geral"
            return "linha-total" if lin["Nível"] != "Ativo" else ""

        html(tabela_html(lay, numericas=set(cv.VALORES) | {"% Rent"}, formatos=formatos, destaque=realce, altura=460,
                         ocultar={"Nível"}))

    if a["avisos"]:
        with st.expander(f"Avisos ({len(a['avisos'])})"):
            html("".join(aviso_html("alerta", t) for t in a["avisos"]))


# ----------------------------------------------------------------------------
# Fonte dos dados (fica no fim da página; o valor digitado é lido pelo session_state)
# ----------------------------------------------------------------------------
def fonte_dados() -> None:
    with st.expander("Fonte dos dados", icon=":material/folder_open:"):
        c1, c2 = st.columns(2)
        c1.text_input("Arquivo da API", value=str(ARQUIVO_PADRAO), key=CHAVE_ARQUIVO, help="Gerado pela busca na API")
        c2.text_input("Pasta dos extratos", value=str(cv.EXTRATOS_PADRAO), key=CHAVE_EXTRATOS)
        st.text_input("Histórico de cotas (Monitor de cotas)", value=str(tb.ARQUIVO_HISTORICO), key=CHAVE_HISTORICO,
                      help="CSV gerado pela busca na API, com Data, IdCarteira e CotaBruta. Também aceita a base original "
                           "(DATA_COTA, ID_CARTEIRA, COTA_BRUTA) em CSV.")


# ----------------------------------------------------------------------------
# Busca na API: o usuário escolhe os IDs e o período (o mesmo que o teste.py faz)
# ----------------------------------------------------------------------------
def usar_ids_da_planilha() -> None:
    st.session_state["ids_busca"] = tb.ler_ids()


def limpar_ids() -> None:
    st.session_state["ids_busca"] = []


def resumo_busca(res: dict) -> None:
    """Mostra o resultado da última busca (depois que a página recarregou com os dados novos)."""
    html(aviso_html("ok", f"{len(res['clientes'])} carteira(s) buscada(s) na API e os dados foram atualizados.", "Busca concluída:"))
    h = res.get("historico")
    if h and "erro" in h:
        html(aviso_html("alerta", f"o histórico para o Monitor de cotas não foi gravado ({h['erro']}). A conferência foi atualizada "
                                  "normalmente.", "Atenção:"))
    elif h and h["sem_dados"]:
        html(aviso_html("alerta", f"a API não devolveu histórico de cotas para: {', '.join(h['sem_dados'])}.", "Atenção:"))
    sem = {aba: ids for aba, ids in res["sem_dados"].items() if ids and aba != "OpCotista"}
    for aba, ids in sem.items():
        html(aviso_html("alerta", f"a API não devolveu '{aba}' para: {', '.join(ids)}.", "Atenção:"))
    if res["extratos_falhos"]:
        html(aviso_html("alerta", f"não consegui baixar o extrato de conta corrente de: {', '.join(res['extratos_falhos'])}.",
                        "Atenção:"))


def secao_busca(analises: dict, caminho: Path, pasta_extratos: str, arquivo_historico: Path, tem_dados: bool) -> None:
    ids_planilha = tb.ler_ids()
    na_tela = list(analises)
    if analises:
        p = next(iter(analises.values()))["ativos"].attrs
        ini_padrao, fim_padrao = p["inicio"].date(), p["fim"].date()
    else:
        ini_padrao, fim_padrao = date.fromisoformat(tb.DATA_INICIO_PADRAO), date.fromisoformat(tb.DATA_FIM_PADRAO)

    st.session_state.setdefault("ids_busca", na_tela or ids_planilha)
    opcoes = list(dict.fromkeys(ids_planilha + na_tela + list(st.session_state["ids_busca"])))

    with st.expander("Buscar carteiras na API", expanded=not tem_dados, icon=":material/cloud_download:"):
        html('<div class="secao-sub">Escolha as carteiras e o período. A busca <b>substitui</b> os dados atuais pelas carteiras '
             'escolhidas (as que não estiverem na lista saem da análise).</div>')
        c_ids, c_per = st.columns([3, 1.6])
        c_ids.multiselect("IDs das carteiras", options=opcoes, key="ids_busca", accept_new_options=True,
                          placeholder="Escolha da lista ou digite um ID e tecle Enter",
                          help="Os IDs da planilha ids_cliente.xlsx aparecem na lista. Também dá para digitar um ID novo.")
        periodo = c_per.date_input("Período (dia-base e último dia)", value=(ini_padrao, fim_padrao), format="DD/MM/YYYY",
                                   key="periodo_busca")
        com_historico = False  # o histórico do Monitor de cotas é buscado só na aba do monitor

        b1, b2, b3, _ = st.columns([1.5, 1.8, 1, 2.2])
        buscar = b1.button("Buscar na API", icon=":material/search:", key="btn_buscar", width="stretch",
                           disabled=not st.session_state.get("ids_busca"))
        b2.button("Usar IDs da planilha", icon=":material/table_view:", key="btn_planilha", width="stretch",
                  on_click=usar_ids_da_planilha, disabled=not ids_planilha,
                  help=f"Carrega os {len(ids_planilha)} ID(s) de {tb.ARQUIVO_IDS.name}")
        b3.button("Limpar", icon=":material/close:", key="btn_limpar", width="stretch", on_click=limpar_ids)

        if not buscar:
            return
        ids = [str(i).strip() for i in st.session_state["ids_busca"] if str(i).strip()]
        invalidos = [i for i in ids if not i.isdigit()]
        if invalidos:
            html(aviso_html("div", f"IDs inválidos (use só números): {', '.join(invalidos)}."))
            return
        if not (isinstance(periodo, (tuple, list)) and len(periodo) == 2):
            html(aviso_html("div", "Escolha as duas datas do período (dia-base e último dia)."))
            return

        resultado, erro = None, None
        with st.status(f"Buscando {len(ids)} carteira(s) na API Britech…", expanded=True) as status:
            barra, andamento = st.progress(0.0), st.empty()

            def mostrar(fracao: float, texto: str) -> None:
                barra.progress(fracao)
                andamento.write(texto)

            try:
                resultado = tb.buscar_dados(ids, periodo[0].isoformat(), periodo[1].isoformat(), arquivo_saida=caminho,
                                            pasta_extratos=pasta_extratos, progresso=mostrar,
                                            historico_inicio=None, arquivo_historico=arquivo_historico)
            except (ValueError, RuntimeError, PermissionError) as e:
                erro = str(e)
            except Exception as e:  # rede, credencial etc.: mostra em vez de quebrar a página
                erro = f"Erro inesperado na busca: {e}"
            status.update(label="Busca concluída" if erro is None else "A busca falhou",
                          state="complete" if erro is None else "error")
        if erro:
            html(aviso_html("div", erro, "Não foi possível buscar:"))
        else:
            st.session_state["resultado_busca"] = resultado
            st.cache_data.clear()
            st.rerun()


# ----------------------------------------------------------------------------
# Aba: Conferência de Rentabilidade
# ----------------------------------------------------------------------------
def aba_conferencia(analises: dict, erros: dict, caminho: Path, pasta_extratos: str, erro_leitura) -> None:
    if not caminho.exists():
        html(aviso_html("alerta", f"Ainda não há dados ({caminho.name}). Escolha as carteiras acima e clique em 'Buscar na API'.",
                        "Primeira vez:"))
        return
    if erro_leitura:
        html(aviso_html("div", erro_leitura, "Não foi possível ler os dados."))
        return
    if not analises:
        html(aviso_html("alerta", "Nenhuma carteira pôde ser analisada.") + "".join(
            aviso_html("alerta", f"Carteira {i}: {msg}") for i, msg in erros.items()))
        return

    # Barra de ações
    a1, a2, a3, a4 = st.columns([3.2, 1.4, 1.6, 1.1], vertical_alignment="center")
    a1.markdown(f'<div class="fonte-dados">Dados de {datetime.fromtimestamp(caminho.stat().st_mtime):%d/%m/%Y %H:%M} · '
                f'{esc(caminho.name)}</div>', unsafe_allow_html=True)
    if a2.button("Recarregar dados", icon=":material/refresh:", width="stretch"):
        st.cache_data.clear()
        st.rerun()
    a3.download_button("Baixar Excel", data=excel_bytes(str(caminho), caminho.stat().st_mtime, pasta_extratos, versao_codigo()),
                       file_name="Rentabilidade_Tratada.xlsx", icon=":material/download:", width="stretch",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with a4.popover("Como ler", icon=":material/help:", width="stretch"):
        st.markdown(
            "<b>PL esperado pela cota</b> = PL inicial × (1 + cota). É o PL se a carteira só tivesse rentabilidade.<br><br>"
            f"{selo(cv.STATUS_OK)} o PL real fechou com o esperado.<br><br>"
            f"{selo(cv.STATUS_EXPLICADA)} o PL real difere, mas a diferença é igual ao efeito de aplicações e resgates de cotas.<br><br>"
            f"{selo(cv.STATUS_DIVERGENTE)} sobra diferença sem explicação nos dados disponíveis.<br><br>"
            "Clique em <b>Ver passo a passo</b> em qualquer carteira para abrir a conta completa.", unsafe_allow_html=True)

    # Indicadores
    total = len(analises)
    contagem = {s: sum(1 for a in analises.values() if a["status"] == s) for s in SEVERIDADE}

    def pct(n: int) -> str:
        return f"{n / total * 100:.0f}%" if total else "0%"

    html(kpis_html([
        ("total", "Carteiras analisadas", total, "na planilha"),
        ("ok", "Sem divergência", contagem[cv.STATUS_OK], pct(contagem[cv.STATUS_OK])),
        ("alerta", "Explicadas por cotistas", contagem[cv.STATUS_EXPLICADA], pct(contagem[cv.STATUS_EXPLICADA])),
        ("div", "Divergentes", contagem[cv.STATUS_DIVERGENTE], pct(contagem[cv.STATUS_DIVERGENTE])),
    ]))

    if contagem[cv.STATUS_DIVERGENTE]:
        html(aviso_html("div", f"{contagem[cv.STATUS_DIVERGENTE]} carteira(s) com divergência sem explicação. Elas aparecem primeiro.",
                        "Atenção:"))
    elif contagem[cv.STATUS_EXPLICADA]:
        html(aviso_html("alerta", "Nenhuma divergência sem explicação. As diferenças encontradas vêm de movimentação de cotas.",
                        "Tudo explicado:"))
    else:
        html(aviso_html("ok", "Todas as carteiras fecharam com a rentabilidade da cota.", "Tudo certo:"))
    for i, msg in erros.items():
        html(aviso_html("alerta", f"Carteira {i} não foi analisada: {msg}"))

    # Filtros
    f1, f2 = st.columns([3, 2], vertical_alignment="center")
    with f1:
        filtro_status = st.pills("Mostrar", options=list(SEVERIDADE), selection_mode="multi", default=list(SEVERIDADE),
                                 label_visibility="collapsed", key="conf_pills")
    with f2:
        busca = st.text_input("Buscar", placeholder="Buscar por nome ou ID da carteira", label_visibility="collapsed",
                              icon=":material/search:", key="conf_busca")

    itens = [a for a in analises.values()
             if a["status"] in (filtro_status or [])
             and (not busca or busca.lower() in a["nome"].lower() or busca in a["id"])]
    itens.sort(key=lambda a: (SEVERIDADE[a["status"]], -abs(a["resumo"]["dif_pl"])))

    if not itens:
        html(aviso_html("alerta", "Nenhuma carteira com esses filtros."))

    # Cartões
    por_linha = 3
    limite = st.session_state.setdefault("limite_cards", 12)
    for ini in range(0, min(len(itens), limite), por_linha):
        cols = st.columns(por_linha)
        for col, a in zip(cols, itens[ini:ini + por_linha]):
            r = a["resumo"]
            neg = " neg" if abs(r["dif_pl"]) > cv.TOLERANCIA_REAIS and r["dif_pl"] < 0 else ""
            with col, st.container(key=f"cartCard_{STATUS_UI[a['status']]['slug']}_{a['id']}"):
                html(f'<div class="card-topo">{selo(a["status"])}<span class="card-id">#{esc(a["id"])}</span></div>'
                     f'<div class="card-nome">{esc(a["nome"])}</div>'
                     f'<div class="kv"><span>PL final real</span><b>{esc(cv.fmt_brl(r["pl_final"]))}</b></div>'
                     f'<div class="kv"><span>Esperado pela cota</span><b>{esc(cv.fmt_brl(r["pl_final_cota"]))}</b></div>'
                     f'<div class="kv"><span>Diferença</span><b class="{neg.strip()}">{esc(cv.fmt_brl(r["dif_pl"]))}</b></div>'
                     f'<div class="resumo">{esc(resumo_curto(a))}</div>')
                if st.button("Ver passo a passo", key=f"ver_{a['id']}", icon=":material/list_alt:", width="stretch"):
                    detalhe(a)

    if len(itens) > limite:
        if st.button(f"Mostrar mais ({len(itens) - limite} restantes)", icon=":material/expand_more:", key="conf_mais"):
            st.session_state["limite_cards"] = limite + 12
            st.rerun()

    # Tabela comparativa
    st.write("")
    with st.expander("Ver tabela comparativa", icon=":material/table:"):
        if itens:
            tabela = pd.DataFrame([{
                "Carteira": a["id"], "Nome": a["nome"], "Status": a["status"],
                "PL inicial": a["resumo"]["pl_inicial"], "PL final real": a["resumo"]["pl_final"],
                "Esperado pela cota": a["resumo"]["pl_final_cota"], "Diferença": a["resumo"]["dif_pl"],
                "Efeito dos cotistas": a["resumo"]["efeito_cotas"], "Resíduo": a["resumo"]["residuo_cotistas"],
            } for a in itens])
            colunas_brl = ["PL inicial", "PL final real", "Esperado pela cota", "Diferença", "Efeito dos cotistas", "Resíduo"]
            html(tabela_html(tabela, numericas=set(colunas_brl), altura=420, dicas=DICAS_COMPARATIVA,
                             formatos={**{c: txt_brl for c in colunas_brl}, "Status": selo}))

        # Legenda: o que cada coluna significa
        html('<div class="rotulo">O que é cada coluna</div>')
        legenda = pd.DataFrame({"Coluna": list(DICAS_COMPARATIVA), "O que é": list(DICAS_COMPARATIVA.values())})
        html(tabela_html(legenda, altura=520, quebrar=True, formatos={"Coluna": lambda v: f"<b>{esc(v)}</b>"}))


# ----------------------------------------------------------------------------
# Aba: Monitor de cotas
# ----------------------------------------------------------------------------
# Acima/Abaixo: o retorno do dia saiu da banda de normalidade (para cima ou para baixo).
STATUS_MON = {
    mon.ACIMA: {"cls": "ok", "selo": "ok", "icone": "▲", "slug": "acima"},
    mon.ABAIXO: {"cls": "div", "selo": "div", "icone": "▼", "slug": "abaixo"},
    mon.DENTRO: {"cls": "total", "selo": "neutro", "icone": "●", "slug": "dentro"},
    mon.SEM_STATUS: {"cls": "alerta", "selo": "warn", "icone": "—", "slug": "sem"},
}
ORDEM_MON = {mon.ABAIXO: 0, mon.ACIMA: 1, mon.SEM_STATUS: 2, mon.DENTRO: 3}
NIVEL_ALERTA = {"erro": "div", "alerta": "alerta", "info": ""}
DICAS_MONITOR = {
    "Carteira": "ID da carteira no Britech.",
    "Nome": "Nome do cliente ou fundo (do extrato de conta corrente ou da conferência).",
    "Última cota": "Data da cota mais recente da carteira.",
    "Retorno": "Cota bruta de hoje ÷ cota bruta de ontem − 1, no último dia.",
    "252d": "Status do último dia com a janela de 252 dias (meia-vida de 60 dias).",
    "126d": "Status do último dia com a janela de 126 dias (meia-vida de 30 dias).",
    "63d": "Status do último dia com a janela de 63 dias (meia-vida de 15 dias).",
    "21d": "Status do último dia com a janela de 21 dias (meia-vida de 5 dias).",
    "Alertas": "Avisos sobre os dados da carteira: cota menor ou igual a zero, saltos acima de 50%, histórico curto, "
               "intervalos sem cota e cota desatualizada.",
}


def selo_mon(status: str) -> str:
    ui = STATUS_MON[status]
    return badge(ui["selo"], f"{ui['icone']} {status}")


def pct3(v) -> str:
    return "—" if v is None or pd.isna(v) else cv.fmt_pct(v, 3)


def bps_txt(v) -> str:
    return "—" if v is None or pd.isna(v) else f"{fmt_num(v * 10000, 1)} bps"


def nome_carteira(id_: str, analises: dict, pasta_extratos: str) -> str:
    if id_ in analises:
        return analises[id_]["nome"]
    arquivo = Path(pasta_extratos) / f"ReportExtratoContaCorrente_{id_}.xlsx"
    return (cv.ler_nome_extrato(arquivo) if arquivo.exists() else None) or f"Carteira {id_}"


def versao_monitor() -> float:
    """Data do monitor.py: se a lógica mudar, o cache é refeito."""
    return Path(mon.__file__).stat().st_mtime


@st.cache_data(show_spinner="Calculando o monitor…")
def carregar_monitor(caminho: str, mtime: float, n_desvios: float, piso_bps: float, codigo: float):
    ext = Path(caminho).suffix.lower()
    base = pd.read_excel(caminho) if ext in (".xlsx", ".xls") else pd.read_csv(caminho)
    return mon.analisar_monitor(base, {"n_desvios": n_desvios, "piso_bps": piso_bps})


def texto_monitor(r, info: dict) -> str:
    if not r["Atual"]:
        return f"Cota desatualizada: a última é de {info['ultima_data']:%d/%m/%Y}."
    if r["Status"] == mon.ACIMA:
        return f"O retorno ficou {bps_txt(r['Retorno'] - r['Lim_Sup_Hoje'])} acima do limite superior da banda."
    if r["Status"] == mon.ABAIXO:
        return f"O retorno ficou {bps_txt(r['Lim_Inf_Hoje'] - r['Retorno'])} abaixo do limite inferior da banda."
    if r["Status"] == mon.DENTRO:
        return "Retorno dentro da banda de normalidade."
    return "Sem status nesta janela: histórico curto."


def grafico_monitor(serie: pd.DataFrame, janela: int, meses: int):
    """Retornos diários (pontos), banda de normalidade e limites. Zoom e arraste com o mouse."""
    col_s, col_i, col_m = f"Status_{janela}d", f"Lim_Inf_{janela}d_Ant", f"Lim_Sup_{janela}d_Ant"
    d = serie[["Data", "Retorno", col_s, col_i, col_m]].rename(columns={col_s: "Status", col_i: "Inferior", col_m: "Superior"})
    d = d.dropna(subset=["Inferior", "Superior"])
    if d.empty:
        return None
    if meses:
        d = d[d["Data"] >= d["Data"].max() - pd.DateOffset(months=meses)]
    d = d.assign(Fora=d["Status"].isin([mon.ACIMA, mon.ABAIXO]), Status=d["Status"].fillna("Sem status"))
    x = alt.X("Data:T", title=None, axis=alt.Axis(format="%d/%m/%y", labelAngle=0))
    y_eixo = alt.Axis(format=".2%")
    banda = alt.Chart(d).mark_area(opacity=0.14, color=COR["dourado"]).encode(
        x=x, y=alt.Y("Inferior:Q", title=None, axis=y_eixo, scale=alt.Scale(zero=False)), y2="Superior:Q")
    superior = alt.Chart(d).mark_line(color=COR["ok"], strokeDash=[4, 3], strokeWidth=1.5).encode(x=x, y="Superior:Q")
    inferior = alt.Chart(d).mark_line(color=COR["div"], strokeDash=[4, 3], strokeWidth=1.5).encode(x=x, y="Inferior:Q")
    pontos = alt.Chart(d).mark_circle(opacity=0.9).encode(
        x=x, y="Retorno:Q",
        color=alt.Color("Status:N", legend=alt.Legend(orient="bottom", title=None),
                        scale=alt.Scale(domain=[mon.DENTRO, mon.ACIMA, mon.ABAIXO],
                                        range=[COR["vinho"], COR["ok"], COR["div"]])),
        size=alt.condition("datum.Fora", alt.value(80), alt.value(24)),
        tooltip=[alt.Tooltip("Data:T", title="Data", format="%d/%m/%Y"), alt.Tooltip("Retorno:Q", format=".3%"),
                 alt.Tooltip("Status:N"), alt.Tooltip("Inferior:Q", title="Limite inferior", format=".3%"),
                 alt.Tooltip("Superior:Q", title="Limite superior", format=".3%")])
    return estilizar_altair((banda + superior + inferior + pontos).properties(height=330).interactive())


@st.dialog("Monitor de cotas", width="large")
def detalhe_monitor(id_: str, res: dict, janela_inicial: int, nome: str, status_conferencia):
    serie, info = res["series"][id_], res["carteiras"][id_]
    ini, fim = info["primeira_data"], info["ultima_data"]
    html(f'<div class="dlg-nome">{esc(nome)}</div>'
         f'<div class="dlg-meta">Carteira {esc(id_)} · cotas de {ini:%d/%m/%Y} a {fim:%d/%m/%Y} ({info["n_linhas"]} dias)</div>')
    if status_conferencia:
        html(f'<div class="dlg-meta">Conferência do PL: {selo(status_conferencia)}</div>')

    janela = st.segmented_control("Janela", mon.JANELAS, default=janela_inicial, format_func=lambda j: f"{j} dias",
                                  key=f"mon_dlg_j_{id_}", label_visibility="collapsed") or janela_inicial
    r = res["resumo"][(res["resumo"]["Id"] == id_) & (res["resumo"]["Janela"] == janela)].iloc[0]
    ui = STATUS_MON[r["Status"]]
    html(f'{selo_mon(r["Status"])}<div style="height:.6rem"></div>')
    html(aviso_html(ui["cls"] if ui["cls"] != "total" else "", texto_monitor(r, info)))
    html(kpis_html([
        ("total", "Retorno do último dia", pct3(r["Retorno"]), f"em {fim:%d/%m/%Y}", True),
        ("div", "Limite inferior (hoje)", pct3(r["Lim_Inf_Hoje"]), "média − banda", True),
        ("ok", "Limite superior (hoje)", pct3(r["Lim_Sup_Hoje"]), "média + banda", True),
    ]))

    t_graf, t_dias, t_janelas, t_qual = st.tabs(["Gráfico", "Últimos dias", "Todas as janelas", "Qualidade dos dados"])

    with t_graf:
        meses = st.segmented_control("Período do gráfico", [3, 6, 12, 0], default=12, key=f"mon_dlg_m_{id_}",
                                     format_func=lambda m: "Tudo" if m == 0 else f"{m} meses",
                                     label_visibility="collapsed")
        grafico = grafico_monitor(serie, janela, 12 if meses is None else meses)
        if grafico is None:
            html(aviso_html("alerta", "Sem banda calculada nesta janela: o histórico é curto demais."))
        else:
            st.altair_chart(grafico, theme=None, width="stretch")
            html('<div class="nota-tabela">Pontos: retorno diário (verde = acima, vermelho = abaixo). Faixa dourada: banda de '
                 'normalidade (média ± desvios) calculada no dia anterior. Dá para dar zoom e arrastar com o mouse.</div>')

    with t_dias:
        cols = ["Data", "Cota", "Retorno", f"Status_{janela}d", f"Lim_Inf_{janela}d_Ant", f"Lim_Sup_{janela}d_Ant"]
        stt = serie[f"Status_{janela}d"]
        n_ac, n_ab, n_de = int(stt.eq("Acima").sum()), int(stt.eq("Abaixo").sum()), int(stt.eq("Dentro").sum())
        html(f'<div class="nota-tabela">No histórico carregado ({serie["Data"].min():%d/%m/%Y} a {serie["Data"].max():%d/%m/%Y}), '
             f'janela de {janela} dias: <b>{n_ac}</b> dias Acima, <b>{n_ab}</b> Abaixo e <b>{n_de}</b> Dentro.</div>')
        filtro = st.segmented_control("Mostrar", ["Últimos 15 dias", "Acima", "Abaixo", "Dentro"], default="Últimos 15 dias",
                                      key=f"mon_dlg_f_{id_}", label_visibility="collapsed") or "Últimos 15 dias"
        base = serie[cols] if filtro == "Últimos 15 dias" else serie.loc[stt.eq(filtro), cols]
        if filtro == "Últimos 15 dias":
            base = base.tail(15)
        ult = base.iloc[::-1].rename(columns={
            f"Status_{janela}d": "Status", f"Lim_Inf_{janela}d_Ant": "Limite inferior", f"Lim_Sup_{janela}d_Ant": "Limite superior"})
        ult["Status"] = ult["Status"].fillna(mon.SEM_STATUS)
        if ult.empty:
            html(f'<div class="nota-tabela">Nenhum dia {esc(filtro.lower())} na janela de {janela} dias.</div>')
        else:
          html(tabela_html(ult,numericas={"Cota", "Retorno", "Limite inferior", "Limite superior"}, altura=460,
                         formatos={"Data": lambda v: esc(f"{v:%d/%m/%Y}"), "Cota": txt_num(6), "Retorno": lambda v: esc(pct3(v)),
                                   "Limite inferior": lambda v: esc(pct3(v)), "Limite superior": lambda v: esc(pct3(v)),
                                   "Status": lambda v: selo_mon(v)}))
        html('<div class="nota-tabela">Os limites de cada linha são os calculados no dia anterior.</div>')

    with t_janelas:
        proximo = fim + pd.offsets.BDay(1)
        linhas = []
        for j in mon.JANELAS:
            x = res["resumo"][(res["resumo"]["Id"] == id_) & (res["resumo"]["Janela"] == j)].iloc[0]
            fora = x[mon.ACIMA] + x[mon.ABAIXO]
            linhas.append({"Janela": f"{j} dias", "Status": x["Status"], "Retorno": x["Retorno"],
                           "Limite inferior": x["Lim_Inf_Hoje"], "Limite superior": x["Lim_Sup_Hoje"],
                           "Dias abaixo": x[mon.ABAIXO], "Dias dentro": x[mon.DENTRO], "Dias acima": x[mon.ACIMA],
                           "% fora": fora / x["Dias_Com_Status"] if x["Dias_Com_Status"] else None,
                           "Inferior p/ o próximo dia": x["Lim_Inf_Proximo"], "Superior p/ o próximo dia": x["Lim_Sup_Proximo"]})
        tab = pd.DataFrame(linhas)
        fm_pct = lambda v: esc(pct3(v))
        html(tabela_html(tab, numericas={c for c in tab.columns if c not in ("Janela", "Status")}, altura=300,
                         formatos={"Status": selo_mon, "Retorno": fm_pct, "Limite inferior": fm_pct, "Limite superior": fm_pct,
                                   "% fora": lambda v: esc(cv.fmt_pct(v, 1)), "Inferior p/ o próximo dia": fm_pct,
                                   "Superior p/ o próximo dia": fm_pct}))
        html(f'<div class="nota-tabela">As colunas "p/ o próximo dia" são a banda que vale para o retorno de {proximo:%d/%m/%Y} '
             '(próximo dia útil, sem considerar feriados).</div>')

    with t_qual:
        if not info["qualidade"]:
            html(aviso_html("ok", "Nenhum problema encontrado nos dados desta carteira."))
        for nivel, texto in info["qualidade"]:
            html(aviso_html(NIVEL_ALERTA[nivel], texto))


def busca_monitor(analises: dict, caminho_historico: Path) -> None:
    """Busca do histórico de cotas: só roda quando o usuário clica (não acompanha a busca da conferência)."""
    ids_padrao = list(analises) or tb.ler_ids()
    if analises:
        fim_padrao = next(iter(analises.values()))["ativos"].attrs["fim"].date()
    else:
        fim_padrao = date.fromisoformat(tb.DATA_FIM_PADRAO)
    with st.expander("Buscar histórico de cotas na API", expanded=not caminho_historico.exists(), icon=":material/cloud_download:"):
        html('<div class="secao-sub">O monitor precisa de muito mais histórico que a conferência. Escolha as carteiras e '
             'clique em <b>Buscar cotas</b>. Leva cerca de 1 segundo por carteira.</div>')
        st.session_state.setdefault("ids_monitor", ids_padrao)
        opcoes = list(dict.fromkeys(ids_padrao + list(st.session_state["ids_monitor"])))
        c1, c2, c3 = st.columns([3, 1.3, 1.3], vertical_alignment="bottom")
        c1.multiselect("IDs das carteiras", options=opcoes, key="ids_monitor", accept_new_options=True,
                       placeholder="Escolha da lista ou digite um ID e tecle Enter")
        desde = c2.date_input("Histórico desde", value=fim_padrao - timedelta(days=tb.HISTORICO_DIAS_PADRAO),
                              format="DD/MM/YYYY", key="mon_desde")
        ate = c3.date_input("Até", value=fim_padrao, format="DD/MM/YYYY", key="mon_ate")
        if not st.button("Buscar cotas", icon=":material/search:", key="btn_buscar_cotas",
                         disabled=not st.session_state.get("ids_monitor")):
            return
        ids = [str(i).strip() for i in st.session_state["ids_monitor"] if str(i).strip()]
        invalidos = [i for i in ids if not i.isdigit()]
        if invalidos:
            html(aviso_html("div", f"IDs inválidos (use só números): {', '.join(invalidos)}."))
            return
        if desde >= ate:
            html(aviso_html("div", "A data inicial precisa ser anterior à data final."))
            return
        erro, res = None, None
        with st.status(f"Buscando cotas de {len(ids)} carteira(s)…", expanded=True) as status:
            barra, andamento = st.progress(0.0), st.empty()

            def mostrar(fracao: float, texto: str) -> None:
                barra.progress(fracao)
                andamento.write(texto)

            try:
                res = tb.buscar_historico_cotas(ids, desde.isoformat(), ate.isoformat(), arquivo_saida=caminho_historico,
                                                progresso=mostrar)
            except (ValueError, RuntimeError, PermissionError) as e:
                erro = str(e)
            except Exception as e:
                erro = f"Erro inesperado na busca: {e}"
            status.update(label="Busca concluída" if erro is None else "A busca falhou",
                          state="complete" if erro is None else "error")
        if erro:
            html(aviso_html("div", erro, "Não foi possível buscar:"))
        else:
            st.cache_data.clear()
            st.rerun()


def aba_monitor(analises: dict, pasta_extratos: str, caminho_historico: Path) -> None:
    busca_monitor(analises, caminho_historico)
    if not caminho_historico.exists():
        html(aviso_html("alerta", "Ainda não há histórico de cotas. Clique em 'Buscar cotas' acima.", "Monitor sem dados:"))
        return

    c1, c2, c3 = st.columns([2.4, 1.6, 1.1], vertical_alignment="center")
    janela = c1.segmented_control("Janela", mon.JANELAS, default=63, format_func=lambda j: f"{j} dias", key="mon_janela",
                                  label_visibility="collapsed") or 63
    with c2.popover("Parâmetros", icon=":material/tune:", width="stretch"):
        n_desvios = st.number_input("Largura da banda (desvios-padrão)", min_value=0.5, max_value=4.0, value=1.5, step=0.1,
                                    key="mon_desvios", help="Quanto maior, menos dias saem da banda.")
        piso_bps = st.number_input("Piso do desvio (bps por dia)", min_value=0.0, max_value=10.0, value=0.5, step=0.1,
                                   key="mon_piso", help="Evita banda de largura zero em carteiras de retorno quase constante. "
                                                        "Use 0 para o cálculo sem piso.")
    with c3.popover("Como ler", icon=":material/help:", width="stretch"):
        st.markdown(
            "O retorno de <b>hoje</b> é comparado com a banda calculada <b>ontem</b>: média ± desvios dos retornos dos dias "
            "anteriores, com mais peso para os dias recentes (EWMA).<br><br>"
            f"{selo_mon(mon.ACIMA)} retorno acima do limite superior.<br><br>"
            f"{selo_mon(mon.ABAIXO)} retorno abaixo do limite inferior.<br><br>"
            f"{selo_mon(mon.DENTRO)} dentro da banda.<br><br>"
            f"{selo_mon(mon.SEM_STATUS)} histórico curto demais para a janela, ou cota desatualizada.<br><br>"
            "As janelas (252, 126, 63 e 21 dias) mudam o quanto o passado recente pesa. "
            "Fora da banda não é erro: é um dia incomum para aquela carteira.", unsafe_allow_html=True)

    try:
        res = carregar_monitor(str(caminho_historico), caminho_historico.stat().st_mtime, float(n_desvios), float(piso_bps),
                               versao_monitor())
    except ValueError as e:
        html(aviso_html("div", str(e), "Não foi possível ler o histórico de cotas."))
        return

    resumo = mon.resumo_da_janela(res, janela)
    ids = list(resumo.index)
    atual = resumo["Atual"]
    contagem = {s: int(((resumo["Status"] == s) & atual).sum()) for s in (mon.DENTRO, mon.ACIMA, mon.ABAIXO)}
    sem_status = len(ids) - sum(contagem.values())

    html(kpis_html([
        ("total", "Carteiras monitoradas", len(ids), f"cotas até {res['data_ref']:%d/%m/%Y}"),
        ("total", "Dentro da banda", contagem[mon.DENTRO], f"janela de {janela} dias"),
        ("ok", "Acima da banda", contagem[mon.ACIMA], "no último dia"),
        ("div", "Abaixo da banda", contagem[mon.ABAIXO], "no último dia"),
        ("alerta", "Sem status", sem_status, "histórico curto ou cota desatualizada"),
    ]))
    fora = contagem[mon.ACIMA] + contagem[mon.ABAIXO]
    if fora:
        html(aviso_html("alerta", f"{fora} carteira(s) fora da banda no último dia, na janela de {janela} dias. Elas aparecem primeiro.",
                        "Atenção:"))
    elif contagem[mon.DENTRO]:
        html(aviso_html("ok", f"Todas as carteiras com status estão dentro da banda na janela de {janela} dias.", "Tudo dentro do normal:"))
    desatualizadas = [i for i in ids if not atual[i]]
    if desatualizadas:
        html(aviso_html("alerta", f"cota mais antiga que a data de referência ({res['data_ref']:%d/%m/%Y}) em: "
                                  f"{', '.join(desatualizadas)}.", "Cota desatualizada:"))
    html(f'<div class="fonte-dados">Histórico de {datetime.fromtimestamp(caminho_historico.stat().st_mtime):%d/%m/%Y %H:%M} · '
         f'{esc(caminho_historico.name)} · banda de {res["params"]["n_desvios"]:g} desvio(s) · piso de {res["params"]["piso_bps"]:g} bps</div>')

    # Filtros
    f1, f2 = st.columns([3, 2], vertical_alignment="center")
    with f1:
        filtro = st.pills("Mostrar", options=list(ORDEM_MON), selection_mode="multi", default=list(ORDEM_MON),
                          label_visibility="collapsed", key="mon_pills")
    with f2:
        busca = st.text_input("Buscar", placeholder="Buscar por nome ou ID da carteira", label_visibility="collapsed",
                              icon=":material/search:", key="mon_busca")
    nomes = {i: nome_carteira(i, analises, pasta_extratos) for i in ids}
    itens = [i for i in ids if resumo.loc[i, "Status"] in (filtro or []) and (not busca or busca.lower() in nomes[i].lower() or busca in i)]
    itens.sort(key=lambda i: (ORDEM_MON[resumo.loc[i, "Status"]], -abs(resumo.loc[i, "Retorno"] if pd.notna(resumo.loc[i, "Retorno"]) else 0)))
    if not itens:
        html(aviso_html("alerta", "Nenhuma carteira com esses filtros."))

    # Cartões
    por_linha = 3
    limite = st.session_state.setdefault("limite_mon", 12)
    for ini in range(0, min(len(itens), limite), por_linha):
        cols = st.columns(por_linha)
        for col, i in zip(cols, itens[ini:ini + por_linha]):
            r, info = resumo.loc[i], res["carteiras"][i]
            status = r["Status"]
            fora_n = int(r[mon.ACIMA] + r[mon.ABAIXO])
            fora_pct = f"{fora_n / r['Dias_Com_Status'] * 100:.1f}%".replace(".", ",") if r["Dias_Com_Status"] else "—"
            alertas = len(info["qualidade"])
            with col, st.container(key=f"monCard_{STATUS_MON[status]['slug']}_{i}"):
                html(f'<div class="card-topo">{selo_mon(status)}<span class="card-id">#{esc(i)}</span></div>'
                     f'<div class="card-nome">{esc(nomes[i])}</div>'
                     f'<div class="kv"><span>Retorno do último dia</span><b>{esc(pct3(r["Retorno"]))}</b></div>'
                     f'<div class="kv"><span>Banda de hoje</span><b>{esc(pct3(r["Lim_Inf_Hoje"]))} a {esc(pct3(r["Lim_Sup_Hoje"]))}</b></div>'
                     f'<div class="kv"><span>Dias fora da banda</span><b>{fora_pct} ({fora_n})</b></div>'
                     f'<div class="kv"><span>Última cota</span><b>{info["ultima_data"]:%d/%m/%Y}</b></div>'
                     f'<div class="resumo">{esc(texto_monitor(r, info))}'
                     f'{f" · {alertas} alerta(s) de dados" if alertas else ""}</div>')
                if st.button("Ver gráfico", key=f"mver_{i}", icon=":material/show_chart:", width="stretch"):
                    detalhe_monitor(i, res, janela, nomes[i], analises[i]["status"] if i in analises else None)

    if len(itens) > limite:
        if st.button(f"Mostrar mais ({len(itens) - limite} restantes)", icon=":material/expand_more:", key="mon_mais"):
            st.session_state["limite_mon"] = limite + 12
            st.rerun()

    # Tabela com todas as janelas
    st.write("")
    with st.expander("Ver tabela com todas as janelas", icon=":material/table:"):
        linhas = []
        for i in itens:
            info = res["carteiras"][i]
            por_janela = {j: res["resumo"][(res["resumo"]["Id"] == i) & (res["resumo"]["Janela"] == j)].iloc[0] for j in mon.JANELAS}
            linhas.append({"Carteira": i, "Nome": nomes[i], "Última cota": info["ultima_data"], "Retorno": por_janela[janela]["Retorno"],
                           **{f"{j}d": por_janela[j]["Status"] for j in mon.JANELAS}, "Alertas": len(info["qualidade"])})
        if linhas:
            tab = pd.DataFrame(linhas)
            fm = {"Última cota": lambda v: esc(f"{v:%d/%m/%Y}"), "Retorno": lambda v: esc(pct3(v))}
            fm.update({f"{j}d": selo_mon for j in mon.JANELAS})
            html(tabela_html(tab, numericas={"Retorno", "Alertas"}, formatos=fm, altura=420, dicas=DICAS_MONITOR))
        html('<div class="rotulo">O que é cada coluna</div>')
        legenda = pd.DataFrame({"Coluna": list(DICAS_MONITOR), "O que é": list(DICAS_MONITOR.values())})
        html(tabela_html(legenda, altura=520, quebrar=True, formatos={"Coluna": lambda v: f"<b>{esc(v)}</b>"}))

    # Excel do monitor (gerado sob demanda: os gráficos levam alguns segundos)
    st.write("")
    x1, x2, x3 = st.columns([1.6, 1.6, 3], vertical_alignment="center")
    com_graficos = x3.checkbox("Incluir gráficos no Excel", value=len(ids) <= 40, key="mon_xl_graficos",help="Um gráfico por carteira e janela (até 160 no total).")
    assinatura = (str(caminho_historico), caminho_historico.stat().st_mtime, float(n_desvios), float(piso_bps), com_graficos)
    if x1.button("Gerar Excel do monitor", icon=":material/table_chart:", width="stretch", key="mon_xl_gerar"):
        with st.spinner("Gerando o Excel…"):
            dados, info_xl = mon.excel_monitor_bytes(res, nomes, com_graficos)
        st.session_state["mon_xl"] = {"assinatura": assinatura, "dados": dados, "info": info_xl,"nome": f"Monitor_Cotas_{datetime.now():%d%m%Y_%H%M%S}.xlsx"}
    guardado = st.session_state.get("mon_xl")
    if guardado and guardado["assinatura"] == assinatura:
        x2.download_button("Baixar Excel do monitor", data=guardado["dados"], file_name=guardado["nome"], width="stretch",
icon=":material/download:", key="mon_xl_baixar",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if guardado["info"]["omitidos"]:
            html(aviso_html("alerta", "Muitas carteiras: os gráficos foram omitidos do Excel (o limite é de 160 gráficos)."))
    elif guardado:
        html('<div class="nota-tabela">Os dados ou parâmetros mudaram desde o último Excel: gere de novo.</div>')


# ----------------------------------------------------------------------------
# Página
# ----------------------------------------------------------------------------
arquivo = st.session_state.get(CHAVE_ARQUIVO, str(ARQUIVO_PADRAO))
pasta_extratos = st.session_state.get(CHAVE_EXTRATOS, str(cv.EXTRATOS_PADRAO))
caminho = Path(arquivo)
caminho_historico = Path(st.session_state.get(CHAVE_HISTORICO, str(tb.ARQUIVO_HISTORICO)))

ANALISES, ERROS, erro_leitura = {}, {}, None
if caminho.exists():
    try:
        ANALISES, ERROS = carregar(str(caminho), caminho.stat().st_mtime, pasta_extratos, versao_codigo())
    except ValueError as e:
        erro_leitura = str(e)

if ANALISES:
    primeiro = next(iter(ANALISES.values()))["ativos"]
    topo("Rentabilidade e Cotas",
         f"Conferência do PL com a cota ({primeiro.attrs['inicio']:%d/%m/%Y} a {primeiro.attrs['fim']:%d/%m/%Y}) "
         "e monitor da rentabilidade diária")
else:
    topo("Rentabilidade e Cotas", "Conferência do PL com a cota e monitor da rentabilidade diária")

ultima_busca = st.session_state.pop("resultado_busca", None)
if ultima_busca:
    resumo_busca(ultima_busca)
secao_busca(ANALISES, caminho, pasta_extratos, caminho_historico, tem_dados=bool(ANALISES) or caminho_historico.exists())

tab_conferencia, tab_monitor = st.tabs(["Conferência de Rentabilidade", "Monitor de Cotas"])
with tab_conferencia:
    aba_conferencia(ANALISES, ERROS, caminho, pasta_extratos, erro_leitura)
with tab_monitor:
    aba_monitor(ANALISES, pasta_extratos, caminho_historico)

fonte_dados()

