import html
import io
import json
import re
import shutil
import subprocess
import sys
import time
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# O portal roda este arquivo via st.navigation: garante a pasta do app (importador) e a raiz (tema) no sys.path.
for _pasta in (Path(__file__).resolve().parent, Path(__file__).resolve().parents[2]):
    if str(_pasta) not in sys.path:
        sys.path.insert(0, str(_pasta))

from importador import MARCADOR_ARQUIVO, TEMPO_MAX_LOGIN  # noqa: E402
from tema import COR, CORES_PRIORIDADE, CSS_ERROS, aplicar_tema, estilizar_figura, topo  # noqa: E402

# --- INÍCIO CONCILIAÇÃO -----------------------------------------------------
# Colunas usadas para localizar o cabeçalho de cada planilha
COLS_TAG = ["Operação", "Valor Solicitado"]
COLS_XP = ["Código do Cliente", "Operação", "Valor Solicitado"]
CANDIDATOS_NOME_TAG = ["Nome Cliente", "Cliente", "Nome"]
CANDIDATOS_CONTA_TAG = ["Conta Cliente", "Conta", "Número da conta"]

ALIASES = {
    "Código do Cliente": ["Código do Cliente", "Código Cliente", "Cód. do Cliente", "Cód Cliente", "Conta", "Conta Cliente", "Número da conta"],
    "Operação": ["Operação", "Tipo de Operação", "Tipo Operação", "Tipo"],
    "Valor Solicitado": ["Valor Solicitado", "Valor", "Valor da Operação"],
}

COL_CONC_TAG = "Conciliação TAG x XP"
COL_CONC_XP = "Conciliação XP x TAG"


def normalizar(txt) -> str:
    """Texto sem acento, sem pontuação, maiúsculo e sem espaços extras (para comparar nomes de colunas)."""
    if pd.isna(txt):
        return ""
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    t = re.sub(r"[^A-Za-z0-9 ]", " ", t)
    return re.sub(r"\s+", " ", t).strip().upper()


def achar_coluna(df: pd.DataFrame, candidatos: list):
    """Acha a coluna do DataFrame cujo nome (normalizado) esteja entre os candidatos."""
    mapa = {normalizar(c): c for c in df.columns}
    for cand in candidatos:
        if normalizar(cand) in mapa:
            return mapa[normalizar(cand)]
    return None


def para_float(v):
    """Converte número ou texto (inclusive '1.234,56') em float; None se não der."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace(" ", "")
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def para_conta(v):
    """Converte a conta em int (3832167.0 -> 3832167); None se não der."""
    f = para_float(v)
    return int(f) if f is not None else None


def ler_com_cabecalho(origem, obrigatorias, max_linhas: int = 15) -> pd.DataFrame:
    """Lê um Excel procurando o cabeçalho nas primeiras linhas de TODAS as abas.

    Ignora acento, maiúsculas e espaços, aceita nomes parecidos (ALIASES) e devolve
    as colunas já com o nome padrão (ex.: 'Codigo do cliente' vira 'Código do Cliente').
    """
    if hasattr(origem, "seek"):
        origem.seek(0)
    abas = pd.read_excel(origem, header=None, sheet_name=None)

    vistos = []  # para mostrar no erro o que foi encontrado
    for nome_aba, bruto in abas.items():
        primeira = True
        for i in range(min(max_linhas, len(bruto))):
            valores = bruto.iloc[i].tolist()
            textos = [str(v).strip() for v in valores if pd.notna(v)]
            if primeira and textos:
                vistos.append(f"aba '{nome_aba}', linha {i + 1}: {', '.join(textos[:10])}")
                primeira = False

            posicao = {normalizar(v): j for j, v in enumerate(valores) if pd.notna(v)}
            achados = {}
            for col in obrigatorias:
                for alias in ALIASES.get(col, [col]):
                    if normalizar(alias) in posicao:
                        achados[col] = posicao[normalizar(alias)]
                        break
            if len(achados) == len(obrigatorias):
                df = bruto.iloc[i + 1 :].copy()
                nomes = [str(c).strip() for c in valores]
                for col, j in achados.items():
                    nomes[j] = col  # nome padrão
                df.columns = nomes
                return df.dropna(how="all").reset_index(drop=True)

    raise ValueError(
        "Não encontrei o cabeçalho esperado. Colunas necessárias: "
        + ", ".join(obrigatorias)
        + ". O que encontrei no arquivo: "
        + (" | ".join(vistos) if vistos else "arquivo vazio")
    )


def ler_tag(origem) -> pd.DataFrame:
    """Lê a planilha da TAG (caminho ou arquivo enviado). Aceita xlsx/xls e csv."""
    nome = str(getattr(origem, "name", origem)).lower()
    if nome.endswith(".csv"):
        return pd.read_csv(origem, sep=";", encoding="utf-8-sig")
    try:
        return ler_com_cabecalho(origem, COLS_TAG)
    except ValueError:
        if hasattr(origem, "seek"):
            origem.seek(0)
        return pd.read_excel(origem)


def exigir_colunas(df: pd.DataFrame, colunas: list, nome_planilha: str) -> None:
    faltando = [c for c in colunas if c not in df.columns]
    if faltando:
        raise ValueError(
            f"Na planilha {nome_planilha} não encontrei: {', '.join(faltando)}. "
            f"Colunas encontradas: {', '.join(map(str, df.columns))}"
        )


def extrair_contas_tag(df_tag: pd.DataFrame) -> pd.DataFrame:
    """DataFrame só com Nome Cliente e Conta Cliente da planilha da TAG (sem repetições)."""
    col_nome = achar_coluna(df_tag, CANDIDATOS_NOME_TAG)
    col_conta = achar_coluna(df_tag, CANDIDATOS_CONTA_TAG)
    if col_nome is None or col_conta is None:
        raise ValueError(
            "Na planilha da TAG preciso das colunas do nome (Nome Cliente ou Cliente) e 'Conta Cliente'. "
            f"Colunas encontradas: {', '.join(map(str, df_tag.columns))}"
        )
    contas = df_tag[[col_nome, col_conta]].copy()
    contas.columns = ["Nome Cliente", "Conta Cliente"]
    contas["Conta Cliente"] = contas["Conta Cliente"].map(para_conta)
    contas = contas.dropna(subset=["Conta Cliente"]).drop_duplicates().reset_index(drop=True)
    contas["Conta Cliente"] = contas["Conta Cliente"].astype("Int64")
    return contas


def normalizar_operacao(x):
    """Padroniza a operação: maiúsculas, sem acento, 1.0 vira '1'."""
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return None
    s = str(x).strip().upper()
    if not s:
        return None

    # Remove acentos para garantir que APLICAÇÃO == APLICACAO
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()

    try:
        f = float(s.replace(",", "."))
        if f.is_integer():
            return str(int(f))
    except ValueError:
        pass
    return s


def chaves(contas, operacoes, valores):
    """Listas padronizadas: conta (int), operação e valor (2 casas decimais, positivo)."""
    c = [para_conta(x) for x in contas]
    o = [normalizar_operacao(x) for x in operacoes]

    v = []
    for x in valores:
        f = para_float(x)
        if f is None:
            v.append(None)
        else:
            # Valor absoluto (ignora resgates negativos) arredondado a 2 casas
            v.append(round(abs(f), 2))

    return c, o, v


def avaliar(origem, destino, comparar_op_valor: bool, nome_destino: str):
    """Cruza os dados e verifica explicitamente o Valor Solicitado."""
    contas_dest, ops_dest, vals_dest = destino

    # Agrupa as boletas por conta no destino
    mapa_dest = defaultdict(list)
    for c, o, v in zip(contas_dest, ops_dest, vals_dest):
        if c is not None:
            mapa_dest[c].append({"operacao": o, "valor": v})

    presenca, conciliacao, motivo = [], [], []

    for c, o, v in zip(*origem):
        if c is None:
            presenca.append("")
            conciliacao.append("")
            motivo.append("Conta em branco ou inválida")
            continue

        if c not in mapa_dest:
            presenca.append("Não")
            conciliacao.append("DIVERGENTE")
            motivo.append(f"Conta não encontrada na {nome_destino}")
            continue

        presenca.append("Sim")

        if not comparar_op_valor:
            conciliacao.append("CONCILIADO")
            motivo.append("")
            continue

        if v is None:
            conciliacao.append("")
            motivo.append("Valor Solicitado em branco na origem")
            continue

        boletas_destino = mapa_dest[c]
        valores_destino = [b["valor"] for b in boletas_destino]

        # 1º Passo: verifica se o Valor Solicitado exato existe para esta conta
        if v in valores_destino:
            # 2º Passo: verifica se a Operação também coincide
            boleta_exata = next((b for b in boletas_destino if b["valor"] == v and b["operacao"] == o), None)

            if boleta_exata:
                conciliacao.append("CONCILIADO")
                motivo.append("")
                boletas_destino.remove(boleta_exata)  # consome o registro para evitar falsos duplicados
            else:
                conciliacao.append("DIVERGENTE")
                ops_encontradas = [b["operacao"] for b in boletas_destino if b["valor"] == v]
                motivo.append(f"Valor Solicitado coincide ({v}), mas a Operação é diferente. Encontrado: {ops_encontradas}")
        else:
            conciliacao.append("DIVERGENTE")
            motivo.append(f"Valor divergente! Solicitado: {v} | Valores registrados na {nome_destino}: {valores_destino}")

    return presenca, conciliacao, motivo


def rodar_conciliacao(df_tag: pd.DataFrame, df_xp: pd.DataFrame, comparar_op_valor: bool = True) -> dict:
    """Cruza TAG x XP pela conta do cliente, nos dois sentidos."""
    exigir_colunas(df_tag, ["Operação", "Valor Solicitado"], "da TAG")
    exigir_colunas(df_xp, COLS_XP, "da XP")

    # 1. DataFrame só com Nome Cliente + Conta Cliente da TAG
    contas_tag = extrair_contas_tag(df_tag)
    nome_por_conta = {}
    for nome, conta in zip(contas_tag["Nome Cliente"], contas_tag["Conta Cliente"]):
        nome_por_conta.setdefault(int(conta), nome)  # primeiro nome da conta

    col_conta_tag = achar_coluna(df_tag, CANDIDATOS_CONTA_TAG)
    k_tag = chaves(df_tag[col_conta_tag], df_tag["Operação"], df_tag["Valor Solicitado"])
    k_xp = chaves(df_xp["Código do Cliente"], df_xp["Operação"], df_xp["Valor Solicitado"])

    # 2. TAG -> XP: cada boleta da TAG existe na XP?
    res_tag = df_tag.copy()
    res_tag["Conta Cliente"] = pd.array(k_tag[0], dtype="Int64")
    if "Status" in res_tag.columns:
        res_tag["Status TAG"] = ["OK" if str(s).strip() == "Confirmada" else "VERIFICAR" for s in res_tag["Status"]]
    pres, conc, mot = avaliar(k_tag, k_xp, comparar_op_valor, "XP")
    res_tag["Conta na XP"], res_tag[COL_CONC_TAG], res_tag["Motivo"] = pres, conc, mot

    # 3. XP -> TAG: cada boleta da XP existe na TAG? (o nome do cliente vem do DataFrame da TAG)
    res_xp = df_xp.copy()
    res_xp["Nome Cliente (TAG)"] = [nome_por_conta.get(c, "") if c is not None else "" for c in k_xp[0]]
    pres, conc, mot = avaliar(k_xp, k_tag, comparar_op_valor, "TAG")
    res_xp["Conta na TAG"], res_xp[COL_CONC_XP], res_xp["Motivo"] = pres, conc, mot

    return {"contas": contas_tag, "tag": res_tag, "xp": res_xp, "comparar": comparar_op_valor}


# ----------------------------------------------------------------------------
# Tabela estilizada (HTML)
# ----------------------------------------------------------------------------
COLUNAS_VALOR = {"Valor Solicitado"}
COLUNAS_ID = {"Conta Cliente", "Código do Cliente", "Código da Boleta", "Código da operação"}
COLUNAS_SIM_NAO = {"Conta na XP", "Conta na TAG"}


def esc(txt) -> str:
    """Escapa o texto para HTML (e o $, que o Streamlit trata como fórmula)."""
    return html.escape(str(txt).strip()).replace("$", "&#36;")


def valor_br(v) -> str:
    f = para_float(v)
    if f is None:
        return esc(v)
    return f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def celula_html(col, v, col_conc) -> str:
    vazio = v is None or (not isinstance(v, str) and pd.isna(v)) or (isinstance(v, str) and not v.strip())
    if vazio:
        return '<span class="vazio">—</span>' if col == col_conc else ""
    if col == col_conc:
        cls = {"CONCILIADO": "ok", "DIVERGENTE": "div"}.get(v, "neutro")
        return f'<span class="badge {cls}">{esc(v)}</span>'
    if col in COLUNAS_SIM_NAO:
        return f'<span class="badge {"ok-suave" if v == "Sim" else "div-suave"}">{esc(v)}</span>'
    if col == "Status TAG":
        return f'<span class="badge {"ok-suave" if v == "OK" else "warn"}">{esc(v)}</span>'
    if col in COLUNAS_VALOR:
        return valor_br(v)
    if col in COLUNAS_ID:
        n = para_conta(v)
        return esc(n) if n is not None else esc(v)
    if col == "Operação":
        f = para_float(v)
        return esc(int(f)) if f is not None and f.is_integer() else esc(v)
    if col == "Motivo":
        return f'<span class="motivo">{esc(v)}</span>'
    return esc(v)


def tabela_html(df: pd.DataFrame, colunas: list, col_conc=None, limite: int = 1500) -> str:
    """Monta a tabela em HTML (cabeçalho fixo, linhas zebradas, selos coloridos)."""
    df = df.loc[:, ~df.columns.duplicated()]
    colunas = [c for c in colunas if c in df.columns]
    vis = df[colunas].head(limite)

    cab = "".join(f'<th class="{"num" if c in COLUNAS_VALOR else ""}">{esc(c)}</th>' for c in colunas)
    linhas = []
    for _, lin in vis.iterrows():
        classe = ' class="linha-div"' if col_conc and lin[col_conc] == "DIVERGENTE" else ""
        tds = "".join(
            f'<td class="{"num" if c in COLUNAS_VALOR else ""}">{celula_html(c, lin[c], col_conc)}</td>'
            for c in colunas
        )
        linhas.append(f"<tr{classe}>{tds}</tr>")

    nota = ""
    if len(df) > limite:
        nota = f'<div class="nota-tabela">Mostrando as primeiras {limite} de {len(df)} linhas. O Excel traz todas.</div>'
    return (
        '<div class="tabela-wrap"><table class="tabela">'
        f"<thead><tr>{cab}</tr></thead><tbody>{''.join(linhas)}</tbody></table></div>{nota}"
    )


def kpis_html(total: int, n_ok: int, n_div: int, n_vazio: int) -> str:
    def pct(n):
        return f"{(n / total * 100):.1f}%".replace(".", ",") if total else "0%"

    cartoes = [
        ("total", "Boletas", total, "na planilha"),
        ("ok", "Conciliadas", n_ok, pct(n_ok)),
        ("div", "Divergentes", n_div, pct(n_div)),
    ]
    if n_vazio:
        cartoes.append(("vazio", "Em branco", n_vazio, "sem dados para conferir"))
    pct_ok = (n_ok / total * 100) if total else 0
    progresso = (
        f'<div class="progresso-txt">{pct(n_ok)} conciliado ({n_ok} de {total} boletas)</div>'
        f'<div class="progresso"><div style="width:{pct_ok:.1f}%"></div></div>'
    )
    return '<div class="kpis">' + "".join(
        f'<div class="kpi {cls}"><div class="kpi-label">{rot}</div>'
        f'<div class="kpi-valor">{valor}</div><div class="kpi-sub">{sub_}</div></div>'
        for cls, rot, valor, sub_ in cartoes
    ) + "</div>" + progresso


def etapas_html(feitas: list) -> str:
    """Faixa com as etapas da conferência (a primeira não concluída fica em destaque)."""
    nomes = ["Base da TAG", "Planilha da XP", "Conciliação"]
    atual = next((i for i, ok in enumerate(feitas) if not ok), None)
    itens = []
    for i, (nome, ok) in enumerate(zip(nomes, feitas), start=1):
        cls = "feita" if ok else ("atual" if i - 1 == atual else "")
        marca = "✓" if ok else str(i)
        itens.append(f'<div class="etapa {cls}"><span class="n">{marca}</span>{nome}</div>')
    return '<div class="etapas">' + "".join(itens) + "</div>"


def _como_arquivo(dados: bytes, nome: str) -> io.BytesIO:
    f = io.BytesIO(dados)
    f.name = nome
    return f


@st.cache_data(show_spinner=False, max_entries=8)
def ler_tag_cache(dados: bytes, nome: str) -> pd.DataFrame:
    """Lê a planilha da TAG uma única vez por arquivo (cache pelo conteúdo)."""
    return ler_tag(_como_arquivo(dados, nome))


@st.cache_data(show_spinner=False, max_entries=8)
def ler_xp_cache(dados: bytes, nome: str) -> pd.DataFrame:
    """Lê a planilha da XP uma única vez por arquivo (cache pelo conteúdo)."""
    return ler_com_cabecalho(_como_arquivo(dados, nome), COLS_XP)


def gerar_excel(r: dict) -> bytes:
    """Monta o Excel da conciliação (feito uma vez por RUN, não a cada clique na tela)."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        r["tag"].to_excel(escritor, sheet_name="TAG x XP", index=False)
        r["xp"].to_excel(escritor, sheet_name="XP x TAG", index=False)
        r["contas"].to_excel(escritor, sheet_name="Contas TAG", index=False)
    return buffer.getvalue()


# --- FIM CONCILIAÇÃO --------------------------------------------------------

# ----------------------------------------------------------------------------
# Página, tema e cabeçalho
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Mesa de Boletas | TAG Investimentos",
    page_icon="logo_tag.png" if Path(__file__).resolve().with_name("logo_tag.png").exists() else None,
    layout="wide",
    initial_sidebar_state="collapsed",
)
aplicar_tema()
topo()


def estado_vazio(titulo: str, descricao: str) -> None:
    st.markdown(
        f'<div class="card"><h3>{titulo}</h3><p>{descricao}</p></div>',
        unsafe_allow_html=True,
    )


# --- INÍCIO CONTROLE DE ERROS -----------------------------------------------
ARQUIVO_ERROS = Path(__file__).resolve().with_name("controle_erros.json")
ARQUIVO_SNAPSHOT = Path(__file__).resolve().with_name("snapshot_ata.json")
COLUNAS_ERROS = ["Problemas / Gargalos", "Em análise", "Em tratamento", "Solucionados"]
COL_SOLUCIONADOS = COLUNAS_ERROS[-1]
PRIORIDADES = ["Baixa", "Média", "Alta", "Crítica"]
SLUG_PRIORIDADE = {"Baixa": "baixa", "Média": "media", "Alta": "alta", "Crítica": "critica"}

# Categorias de problema (edite esta lista para ajustar às necessidades da mesa)
SEM_CATEGORIA = "Sem categoria"
CATEGORIAS = [
    SEM_CATEGORIA,
    "Cadastro / conta do cliente",
    "Valor divergente",
    "Operação incorreta",
    "Horário / cut-off",
    "Sistema / integração",
    "Comunicação com a XP",
    "Processo manual",
    "Outros",
]


def categoria_do_card(c: dict) -> str:
    return c.get("categoria") or SEM_CATEGORIA


def opcoes_categoria(atual: str) -> list:
    """Lista de categorias; mantém a atual mesmo que tenha saído da lista."""
    return CATEGORIAS if atual in CATEGORIAS else CATEGORIAS + [atual]


def agora() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M")


def hoje() -> str:
    return datetime.now().strftime("%d/%m/%Y")


def carregar_erros() -> list:
    if not ARQUIVO_ERROS.exists():
        return []
    try:
        return json.loads(ARQUIVO_ERROS.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        # arquivo corrompido: guarda uma cópia e começa limpo, sem perder o original
        ARQUIVO_ERROS.replace(ARQUIVO_ERROS.with_name("controle_erros_corrompido.json"))
        return []


PASTA_BACKUP = Path(__file__).resolve().with_name("backups_erros")
MAX_BACKUPS = 30  # quantos dias de backup manter


def fazer_backup_diario() -> None:
    """Copia o controle_erros.json para backups_erros/ uma vez por dia (estado do início do dia)."""
    if not ARQUIVO_ERROS.exists():
        return
    destino = PASTA_BACKUP / f"controle_erros_{datetime.now():%Y-%m-%d}.json"
    if destino.exists():
        return
    try:
        PASTA_BACKUP.mkdir(exist_ok=True)
        shutil.copy2(ARQUIVO_ERROS, destino)
        for antigo in sorted(PASTA_BACKUP.glob("controle_erros_*.json"))[:-MAX_BACKUPS]:
            antigo.unlink()
    except OSError:
        pass  # o backup nunca pode impedir o salvamento


def salvar_erros(cards: list) -> None:
    fazer_backup_diario()
    tmp = ARQUIVO_ERROS.with_name("controle_erros.tmp")
    tmp.write_text(json.dumps(cards, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(ARQUIVO_ERROS)


def criar_card(
    titulo: str, descricao: str, prioridade: str, ocorrencias: int,
    categoria: str = SEM_CATEGORIA, causa_raiz: str = "",
) -> None:
    cards = carregar_erros()
    cards.append({
        "id": max((c["id"] for c in cards), default=0) + 1,
        "titulo": titulo.strip(),
        "descricao": descricao.strip(),
        "coluna": COLUNAS_ERROS[0],
        "prioridade": prioridade,
        "ocorrencias": int(ocorrencias),
        "solucao": "",
        "categoria": categoria,
        "causa_raiz": causa_raiz.strip(),
        "conexoes": [],
        "historico": [{"data": agora(), "qtd": int(ocorrencias), "nota": "Registro inicial"}],
        "criado_em": agora(),
        "atualizado_em": agora(),
        "ultima_ocorrencia": hoje(),
    })
    salvar_erros(cards)


def atualizar_card(card_id: int, **campos) -> None:
    cards = carregar_erros()
    for c in cards:
        if c["id"] == card_id:
            c.update(campos)
            c["atualizado_em"] = agora()
            if c["coluna"] == COL_SOLUCIONADOS:
                c.setdefault("resolvido_em", agora())
            else:
                c.pop("resolvido_em", None)
    salvar_erros(cards)


FORMATO_DATA_HORA = "%d/%m/%Y %H:%M"


def historico_do_card(c: dict) -> list:
    """Linha do tempo de ocorrências. Cards antigos (sem histórico) ganham uma entrada inicial."""
    if c.get("historico"):
        return list(c["historico"])
    return [{
        "data": c.get("criado_em", agora()),
        "qtd": int(c["ocorrencias"]),
        "nota": "Registro inicial (anterior ao histórico)",
    }]


def registrar_ocorrencia(card_id: int, qtd: int, nota: str) -> None:
    """Soma ocorrências ao card e registra a data e a nota na linha do tempo."""
    cards = carregar_erros()
    for c in cards:
        if c["id"] == card_id:
            historico = historico_do_card(c)  # montado antes de alterar a contagem
            historico.append({"data": agora(), "qtd": int(qtd), "nota": nota.strip()})
            c["historico"] = historico
            c["ocorrencias"] = int(c["ocorrencias"]) + int(qtd)
            c["ultima_ocorrencia"] = hoje()
            c["atualizado_em"] = agora()
    salvar_erros(cards)


def excluir_card(card_id: int) -> None:
    restantes = [c for c in carregar_erros() if c["id"] != card_id]
    for c in restantes:  # remove as conexões que apontavam para o card excluído
        if card_id in c.get("conexoes", []):
            c["conexoes"] = [i for i in c["conexoes"] if i != card_id]
    salvar_erros(restantes)


@st.dialog("Registrar ocorrência")
def dialogo_ocorrencia(card_id: int) -> None:
    card = next((c for c in carregar_erros() if c["id"] == card_id), None)
    if card is None:
        st.warning("Este card não existe mais.")
        return
    st.markdown(f"**#{card_id} · {esc(card['titulo'])}**")
    nota = st.text_area(
        "O que aconteceu? (opcional)", height=110, key=f"oc_nota_{card_id}",
        placeholder="Ex.: cliente 123456, boleta 789, divergência de valor...",
    )
    qtd = st.number_input("Quantidade de ocorrências", min_value=1, value=1, step=1, key=f"oc_qtd_{card_id}")
    if st.button("Registrar", key=f"oc_ok_{card_id}"):
        registrar_ocorrencia(card_id, int(qtd), nota)
        st.toast("Ocorrência registrada.")
        st.rerun()


@st.dialog("Novo problema / gargalo", width="large")
def dialogo_novo() -> None:
    titulo = st.text_input("Título", key="novo_erro_titulo")
    descricao = st.text_area("Descrição", height=200, key="novo_erro_desc")
    c1, c2 = st.columns(2)
    prioridade = c1.selectbox("Prioridade", PRIORIDADES, index=1, key="novo_erro_prio")
    ocorrencias = c2.number_input("Ocorrências", min_value=1, value=1, step=1, key="novo_erro_oc")
    categoria = st.selectbox("Categoria", CATEGORIAS, key="novo_erro_cat")
    causa_raiz = st.text_area(
        "Causa raiz (opcional)", height=120, key="novo_erro_causa",
        placeholder="Por que o problema acontece, e não só o sintoma...",
    )
    if st.button("Criar", key="criar_erro"):
        if not titulo.strip():
            st.warning("Informe um título.")
            return
        criar_card(titulo, descricao, prioridade, ocorrencias, categoria, causa_raiz)
        st.toast("Problema registrado com sucesso.")
        st.rerun()


@st.dialog("Detalhes do problema", width="large")
def dialogo_card(card_id: int) -> None:
    card = next((c for c in carregar_erros() if c["id"] == card_id), None)
    if card is None:
        st.warning("Este card não existe mais.")
        return

    k = f"erro_{card_id}_"
    legenda = f"#{card_id} · criado em {card.get('criado_em', '—')} · atualizado em {card.get('atualizado_em', '—')}"
    if card.get("resolvido_em"):
        legenda += f" · resolvido em {card['resolvido_em']}"
    st.caption(legenda)

    aba_detalhes, aba_causa, aba_conexoes, aba_historico = st.tabs(["Detalhes", "Causa raiz", "Conexões", "Histórico"])

    with aba_detalhes:
        titulo = st.text_input("Título", value=card["titulo"], key=k + "titulo")
        descricao = st.text_area("Descrição completa", value=card["descricao"], height=180, key=k + "desc")
        c1, c2, c3 = st.columns(3)
        prioridade = c1.selectbox("Prioridade", PRIORIDADES, index=PRIORIDADES.index(card["prioridade"]), key=k + "prio")
        coluna = c2.selectbox("Status", COLUNAS_ERROS, index=COLUNAS_ERROS.index(card["coluna"]), key=k + "col")
        ocorrencias = c3.number_input("Ocorrências", min_value=1, value=int(card["ocorrencias"]), step=1, key=k + "oc")
        atual = categoria_do_card(card)
        opcoes = opcoes_categoria(atual)
        categoria = st.selectbox("Categoria", opcoes, index=opcoes.index(atual), key=k + "cat")
        solucao = st.text_area("Solução / como foi resolvido", value=card.get("solucao", ""), height=110, key=k + "sol")

    with aba_causa:
        st.caption("Registre por que o problema acontece, e não só o sintoma. Ex.: onde o processo falhou e o que o originou.")
        causa_raiz = st.text_area(
            "Causa raiz", value=card.get("causa_raiz", ""), height=300, key=k + "causa",
            placeholder="Descreva a causa raiz identificada...",
        )

    with aba_conexoes:
        st.caption("Ligue este problema aos que ele provoca ou agrava. As ligações aparecem no Mapa de Conexões.")
        outros = [c for c in carregar_erros() if c["id"] != card_id]
        rotulo = {c["id"]: f'#{c["id"]} · {c["titulo"]}' for c in outros}
        validos = [i for i in card.get("conexoes", []) if i in rotulo]
        conexoes = st.multiselect(
            "Este problema impacta:", list(rotulo), default=validos, format_func=rotulo.get,
            key=k + "conex", placeholder="Selecione os problemas impactados",
        )
        impactado_por = [c for c in outros if card_id in c.get("conexoes", [])]
        if impactado_por:
            st.markdown("**Impactado por:**\n" + "\n".join(f"- #{c['id']} · {c['titulo']}" for c in impactado_por))

    with aba_historico:
        st.caption("Cada ocorrência registrada, da mais recente para a mais antiga. Use o botão ＋1 do card para anotar uma nova.")
        for h in reversed(historico_do_card(card)):
            sinal = f"+{h['qtd']}" if int(h["qtd"]) > 0 else str(h["qtd"])
            nota = f" · {esc(h['nota'])}" if h.get("nota") else ""
            st.markdown(f"**{h['data']}** · {sinal} ocorrência(s){nota}")

    chave_exclusao = f"confirmar_exclusao_{card_id}"

    # Callbacks alteram o estado sem fechar o modal
    def iniciar_exclusao():
        st.session_state[chave_exclusao] = True

    def cancelar_exclusao():
        st.session_state[chave_exclusao] = False

    st.write("")

    if st.session_state.get(chave_exclusao, False):
        st.error("Tem certeza que deseja excluir este problema? Esta ação não pode ser desfeita.")
        c_sim, c_nao = st.columns(2)

        if c_sim.button("Sim, Excluir", use_container_width=True, key=k + "sim_excluir"):
            excluir_card(card_id)
            st.session_state[chave_exclusao] = False
            st.toast("Card excluído com sucesso.")
            st.rerun()

        c_nao.button("Cancelar", use_container_width=True, key=k + "nao_excluir", on_click=cancelar_exclusao)

    else:
        b_salvar, _, b_excluir = st.columns([1, 2.4, 1])

        if b_salvar.button("Salvar", key=k + "salvar"):
            if not titulo.strip():
                st.warning("O título não pode ficar vazio.")
            else:
                campos = dict(
                    titulo=titulo.strip(), descricao=descricao.strip(), prioridade=prioridade,
                    coluna=coluna, ocorrencias=int(ocorrencias), solucao=solucao.strip(),
                    categoria=categoria, causa_raiz=causa_raiz.strip(),
                    conexoes=[int(i) for i in conexoes],
                )
                delta = int(ocorrencias) - int(card["ocorrencias"])
                if delta > 0:
                    campos["ultima_ocorrencia"] = hoje()
                if delta != 0:  # ajuste manual da contagem também entra na linha do tempo
                    campos["historico"] = historico_do_card(card) + [
                        {"data": agora(), "qtd": delta, "nota": "Ajuste manual na edição do card"}
                    ]
                atualizar_card(card_id, **campos)
                st.toast("Card atualizado.")
                st.rerun()

        b_excluir.button("Excluir", key=k + "excluir", on_click=iniciar_exclusao)


def kpis_erros_html(cards: list) -> str:
    abertos = [c for c in cards if c["coluna"] != COL_SOLUCIONADOS]
    cartoes = [
        ("total", "Em aberto", len(abertos), "problemas e gargalos"),
        ("vazio", "Ocorrências", sum(int(c["ocorrencias"]) for c in abertos), "nos itens em aberto"),
        ("div", "Críticos", sum(1 for c in abertos if c["prioridade"] == "Crítica"), "em aberto"),
        ("ok", "Solucionados", len(cards) - len(abertos), "no total"),
    ]
    return '<div class="kpis">' + "".join(
        f'<div class="kpi {cls}"><div class="kpi-label">{rot}</div>'
        f'<div class="kpi-valor">{valor}</div><div class="kpi-sub">{sub_}</div></div>'
        for cls, rot, valor, sub_ in cartoes
    ) + "</div>"


DIAS_PARADO = 7  # a partir de quantos dias sem atualização o card ganha o alerta


def dias_sem_atualizar(c: dict):
    """Dias desde a última atualização (None para solucionados ou datas inválidas)."""
    if c["coluna"] == COL_SOLUCIONADOS:
        return None
    try:
        ultima = datetime.strptime(c.get("atualizado_em", ""), "%d/%m/%Y %H:%M")
    except ValueError:
        return None
    return (datetime.now() - ultima).days


def desenhar_card(c: dict) -> None:
    cid, n = c["id"], int(c["ocorrencias"])
    idx = COLUNAS_ERROS.index(c["coluna"])
    slug = SLUG_PRIORIDADE.get(c["prioridade"], "media")
    selo_causa = '<span class="badge causa">Causa raiz registrada</span>' if c.get("causa_raiz") else ""
    n_conex = len(c.get("conexoes", []))
    if n_conex:
        selo_causa += f'<span class="badge causa">Impacta {n_conex}</span>'
    parado = dias_sem_atualizar(c)
    if parado is not None and parado >= DIAS_PARADO:
        selo_causa += f'<span class="badge warn">Sem atualização há {parado} dias</span>'
    with st.container(key=f"cardErro_{slug}_{cid}"):
        st.markdown(
            f'<div class="erro-card-topo"><span class="erro-id">#{cid}</span>'
            f'<span class="badge prio-{slug}">{esc(c["prioridade"])}</span></div>'
            f'<div class="erro-titulo">{esc(c["titulo"])}</div>'
            f'<div class="erro-cat"><span class="badge cat">{esc(categoria_do_card(c))}</span>'
            f'{selo_causa}</div>'
            f'<div class="erro-meta">'
            f'<span class="destaque-ocorrencia">{n} ocorrência{"s" if n != 1 else ""}</span>'
            f' · última: {esc(c.get("ultima_ocorrencia", "—"))}</div>',
            unsafe_allow_html=True,
        )
        b_esq, b_dir, b_mais, b_abrir = st.columns([1, 1, 1, 1.7])
        if b_esq.button("◀", key=f"esq_{cid}", disabled=idx == 0, help="Mover para a coluna anterior"):
            atualizar_card(cid, coluna=COLUNAS_ERROS[idx - 1])
            st.toast("Card movido para a esquerda.")
            st.rerun()
        if b_dir.button("▶", key=f"dir_{cid}", disabled=idx == len(COLUNAS_ERROS) - 1, help="Mover para a próxima coluna"):
            atualizar_card(cid, coluna=COLUNAS_ERROS[idx + 1])
            st.toast("Card movido para a direita.")
            st.rerun()
        if b_mais.button("＋1", key=f"mais_{cid}", help="Registrar nova ocorrência (com anotação)"):
            dialogo_ocorrencia(cid)
        if b_abrir.button("Abrir", key=f"abrir_{cid}"):
            dialogo_card(cid)


def exibir_dashboard_erros(cards: list) -> None:
    if not cards:
        estado_vazio("Sem dados", "Nenhum problema registrado para gerar o dashboard.")
        return

    df = pd.DataFrame(cards)
    df["ocorrencias"] = pd.to_numeric(df["ocorrencias"], errors="coerce").fillna(0)
    if "categoria" not in df:
        df["categoria"] = SEM_CATEGORIA
    df["categoria"] = df["categoria"].replace("", SEM_CATEGORIA).fillna(SEM_CATEGORIA)
    df["Data"] = pd.to_datetime(df["criado_em"], format="%d/%m/%Y %H:%M", errors="coerce")

    st.write("")
    c1, c2 = st.columns(2)

    # --- Gráfico 1: Ocorrências por Prioridade (donut) ---
    fig_prio = px.pie(
        df, values="ocorrencias", names="prioridade",
        title="Volume de Ocorrências por Prioridade",
        color="prioridade", color_discrete_map=CORES_PRIORIDADE, hole=0.45,
    )
    estilizar_figura(fig_prio)
    c1.plotly_chart(fig_prio, use_container_width=True)

    # --- Gráfico 2: Top 5 Problemas ---
    top_problemas = df.nlargest(5, "ocorrencias")
    fig_top = px.bar(
        top_problemas, x="ocorrencias", y="titulo", orientation="h",
        title="Top 5 Gargalos Recorrentes",
        color="prioridade", color_discrete_map=CORES_PRIORIDADE, text="ocorrencias",
    )
    estilizar_figura(fig_top)
    fig_top.update_layout(yaxis={"categoryorder": "total ascending", "title": ""}, showlegend=False)
    c2.plotly_chart(fig_top, use_container_width=True)

    # --- Gráfico: Ocorrências por Categoria ---
    por_cat = df.groupby("categoria", as_index=False)["ocorrencias"].sum().sort_values("ocorrencias")
    fig_cat = px.bar(
        por_cat, x="ocorrencias", y="categoria", orientation="h",
        title="Ocorrências por Categoria", text="ocorrencias",
        color_discrete_sequence=[COR["vinho"]],
    )
    estilizar_figura(fig_cat)
    fig_cat.update_layout(yaxis_title="", xaxis_title="Nº de ocorrências")
    st.plotly_chart(fig_cat, use_container_width=True)

    st.write("")

    # --- Gráfico 3: Histórico (diário / mensal) ---
    granularidade = st.radio(
        "Agrupamento temporal:", ["Visão Diária", "Visão Mensal"], horizontal=True, key="gran_evolucao",
    )

    # Uma linha por ocorrência registrada (data real da ocorrência, não a de criação do card)
    eventos = pd.DataFrame([
        {
            "Data": pd.to_datetime(h.get("data"), format=FORMATO_DATA_HORA, errors="coerce"),
            "prioridade": c["prioridade"],
            "ocorrencias": int(h.get("qtd", 0)),
        }
        for c in cards for h in historico_do_card(c)
    ])

    if granularidade == "Visão Diária":
        eventos["Periodo"] = eventos["Data"].dt.strftime("%d/%m/%Y").fillna("Desconhecido")
        formato_data, rotulo_x = "%d/%m/%Y", "Data"
    else:
        eventos["Periodo"] = eventos["Data"].dt.strftime("%m/%Y").fillna("Desconhecido")
        formato_data, rotulo_x = "%m/%Y", "Mês / Ano"

    ev_tempo = eventos.groupby(["Periodo", "prioridade"])["ocorrencias"].sum().reset_index()
    ev_tempo["Data_Real"] = pd.to_datetime(ev_tempo["Periodo"], format=formato_data, errors="coerce")
    ev_tempo = ev_tempo.dropna(subset=["Data_Real"]).sort_values("Data_Real")

    fig_linha = px.line(
        ev_tempo, x="Periodo", y="ocorrencias", color="prioridade",
        title=f"Evolução de Ocorrências ({granularidade.replace('Visão ', '')})",
        color_discrete_map=CORES_PRIORIDADE, markers=True,
        labels={"Periodo": rotulo_x, "ocorrencias": "Nº de Ocorrências"},
    )
    estilizar_figura(fig_linha)
    fig_linha.update_layout(
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title_text=""),
        hovermode="x unified",
    )
    fig_linha.update_traces(line=dict(width=3), marker=dict(size=8, line=dict(width=1, color="white")))
    st.plotly_chart(fig_linha, use_container_width=True)


def layout_grafo(nos: list, arestas: list, iteracoes: int = 250) -> dict:
    """Posiciona os nós (força de repulsão entre todos + atração nas ligações), em [-1, 1]."""
    import numpy as np

    n = len(nos)
    indice = {no: i for i, no in enumerate(nos)}
    angulos = np.linspace(0, 2 * np.pi, n, endpoint=False)
    pos = np.column_stack([np.cos(angulos), np.sin(angulos)])
    pos += np.random.default_rng(7).normal(scale=0.05, size=pos.shape)  # evita nós sobrepostos

    k = 1.6 / np.sqrt(max(n, 1))
    pares = np.array([(indice[a], indice[b]) for a, b in arestas], dtype=int).reshape(-1, 2)
    temperatura = 0.2
    for _ in range(iteracoes):
        delta = pos[:, None, :] - pos[None, :, :]
        dist = np.maximum(np.linalg.norm(delta, axis=-1), 0.01)
        forca = (delta * (k**2 / dist**2)[..., None]).sum(axis=1)  # repulsão
        for a, b in pares:  # atração
            d = pos[a] - pos[b]
            f = d * (np.linalg.norm(d) / k)
            forca[a] -= f
            forca[b] += f
        tamanho = np.maximum(np.linalg.norm(forca, axis=1), 0.01)
        pos += forca / tamanho[:, None] * np.minimum(tamanho, temperatura)[:, None]
        temperatura = max(temperatura * 0.98, 0.002)

    pos -= pos.mean(axis=0)
    pos /= max(np.abs(pos).max(), 1e-9)
    return {no: (float(pos[i, 0]), float(pos[i, 1])) for no, i in indice.items()}


def exibir_grafo_erros(cards: list) -> None:
    """Mapa de conexões: seta A -> B significa que o problema A impacta o problema B."""
    import math

    import plotly.graph_objects as go

    c1, c2 = st.columns(2)
    ocultar_solucionados = c1.checkbox("Ocultar solucionados", value=True, key="grafo_ocultar_sol")
    so_conectados = c2.checkbox("Mostrar só problemas conectados", value=True, key="grafo_so_conect")

    base = [c for c in cards if not (ocultar_solucionados and c["coluna"] == COL_SOLUCIONADOS)]
    ids = {c["id"] for c in base}
    arestas = [(c["id"], d) for c in base for d in c.get("conexoes", []) if d in ids and d != c["id"]]

    if not arestas:
        estado_vazio(
            "Nenhuma conexão registrada",
            "Abra um card, vá na aba Conexões e selecione os problemas que ele impacta.",
        )
        return

    if so_conectados:
        conectados = {i for par in arestas for i in par}
        base = [c for c in base if c["id"] in conectados]

    por_id = {c["id"]: c for c in base}
    arestas = list(dict.fromkeys(arestas))  # sem ligações repetidas
    saidas = {i: 0 for i in por_id}
    entradas = {i: 0 for i in por_id}
    for origem, destino in arestas:
        saidas[origem] += 1
        entradas[destino] += 1
    pos = layout_grafo(list(por_id), arestas)

    fig = go.Figure()

    # Setas (paradas um pouco antes do nó de destino para não ficarem escondidas atrás dele)
    for origem, destino in arestas:
        x0, y0 = pos[origem]
        x1, y1 = pos[destino]
        dx, dy = x1 - x0, y1 - y0
        dist = math.hypot(dx, dy) or 1.0
        recuo = 0.075
        fig.add_annotation(
            x=x1 - dx / dist * recuo, y=y1 - dy / dist * recuo, ax=x0 + dx / dist * recuo, ay=y0 + dy / dist * recuo,
            xref="x", yref="y", axref="x", ayref="y", text="", showarrow=True,
            arrowhead=3, arrowsize=1.3, arrowwidth=1.6, arrowcolor="#A58A95",
        )

    # Nós: um traço por prioridade (gera a legenda); tamanho = ocorrências
    for prio in PRIORIDADES:
        nos = [c for c in base if c["prioridade"] == prio]
        if not nos:
            continue
        fig.add_trace(go.Scatter(
            x=[pos[c["id"]][0] for c in nos],
            y=[pos[c["id"]][1] for c in nos],
            mode="markers+text",
            name=prio,
            text=[f'#{c["id"]} {c["titulo"][:26]}{"…" if len(c["titulo"]) > 26 else ""}' for c in nos],
            textposition="bottom center",
            textfont=dict(size=11, color=COR["texto"]),
            marker=dict(
                size=[min(18 + 5 * math.sqrt(max(int(c["ocorrencias"]), 1)), 60) for c in nos],
                color=CORES_PRIORIDADE[prio],
                opacity=[0.4 if c["coluna"] == COL_SOLUCIONADOS else 1 for c in nos],
                line=dict(width=2, color="white"),
            ),
            customdata=[[
                c["id"], c["titulo"], categoria_do_card(c), c["coluna"], int(c["ocorrencias"]),
                saidas[c["id"]], entradas[c["id"]],
            ] for c in nos],
            hovertemplate=(
                "<b>#%{customdata[0]} · %{customdata[1]}</b><br>Categoria: %{customdata[2]}<br>"
                "Status: %{customdata[3]}<br>Ocorrências: %{customdata[4]}<br>"
                "Impacta: %{customdata[5]} · Impactado por: %{customdata[6]}<extra></extra>"
            ),
        ))

    estilizar_figura(fig, "Mapa de Conexões entre Problemas")
    fig.update_xaxes(visible=False, range=[-1.35, 1.35])
    fig.update_yaxes(visible=False, range=[-1.35, 1.35])
    fig.update_layout(
        height=640, legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1, title_text=""),
        hoverlabel=dict(font_family="Inter, sans-serif"),
    )
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Seta A → B: o problema A impacta o problema B. O tamanho do círculo representa o número de ocorrências.")

    maior_origem = max(por_id, key=lambda n: saidas[n])
    mais_impactado = max(por_id, key=lambda n: entradas[n])
    k1, k2 = st.columns(2)
    k1.markdown(
        f'<div class="kpi"><div class="kpi-label">Maior origem de impacto</div>'
        f'<div class="kpi-sub"><b>#{maior_origem} · {esc(por_id[maior_origem]["titulo"])}</b><br>'
        f'impacta {saidas[maior_origem]} problema(s)</div></div>',
        unsafe_allow_html=True,
    )
    k2.markdown(
        f'<div class="kpi div"><div class="kpi-label">Mais impactado</div>'
        f'<div class="kpi-sub"><b>#{mais_impactado} · {esc(por_id[mais_impactado]["titulo"])}</b><br>'
        f'é impactado por {entradas[mais_impactado]} problema(s)</div></div>',
        unsafe_allow_html=True,
    )


def carregar_snapshot_ata() -> dict:
    if not ARQUIVO_SNAPSHOT.exists():
        return {}
    try:
        return json.loads(ARQUIVO_SNAPSHOT.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def salvar_snapshot_ata(dados: dict) -> None:
    ARQUIVO_SNAPSHOT.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


@st.dialog("Relatório para Ata de Reunião", width="large")
def dialogo_ata():
    cards = carregar_erros()
    snapshot_antigo = carregar_snapshot_ata()

    abertos = [c for c in cards if c["coluna"] != COL_SOLUCIONADOS]
    abertos.sort(key=lambda c: -int(c["ocorrencias"]))

    linhas_relatorio = ["### Resumo de Problemas e Gargalos Pendentes\n"]
    novo_snapshot = {}

    for c in abertos:
        cid = str(c["id"])
        ocorrencias_atuais = int(c["ocorrencias"])
        ocorrencias_antigas = snapshot_antigo.get(cid, ocorrencias_atuais)

        delta = ocorrencias_atuais - ocorrencias_antigas
        novo_snapshot[cid] = ocorrencias_atuais

        texto_aumento = f" (+{delta} desde a última reunião)" if delta > 0 else ""
        linhas_relatorio.append(
            f"- **{c['titulo']}** [{categoria_do_card(c)}]: {c['descricao']} | "
            f"{ocorrencias_atuais} ocorrência(s){texto_aumento}."
        )
        if c.get("causa_raiz"):
            linhas_relatorio.append(f"  - Causa raiz: {c['causa_raiz']}")

    texto_final = "\n".join(linhas_relatorio)

    st.markdown("Copie o texto abaixo e cole diretamente na sua Ata. O cálculo de aumento é feito com base na última vez que guardou a referência.")
    st.text_area("Texto gerado:", value=texto_final, height=300, label_visibility="collapsed")
    st.write("")

    st.info("Clique no botão abaixo apenas no final da reunião. Isso irá 'zerar' o contador de aumento para a próxima ata.")
    if st.button("Guardar estado atual (Atualizar referência)", use_container_width=True):
        salvar_snapshot_ata(novo_snapshot)
        st.toast("Referência da Ata atualizada com sucesso.")
        st.success("Referência atualizada com sucesso! A partir de agora, o aumento será calculado face a este momento.")


def aba_controle_erros() -> None:
    fazer_backup_diario()
    st.markdown(CSS_ERROS, unsafe_allow_html=True)
    cards = carregar_erros()
    st.markdown(kpis_erros_html(cards), unsafe_allow_html=True)

    modo = st.radio(
        "Modo de visualização",
        ["Quadro Kanban", "Dashboard de Métricas", "Mapa de Conexões"],
        horizontal=True,
        label_visibility="collapsed",
    )

    if modo == "Mapa de Conexões":
        exibir_grafo_erros(cards)
        return

    if modo == "Dashboard de Métricas":
        exibir_dashboard_erros(cards)
        return

    c_novo, c_ata, c_busca, c_prio, c_cat = st.columns([1.2, 1.2, 2.2, 1.2, 1.6])

    if c_novo.button("＋ Novo problema", key="btn_novo_erro"):
        dialogo_novo()

    if c_ata.button("Gerar Ata", key="btn_gerar_ata"):
        dialogo_ata()

    busca = c_busca.text_input(
        "Buscar", placeholder="Buscar por título ou descrição",
        label_visibility="collapsed", key="busca_erros",
    )
    filtro_prio = c_prio.selectbox(
        "Prioridade", ["Todas"] + PRIORIDADES, label_visibility="collapsed", key="filtro_prio_erros",
    )

    filtro_cat = c_cat.selectbox(
        "Categoria", ["Todas as categorias"] + CATEGORIAS, label_visibility="collapsed", key="filtro_cat_erros",
    )

    termo = normalizar(busca)
    visiveis = [
        c for c in cards
        if (filtro_prio == "Todas" or c["prioridade"] == filtro_prio)
        and (filtro_cat == "Todas as categorias" or categoria_do_card(c) == filtro_cat)
        and (not termo or termo in normalizar(f'{c["titulo"]} {c["descricao"]} {c.get("causa_raiz", "")}'))
    ]

    colunas = st.columns(len(COLUNAS_ERROS))
    for coluna, nome in zip(colunas, COLUNAS_ERROS):
        itens = [c for c in visiveis if c["coluna"] == nome]
        itens.sort(key=lambda c: (-PRIORIDADES.index(c["prioridade"]), -int(c["ocorrencias"])))
        with coluna:
            st.markdown(
                f'<div class="erro-col-topo"><span>{esc(nome)}</span>'
                f'<span class="erro-col-qtd">{len(itens)}</span></div>',
                unsafe_allow_html=True,
            )
            if not itens:
                st.markdown('<div class="erro-vazio">Nenhum card</div>', unsafe_allow_html=True)
            for c in itens:
                desenhar_card(c)


# --- FIM CONTROLE DE ERROS --------------------------------------------------

# ----------------------------------------------------------------------------
# Abas
# ----------------------------------------------------------------------------
aba_conferencia, aba_liquidacao, aba_erros = st.tabs(
    ["Conferência de Boletas", "Controle de Liquidação", "Controle de Erros"]
)

with aba_conferencia:
    area_etapas = st.empty()  # preenchida depois, quando já se sabe o que foi enviado

    # ---- 1. Importar a base da TAG (boletador) ----
    st.markdown('<div class="secao">1 · Base da TAG</div>'
                '<div class="secao-sub">Importe a base mais recente do boletador ou envie a planilha abaixo.</div>',
                unsafe_allow_html=True)
    aviso = st.empty()

    if st.button("Importar base mais recente", key="btn_importar"):
        try:
            with st.spinner("Navegador aberto. Aguardando a importação da base..."):
                resultado = subprocess.run(
                    [sys.executable, str(Path(__file__).resolve().with_name("importador.py"))],
                    capture_output=True,
                    text=True,
                    timeout=TEMPO_MAX_LOGIN + 300,
                )
        except subprocess.TimeoutExpired:
            st.error("A importação demorou mais do que o esperado e foi interrompida.")
            resultado = None

        if resultado is not None:
            if resultado.returncode == 0:
                caminho = next(
                    (
                        linha.split(MARCADOR_ARQUIVO, 1)[1].strip()
                        for linha in resultado.stdout.splitlines()
                        if linha.startswith(MARCADOR_ARQUIVO)
                    ),
                    None,
                )
                if caminho:
                    try:
                        st.session_state["df_base"] = ler_tag_cache(Path(caminho).read_bytes(), Path(caminho).name)
                        st.session_state["arquivo_base"] = caminho

                        aviso.success("Planilha baixada com sucesso")
                        time.sleep(3)
                        aviso.empty()
                    except Exception as e:
                        aviso.error(f"A planilha foi baixada, mas não consegui lê-la: {e}")
                else:
                    aviso.warning("Download concluído, mas o caminho do arquivo não foi retornado.")
            elif resultado.returncode == 2:
                st.warning("Tempo esgotado: o login não foi concluído em 5 minutos.")
            elif resultado.returncode == 3:
                st.warning("Houve um erro na execução ou o navegador foi fechado antes de terminar.")
                st.code((resultado.stderr or resultado.stdout)[-1500:] or "Erro detalhado não encontrado.")
            else:
                st.error("Não foi possível abrir o navegador.")
                st.code((resultado.stderr or resultado.stdout)[-1500:] or "sem detalhes")

    # ---- 2. Planilhas de entrada (upload) ----
    st.markdown("---")
    st.markdown('<div class="secao">2 · Planilhas de entrada</div>'
                '<div class="secao-sub">A planilha da TAG é opcional se a base já foi importada.</div>',
                unsafe_allow_html=True)
    col_tag, col_xp = st.columns(2)
    with col_tag:
        arq_tag = st.file_uploader(
            "Planilha da TAG (opcional: use se já tiver baixado)",
            type=["xlsx", "xls", "csv"],
            key="up_tag",
        )
    with col_xp:
        arq_xp = st.file_uploader("Planilha da XP", type=["xlsx", "xls"], key="up_xp")

    if arq_tag is not None:
        st.caption(f"TAG: usando a planilha enviada ({arq_tag.name})")
    elif "df_base" in st.session_state:
        st.caption(f"TAG: usando a base importada ({Path(st.session_state['arquivo_base']).name})")
    else:
        st.caption("TAG: importe a base ou envie a planilha da TAG.")

    comparar = st.checkbox(
        "Conferir também Operação e Valor (além da conta)",
        value=True,
        key="chk_op_valor",
    )

    # ---- 3. RUN ----
    tem_tag = arq_tag is not None or "df_base" in st.session_state
    tem_xp = arq_xp is not None
    faltando = [nome for nome, ok in (("a base da TAG", tem_tag), ("a planilha da XP", tem_xp)) if not ok]

    st.markdown("---")
    if st.button("Conciliar (RUN)", key="btn_run", disabled=bool(faltando)):
        try:
            df_tag = (
                ler_tag_cache(arq_tag.getvalue(), arq_tag.name)
                if arq_tag is not None
                else st.session_state["df_base"]
            )
            df_xp = ler_xp_cache(arq_xp.getvalue(), arq_xp.name)
            resultado_run = rodar_conciliacao(df_tag, df_xp, comparar)
            resultado_run["excel"] = gerar_excel(resultado_run)
            st.session_state["run"] = resultado_run
            # Abre já nas divergentes, que é o que precisa de ação
            tem_div = any(
                (resultado_run[chave][coluna] == "DIVERGENTE").any()
                for chave, coluna in (("tag", COL_CONC_TAG), ("xp", COL_CONC_XP))
            )
            st.session_state["filtro_conc"] = "DIVERGENTE" if tem_div else "Todas"
        except Exception as e:
            st.session_state.pop("run", None)
            st.error(f"Não foi possível fazer a conciliação: {e}")
    if faltando:
        st.caption("Para conciliar, falta: " + " e ".join(faltando) + ".")

    area_etapas.markdown(etapas_html([tem_tag, tem_xp, "run" in st.session_state]), unsafe_allow_html=True)

    # ---- 4. Resultado (fica na tela enquanto você mexe nos filtros) ----
    if "run" in st.session_state:
        r = st.session_state["run"]
        modo = "Conta + Operação + Valor" if r.get("comparar", True) else "Somente pela conta"

        st.markdown("---")
        st.markdown(
            f'<div class="res-titulo"><h3>Resultado da conciliação</h3><span class="chip-modo">{modo}</span></div>',
            unsafe_allow_html=True,
        )
        area_kpis = st.container()

        with st.expander(f"Contas da TAG (Nome Cliente + Conta Cliente): {len(r['contas'])} registros"):
            st.markdown(tabela_html(r["contas"], ["Nome Cliente", "Conta Cliente"]), unsafe_allow_html=True)

        c_visao, c_filtro = st.columns([1, 1.7])
        with c_visao:
            st.markdown('<div class="rotulo">Visão</div>', unsafe_allow_html=True)
            visao = st.radio("Visão", ["TAG → XP", "XP → TAG"], horizontal=True, key="visao_run", label_visibility="collapsed")

        if visao == "TAG → XP":
            res, col_conc = r["tag"], COL_CONC_TAG
            colunas = [
                "Código da Boleta", "Nome Cliente", "Cliente", "Conta Cliente", "Operação", "Valor Solicitado",
                "Status", "Status TAG", "Conta na XP", COL_CONC_TAG, "Motivo",
            ]
        else:
            res, col_conc = r["xp"], COL_CONC_XP
            colunas = [
                "Código da operação", "Cliente", "Código do Cliente", "Nome Cliente (TAG)", "Operação", "Fundo",
                "Valor Solicitado", "Status", "Conta na TAG", COL_CONC_XP, "Motivo",
            ]

        total = len(res)
        n_ok = int((res[col_conc] == "CONCILIADO").sum())
        n_div = int((res[col_conc] == "DIVERGENTE").sum())
        n_vazio = total - n_ok - n_div
        with area_kpis:
            st.markdown(kpis_html(total, n_ok, n_div, n_vazio), unsafe_allow_html=True)

        rotulos = {"Todas": f"Todas ({total})", "CONCILIADO": f"Conciliadas ({n_ok})", "DIVERGENTE": f"Divergentes ({n_div})"}
        with c_filtro:
            st.markdown('<div class="rotulo">Mostrar</div>', unsafe_allow_html=True)
            filtro = st.radio(
                "Mostrar", list(rotulos), format_func=rotulos.get, horizontal=True,
                key="filtro_conc", label_visibility="collapsed",
            )

        mostrar = res if filtro == "Todas" else res[res[col_conc] == filtro]
        if mostrar.empty:
            st.info("Nenhuma linha neste filtro.")
        else:
            st.markdown(tabela_html(mostrar, colunas, col_conc), unsafe_allow_html=True)

        st.download_button(
            "Baixar conciliação (Excel)",
            data=r["excel"],
            file_name="conciliacao_tag_xp.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with aba_liquidacao:
    estado_vazio("Controle de Liquidação", "Em breve.")

with aba_erros:
    aba_controle_erros()
