import base64
import html
import json
import os
import re
import shutil
import time
import unicodedata
from collections import Counter, OrderedDict
from datetime import date, datetime, timedelta
from functools import lru_cache

import pandas as pd
import pythoncom
import streamlit as st
import win32com.client as win32

try:
    from pypdf import PdfReader  # opcional: pip install pypdf
except ImportError:
    PdfReader = None

try:
    from watchdog.events import FileSystemEventHandler  # opcional: pip install watchdog
    from watchdog.observers import Observer
except ImportError:
    Observer = FileSystemEventHandler = None

# Pasta do próprio app.py: não depende de onde o comando foi executado
BASE = os.path.dirname(os.path.abspath(__file__))
ARQUIVO = os.path.join(BASE, "Template_Fundos_Ativos.xlsx")
EMAIL_RE = re.compile(r"^[^@\s;,]+@[^@\s;,]+\.[^@\s;,]+$")

# ==========================================
# ESTILO
# ==========================================
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root{--vinho:#6A0F2B;--vinho-escuro:#4A0A1E;--vinho-medio:#8A1B3D;--vinho-claro:#F4E8EC;
--fundo:#F7F4F5;--card:#FFFFFF;--texto:#26151B;--suave:#7A6770;--borda:#E6DCE0;--dourado:#B8924A;
--ok:#2E8B57;--ok-bg:#DDF1E5;--ok-tx:#1D6B40;--div:#C0392B;--div-bg:#F4D3D8;--div-tx:#8E1B2E;
--alerta:#C98A1B;--alerta-bg:#FBEBC8;--alerta-tx:#8A5A00;--sombra:0 2px 8px rgba(74,10,30,.07);}
html,body,.stApp{font-family:"Inter",system-ui,sans-serif;color-scheme:light;}
.stApp{background:var(--fundo);color:var(--texto);}
#MainMenu,footer,[data-testid="stToolbar"],[data-testid="stDecoration"]{display:none !important;}


.kpi{background:var(--card);border:1px solid var(--borda);border-top:4px solid var(--vinho);border-radius:12px;
box-shadow:var(--sombra);padding:.9rem 1.1rem;}
.kpi-val{font-size:2.1rem;font-weight:700;line-height:1.1;color:var(--texto);}
.kpi-lbl{color:var(--suave);font-size:.78rem;font-weight:600;letter-spacing:.06em;text-transform:uppercase;margin-top:.3rem;}
.kpi.ok{border-top-color:var(--ok);}
.kpi.warn{border-top-color:var(--alerta);}
.kpi.warn .kpi-val{color:var(--div-tx);}

.st-key-card_envio,.st-key-card_lista,.st-key-card_conf{background:var(--card) !important;border:1px solid var(--borda) !important;
border-radius:12px;box-shadow:var(--sombra);}

[data-testid="stWidgetLabel"] *,[data-testid="stCheckbox"] p{color:var(--texto) !important;}
[data-testid="stTooltipIcon"] svg{color:var(--suave) !important;}
[data-testid="stAlert"]{border-radius:10px;}
[data-baseweb="input"],[data-baseweb="base-input"],[data-baseweb="input"]>div{background:#fff !important;border-radius:10px;}
[data-baseweb="input"]{border:1px solid var(--borda) !important;}
[data-baseweb="input"]:focus-within{border-color:var(--vinho) !important;box-shadow:0 0 0 2px rgba(106,15,43,.18);}
[data-baseweb="input"] input{background:#fff !important;color:var(--texto) !important;-webkit-text-fill-color:var(--texto);}
[data-baseweb="input"] input::placeholder{color:var(--suave) !important;-webkit-text-fill-color:var(--suave);opacity:1;}
label[data-baseweb="checkbox"]:not(:has(input:checked))>span:first-child{background:#fff !important;border:2px solid #C9A3B1 !important;}
label[data-baseweb="checkbox"]:has(input:checked)>span:first-child{background:var(--vinho) !important;border-color:var(--vinho) !important;}
.card-title{font-weight:700;font-size:1.08rem;color:var(--vinho);}
.card-hint{color:var(--suave);font-size:.88rem;margin:.1rem 0 .6rem;}

[data-testid="stHorizontalBlock"]{align-items:center;}
.colhead{color:var(--suave);font-size:.78rem;font-weight:600;letter-spacing:.05em;text-transform:uppercase;}
.sep{border-top:1px solid #F0E8EB;margin:.3rem 0;}
.fundo{font-weight:600;}
.ativo{color:var(--suave);font-size:.92rem;}
.badge{display:inline-block;padding:.2rem .7rem;border-radius:999px;font-size:.8rem;font-weight:600;}
.badge.ok{background:var(--ok-bg);color:var(--ok-tx);}
.badge.sem{background:var(--div-bg);color:var(--div-tx);}
.badge.pend{background:var(--alerta-bg);color:var(--alerta-tx);}

.stButton>button{border-radius:10px;font-weight:600;border:1px solid var(--borda);color:var(--vinho);background:#fff;}
.stButton>button:hover{border-color:var(--vinho);color:var(--vinho);background:var(--vinho-claro);}
.stButton>button[kind="primary"],.stButton>button[data-testid="stBaseButton-primary"]{
background:var(--vinho);border-color:var(--vinho);color:#fff;}
.stButton>button[kind="primary"]:hover,.stButton>button[data-testid="stBaseButton-primary"]:hover{
background:var(--vinho-medio);border-color:var(--vinho-medio);color:#fff;box-shadow:0 4px 12px rgba(74,10,30,.25);}
.stButton>button:focus-visible{outline:3px solid rgba(184,146,74,.7);outline-offset:2px;}
.stButton>button:disabled{background:#E3D8DC !important;color:var(--suave) !important;border-color:#E3D8DC !important;opacity:1 !important;cursor:not-allowed;}

.tbl-wrap{max-height:540px;overflow:auto;margin-top:.6rem;border:1px solid var(--borda);border-radius:12px;}
.tbl{width:100%;border-collapse:collapse;font-size:.93rem;}
.tbl th{position:sticky;top:0;background:var(--vinho);color:#fff;text-align:left;font-size:.78rem;font-weight:600;
letter-spacing:.05em;text-transform:uppercase;padding:.6rem .7rem;border-bottom:2px solid var(--dourado);}
.tbl td{padding:.65rem .7rem;border-bottom:1px solid #F0E8EB;vertical-align:middle;}
.tbl tbody tr:nth-child(even){background:#FBF8F9;}
.tbl tbody tr:hover{background:var(--vinho-claro);}
.tbl tr:last-child td{border-bottom:none;}
.tbl tr.novo td{border-top:2px solid #E6DCE0;}
.tbl tr.cont td{border-bottom-color:#F7F1F3;}

[data-testid="stHorizontalBlock"]:has(.st-key-card_envio){align-items:flex-start;}
.st-key-card_envio .stButton>button[kind="primary"],.st-key-card_envio .stButton>button[data-testid="stBaseButton-primary"]{
padding:.7rem 1rem;font-size:1.02rem;}
.vazio{background:var(--card);border:1px dashed #D8C3CB;border-radius:12px;padding:1.4rem;text-align:center;color:var(--suave);}
.vazio b{color:var(--vinho);display:block;margin-bottom:.2rem;font-size:1.02rem;}

[data-testid="stHorizontalBlock"]:has(.row-warn){background:#FDF3F4;box-shadow:inset 4px 0 0 var(--div);border-radius:8px;padding:.15rem .4rem .15rem .7rem;}

[data-testid="stAlert"]{background:#FFF8EA !important;border:1px solid #F0DDB0;border-left:4px solid var(--dourado);color:var(--texto);}
[data-testid="stAlert"] *{color:var(--texto) !important;}

[data-testid="stRadio"] [role="radiogroup"]{gap:.4rem;flex-wrap:wrap;}
[data-testid="stRadio"] label{background:#fff;border:1px solid var(--borda);border-radius:999px;padding:.15rem .9rem;margin:0;cursor:pointer;}
[data-testid="stRadio"] label:hover{border-color:var(--vinho);background:var(--vinho-claro);}
[data-testid="stRadio"] label:has(input:checked){background:var(--vinho);border-color:var(--vinho);}
[data-testid="stRadio"] label:has(input:checked) *{color:#fff !important;}
[data-testid="stRadioOption"]>div>div:first-child:not([data-testid="stMarkdownContainer"]){display:none !important;}
.badge{white-space:nowrap;}
.tbl td.fundo{font-weight:600;}
@media (max-width:640px){.topo .marca{font-size:1.4rem;}.block-container{padding:1rem .8rem 2rem;}}
</style>
"""


# ==========================================
# COLUNAS DA PLANILHA (nomes tolerantes a maiúsculas, acentos e espaços)
# ==========================================
ALIASES = {
    "fundo": {"nome do fundo", "fundo"},
    "cotista": {"cotistas", "cotista", "nome do cotista"},
    "doc": {"cpf/cnpj cotistas", "cpf/cnpj cotista", "cpf/cnpj", "cpf cnpj cotistas"},
    "email": {"emails", "email", "e-mail", "e-mails"},
    "cnpj_fundo": {"cnpj do fundo", "cnpj fundo"},
}
OBRIGATORIAS = {"fundo": "Nome do Fundo", "cotista": "COTISTAS", "email": "emails"}

# ==========================================
# CONFIGURAÇÕES DE ENVIO OCULTAS
# ==========================================
# Coloque os e-mails em cópia aqui. Deixe vazio ("") se não quiser ninguém em cópia.
EMAILS_EM_COPIA = "processamento@taginvest.com.br; bo@taginvest.com.br"

# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================
def norm(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().casefold()


def saudacao():
    hora = datetime.now().hour
    return "bom dia" if hora < 12 else "boa tarde" if hora < 18 else "boa noite"


def _pascoa(ano):
    """Domingo de Páscoa (algoritmo gregoriano anônimo)."""
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes, dia = divmod(h + l - 7 * m + 114, 31)
    return date(ano, mes, dia + 1)


def feriados_nacionais(ano):
    """Feriados nacionais (fixos e móveis) do ano: {data: nome}. Estaduais e municipais não entram."""
    p = _pascoa(ano)
    lista = {
        date(ano, 1, 1): "Confraternização Universal",
        p - timedelta(days=48): "Carnaval",
        p - timedelta(days=47): "Carnaval",
        p - timedelta(days=2): "Sexta-feira Santa",
        date(ano, 4, 21): "Tiradentes",
        date(ano, 5, 1): "Dia do Trabalho",
        p + timedelta(days=60): "Corpus Christi",
        date(ano, 9, 7): "Independência do Brasil",
        date(ano, 10, 12): "Nossa Senhora Aparecida",
        date(ano, 11, 2): "Finados",
        date(ano, 11, 15): "Proclamação da República",
        date(ano, 12, 25): "Natal",
    }
    if ano >= 2024:
        lista[date(ano, 11, 20)] = "Consciência Negra"
    return lista


def motivo_nao_util(d):
    """Por que a data não é dia útil ('sábado', 'domingo' ou nome do feriado); None se for dia útil."""
    if d.weekday() == 5:
        return "um sábado"
    if d.weekday() == 6:
        return "um domingo"
    feriado = feriados_nacionais(d.year).get(d)
    return f"feriado ({feriado})" if feriado else None


def dia_util_anterior(d):
    d -= timedelta(days=1)
    while motivo_nao_util(d):
        d -= timedelta(days=1)
    return d


def texto(valor):
    return "" if pd.isna(valor) else str(valor).strip()


def lista_emails(valor):
    """Separa por vírgula, ';' ou quebra de linha e remove repetidos (sem diferenciar maiúsculas)."""
    if pd.isna(valor):
        return []
    vistos, saida = set(), []
    for p in re.split(r"[;,\r\n]+", str(valor)):
        p = p.strip()
        if p and p.lower() not in vistos:
            vistos.add(p.lower())
            saida.append(p)
    return saida


@st.cache_data
def carregar_planilha(caminho, mtime):
    # dtype=str preserva zeros à esquerda de CPF/CNPJ
    df = pd.read_excel(caminho, dtype=str).dropna(how="all")
    originais = [str(c) for c in df.columns]
    mapa = {}
    for col in df.columns:
        n = norm(col)
        for campo, nomes in ALIASES.items():
            if n in nomes and campo not in mapa.values():
                mapa[col] = campo
    return df.rename(columns=mapa), originais


def montar_itens(df):
    """Um item por destinatário: linhas com o mesmo e-mail (ou o mesmo conjunto de e-mails)
    viram um único e-mail com todos os fundos e cotistas."""
    itens = OrderedDict()
    for n, (_, r) in enumerate(df.iterrows()):
        fundo, cotista = texto(r["fundo"]), texto(r["cotista"])
        doc = texto(r["doc"]) if "doc" in df.columns else ""
        cnpj = texto(r["cnpj_fundo"]) if "cnpj_fundo" in df.columns else ""
        if not (fundo or cotista):
            continue
        emails = lista_emails(r["email"])
        chave = tuple(sorted(e.lower() for e in emails)) if emails else ("__sem_email__", n)
        item = itens.setdefault(
            chave,
            {"emails": "; ".join(emails), "fundos": OrderedDict(), "cnpjs": {}, "linhas": 0,
             "invalidos": [e for e in emails if not EMAIL_RE.match(e)]},
        )
        item["linhas"] += 1
        if cnpj:
            item["cnpjs"].setdefault(fundo, cnpj)
        cots = item["fundos"].setdefault(fundo, [])
        if cotista and (cotista, doc) not in cots:
            cots.append((cotista, doc))
    return list(itens.values())


def nome_item(item):
    fundos = [f for f in item["fundos"] if f]
    if not fundos:
        return ""
    return fundos[0] if len(fundos) == 1 else f"{fundos[0]} e mais {len(fundos) - 1}"


def montar_assunto(item, data_fmt):
    nome = nome_item(item)
    return f"Extrato Mensal Ref. {data_fmt}" + (f" - {nome}" if nome else "")


def montar_corpo_html(item, data_fmt):
    """Um bloco por fundo: frase de pedido + tabela (Nome Cliente / CPF/CNPJ)."""
    fonte = "font-family:Arial,Helvetica,sans-serif;font-size:10pt;"
    borda = "border:1px solid #000000;padding:2px 5px;"
    blocos = []
    for fundo, cots in item["fundos"].items():
        cnpj = item["cnpjs"].get(fundo, "")
        ativo = f"<b>{html.escape(fundo)}</b>" if fundo else ""
        if cnpj:
            ativo += f" <b>({html.escape(cnpj)})</b>"
        pedido = "os extratos dos cotistas abaixo" if len(cots) > 1 else "o extrato do cotista abaixo"
        frase = f"Por favor, poderiam nos enviar {pedido}, ref. <b>{data_fmt}</b>"
        if ativo:
            frase += f", no ativo {ativo}"
        frase += "?"

        linhas = "".join(
            f'<tr><td style="{borda}">{html.escape(c) or "-"}</td>'
            f'<td style="{borda}">{html.escape(d) or "-"}</td></tr>'
            for c, d in (cots or [("", "")])
        )
        tabela = (
            '<table border="1" cellspacing="0" cellpadding="2" width="720" '
            f'style="border-collapse:collapse;border:1px solid #000000;{fonte}">'
            f'<tr><th align="left" width="75%" style="{borda}">Nome Cliente</th>'
            f'<th align="left" width="25%" style="{borda}">CPF/CNPJ</th></tr>'
            f"{linhas}</table>"
        )
        blocos.append(f"<p>{frase}</p>{tabela}<br>Fico no aguardo!<br>")

    return (
        f'<div style="{fonte}">'
        f"<p>Prezados, {saudacao()}! Tudo bem?</p>"
        + "<br>".join(blocos)
        + "<br></div>"
    )


def inserir_no_corpo(html_existente, corpo):
    """Coloca o texto logo após <body>, preservando a assinatura do Outlook."""
    m = re.search(r"<body[^>]*>", html_existente, flags=re.IGNORECASE)
    if m:
        return html_existente[: m.end()] + corpo + html_existente[m.end():]
    return corpo + html_existente


def abrir_emails_outlook(itens, data_ref, enviar=False):
    data_fmt = data_ref.strftime("%d/%m")
    abertos = 0
    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("outlook.application")
        for item in itens:
            try:
                mail = outlook.CreateItem(0)
                mail.To = item["emails"]

                # Insere os e-mails em cópia (CC)
                if EMAILS_EM_COPIA:
                    mail.CC = EMAILS_EM_COPIA

                mail.Subject = montar_assunto(item, data_fmt)
                mail.Display()  # carrega a assinatura padrão
                for _ in range(10):  # espera a assinatura carregar antes de inserir o texto
                    try:
                        if re.search(r"<body", mail.HTMLBody, flags=re.IGNORECASE):
                            break
                    except Exception:
                        pass
                    time.sleep(0.2)
                mail.HTMLBody = inserir_no_corpo(mail.HTMLBody, montar_corpo_html(item, data_fmt))
                if enviar:
                    mail.Send()
                abertos += 1
            except Exception as e:
                st.error(f"Não foi possível abrir o e-mail para {item['emails']}: {e}")
    except Exception as e:
        st.error(f"Não foi possível conectar ao Outlook. Abra o Outlook e tente de novo. Detalhe: {e}")
    finally:
        pythoncom.CoUninitialize()
    return abertos


PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
NDR_MARCAS = (
    "undeliverable", "delivery status notification", "mail delivery", "returned mail",
    "não foi possível entregar", "nao foi possivel entregar", "falha na entrega", "não entregue",
)


def _smtp(rec):
    try:
        v = rec.PropertyAccessor.GetProperty(PR_SMTP)
        if v:
            return str(v).lower()
    except Exception:
        pass
    try:
        return str(rec.Address).lower()
    except Exception:
        return ""


def _destinatarios(msg):
    """Endereços SMTP dos destinatários 'Para' da mensagem."""
    try:
        return {_smtp(r) for r in msg.Recipients if r.Type == 1} - {""}
    except Exception:
        return set()


def _recentes(pasta, campo, desde):
    """Itens da pasta com data >= desde, do mais novo para o mais antigo."""
    itens = pasta.Items
    itens.Sort(f"[{campo}]", True)
    achados = []
    for m in itens:
        try:
            quando = getattr(m, campo).replace(tzinfo=None)
        except Exception:
            continue
        if quando < desde:
            break
        achados.append(m)
    return achados


def conferir_envios(itens, desde):
    """Confere no Outlook o que aconteceu com cada e-mail aberto. Retorna {indice: status}."""
    desde = desde - timedelta(minutes=1)
    status = {i: "Aguardando envio" for i in range(len(itens))}
    pythoncom.CoInitialize()
    try:
        ns = win32.Dispatch("outlook.application").GetNamespace("MAPI")
        enviados = []
        for m in _recentes(ns.GetDefaultFolder(5), "SentOn", desde):
            try:
                enviados.append((_destinatarios(m), str(m.Subject).lower(), m.SentOn.replace(tzinfo=None)))
            except Exception:
                continue
        saida = [(_destinatarios(m), str(m.Subject).lower()) for m in ns.GetDefaultFolder(4).Items]
        devolvidos = []
        for m in _recentes(ns.GetDefaultFolder(6), "ReceivedTime", desde):
            try:
                if any(k in str(m.Subject).lower() for k in NDR_MARCAS):
                    devolvidos.append((m.ReceivedTime.replace(tzinfo=None), str(m.Body).lower()))
            except Exception:
                continue

        for i, it in enumerate(itens):
            esperados = {e.lower() for e in lista_emails(it["emails"])}
            casados = [env for dest, assunto, env in enviados if esperados <= dest and "extrato" in assunto]
            if casados:
                status[i] = "Enviado"
                # só vale devolução recebida depois do envio mais antigo deste e-mail
                primeiro = min(casados)
                if any(rec >= primeiro and e in corpo for rec, corpo in devolvidos for e in esperados):
                    status[i] = "Erro de entrega"
            elif any(esperados <= dest and "extrato" in assunto for dest, assunto in saida):
                status[i] = "Na caixa de saída"
        return status
    except Exception as e:
        st.error(f"Não foi possível conferir no Outlook: {e}")
        return None
    finally:
        pythoncom.CoUninitialize()


PASTA_EXTRATOS = os.path.join(BASE, "Extratos")
PR_ANEXO_OCULTO = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
PREFIXO_REVISAR = "_Revisar"  # anexos sem fundo identificado; também reclassificados a cada conferência


def _nome_seguro(txt):
    limpo = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(txt))
    return re.sub(r"\s+", " ", limpo).strip(" .") or "sem_nome"


def _assunto_base(assunto):
    """Tira prefixos de resposta/encaminhamento (RE:, RES:, ENC:, FW:...)."""
    return re.sub(r"^\s*((re|res|enc|fw|fwd|rv)\s*:\s*)+", "", str(assunto), flags=re.I).strip().lower()


def _remetente(msg):
    try:
        if str(msg.SenderEmailType).upper() == "SMTP":
            return str(msg.SenderEmailAddress).lower()
        return str(msg.Sender.GetExchangeUser().PrimarySmtpAddress).lower()
    except Exception:
        return ""


def _anexo_valido(a):
    """Só arquivos de verdade: ignora itens embutidos, anexos ocultos e logos de assinatura."""
    try:
        if a.Type != 1:
            return False
        try:
            if a.PropertyAccessor.GetProperty(PR_ANEXO_OCULTO):
                return False
        except Exception:
            pass
        nome = str(a.FileName).lower()
        if nome.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp")) and int(a.Size) < 50000:
            return False
        return True
    except Exception:
        return False


# ==========================================
# CLASSIFICAÇÃO DO ANEXO POR FUNDO
# ==========================================
def _so_digitos(s):
    return re.sub(r"\D", "", str(s))


def _limpo(s):
    """norm() tratando _ - . como espaço, para casar nomes de arquivo."""
    return norm(re.sub(r"[_\-.]+", " ", str(s)))


def _contem(texto_limpo, termo):
    """O termo aparece como palavra(s) inteira(s) no texto (já passado por _limpo)?"""
    termo = _limpo(termo)
    if not termo:
        return False
    return re.search(rf"(?<!\w){re.escape(termo)}(?!\w)", texto_limpo) is not None


# palavras genéricas de nomes de fundos/cotistas: não servem para identificar ninguém
PALAVRAS_COMUNS = {
    "fic", "fim", "fif", "fidc", "fia", "fip", "fii", "fi", "cic", "resp", "ltda", "limitada", "sub", "sr",
    "cp", "lp", "rl", "de", "da", "do", "das", "dos", "e", "ou", "mult", "multimercado", "cred", "credito",
    "priv", "privado", "acoes", "ie", "ii", "iii", "iv", "fundo", "investimento", "investimentos", "cotas",
    "cotista", "cotistas", "extrato", "posicao", "perfil", "previdencia", "prev", "institucional",
}


def _tokens(nome):
    """Palavras que distinguem o nome (sem termos genéricos como FIC, FIM, RESP LTDA)."""
    return [p for p in _limpo(nome).split() if p not in PALAVRAS_COMUNS and (len(p) >= 3 or p.isdigit())]


def _forca_nome(t, nome):
    """Quanto do nome aparece no texto: 1.0 = nome completo; fração das palavras-chave quando abreviado
    (ex.: 'Previ Novartis' para 'PREVI NOVARTIS FIF CIC MULT RESP LTDA'); 0 = não aparece."""
    if not nome:
        return 0.0
    if _contem(t, nome):
        return 1.0
    toks = _tokens(nome)
    if not toks:
        return 0.0
    achou = sum(1 for p in toks if _contem(t, p))
    if len(toks) == 1:
        return 0.5 if achou else 0.0
    return achou / len(toks) if achou >= 2 and achou / len(toks) >= 0.5 else 0.0


def _pontuar(item, fonte):
    """Pontos por fundo: nome/CNPJ do fundo valem 10 (nome abreviado vale proporcional);
    nome/documento do cotista valem 3 (abreviado, proporcional)."""
    if not fonte:
        return {}
    t, d = _limpo(fonte), _so_digitos(fonte)
    pontos = {}
    for fundo, cots in item["fundos"].items():
        if not fundo:
            continue
        p = 0
        if _contem(t, fundo):
            p += 10
        else:
            # nome do fundo abreviado no arquivo (ex.: 'Root' para 'Root Capital'): conta cada palavra-chave
            p += 4 * sum(1 for w in _tokens(fundo) if _contem(t, w))
        cnpj = _so_digitos(item["cnpjs"].get(fundo, ""))
        if len(cnpj) >= 11 and cnpj in d:
            p += 10
        for cot, doc in cots:
            p += 3 * _forca_nome(t, cot)
            dd = _so_digitos(doc)
            if len(dd) >= 11 and dd in d:
                p += 3
        if p:
            pontos[fundo] = p
    return pontos


def _vencedor(pontos):
    """Fundo com mais pontos, só se não houver empate."""
    ordem = sorted(pontos.items(), key=lambda x: -x[1])
    if not ordem:
        return None
    if len(ordem) == 1 or ordem[0][1] > ordem[1][1]:
        return ordem[0][0]
    return None


@lru_cache(maxsize=512)
def _texto_pdf_cache(caminho, mtime, tamanho, max_paginas=3):
    try:
        return " ".join((p.extract_text() or "") for p in PdfReader(caminho).pages[:max_paginas])
    except Exception:
        return ""


def _texto_pdf(caminho, max_paginas=3):
    """Texto das primeiras páginas do PDF (com cache por arquivo, para não reler a cada conferência)."""
    if PdfReader is None or not caminho.lower().endswith(".pdf"):
        return ""
    try:
        st_ = os.stat(caminho)
    except OSError:
        return ""
    return _texto_pdf_cache(caminho, st_.st_mtime_ns, st_.st_size, max_paginas)


def _cotistas_cobertos(item, nome_arquivo, caminho, so_fundo=None):
    """Cotistas (fundo, nome) do item identificados no anexo, pelo nome do arquivo ou conteúdo do PDF.
    Nome completo ou documento: todos contam. Só nome abreviado: conta apenas o que casa melhor
    (se houver empate, conta um só, para não inflar o total)."""
    fonte = f"{nome_arquivo} {_texto_pdf(caminho)}"
    t, d = _limpo(fonte), _so_digitos(fonte)
    exatos, parciais = set(), []
    for fundo, cots in item["fundos"].items():
        if so_fundo and fundo != so_fundo:
            continue
        for cot, doc in cots:
            dd = _so_digitos(doc)
            forca = _forca_nome(t, cot) if cot else 0.0
            if forca >= 1.0 or (len(dd) >= 11 and dd in d):
                exatos.add((fundo, cot))
            elif forca > 0:
                parciais.append((forca, fundo, cot))
    if exatos or not parciais:
        return exatos
    melhor = max(p[0] for p in parciais)
    return {next((f, c) for forca, f, c in parciais if forca == melhor)}


def _escolher_item(itens, candidatos, nome_arquivo, caminho):
    """Entre vários itens do mesmo remetente, escolhe o que melhor casa com o conteúdo do anexo."""
    if len(candidatos) == 1:
        return candidatos[0]
    fonte = f"{nome_arquivo} {_texto_pdf(caminho)}"
    notas = {k: sum(_pontuar(itens[k], fonte).values()) for k in candidatos}
    melhor = max(notas.values())
    if melhor > 0:
        vencedores = [k for k in candidatos if notas[k] == melhor]
        if len(vencedores) == 1:
            return vencedores[0]
    return candidatos[0]


def _fundo_identificado(item, nome_arquivo, msg, caminho):
    """Fundo do anexo somando evidências do nome do arquivo, do conteúdo do PDF, do assunto e do corpo
    (peso maior para as duas primeiras). None se não der para identificar com segurança."""
    fundos = [f for f in item["fundos"] if f]
    if len(fundos) == 1:
        return fundos[0]

    def corpo():
        try:
            return str(msg.Body)
        except Exception:
            return ""

    def assunto():
        try:
            return str(msg.Subject)
        except Exception:
            return ""

    total = {}
    for peso, fonte in ((3, nome_arquivo), (3, _texto_pdf(caminho)), (1, assunto()), (0.5, corpo())):
        for fundo, p in _pontuar(item, fonte).items():
            total[fundo] = total.get(fundo, 0) + p * peso
    return _vencedor(total)


def _pasta_sem_fundo(msg):
    """Pasta de revisão para anexos que não foram associados a nenhum fundo."""
    try:
        remetente = re.sub(r"\s+", " ", str(msg.SenderName)).strip()
    except Exception:
        remetente = ""
    return f"{PREFIXO_REVISAR} - {remetente}" if remetente else f"{PREFIXO_REVISAR} - remetente desconhecido"


def _fundo_do_anexo(item, nome_arquivo, msg, caminho):
    """Nome da pasta do anexo: o fundo identificado ou, na falta dele, a pasta de revisão."""
    return _fundo_identificado(item, nome_arquivo, msg, caminho) or _pasta_sem_fundo(msg)


def _em_pasta_sem_fundo(caminho):
    pasta = os.path.basename(os.path.dirname(caminho))
    return pasta.startswith((PREFIXO_REVISAR, "Remetente"))  # "Remetente" = nome antigo da pasta


def _mover_para_pasta(origem, pasta_mes, nome_pasta, nome, evitar_duplicata=False):
    """Move o arquivo para <pasta_mes>/<nome_pasta>/ sem sobrescrever nada. Com evitar_duplicata,
    se a pasta já tem um arquivo de mesmo nome e tamanho, descarta a cópia e aponta para o existente."""
    pasta = os.path.join(pasta_mes, _nome_seguro(nome_pasta))
    os.makedirs(pasta, exist_ok=True)
    base, ext = os.path.splitext(nome)
    destino, n = os.path.join(pasta, nome), 1
    while os.path.exists(destino):
        if evitar_duplicata and os.path.getsize(destino) == os.path.getsize(origem):
            os.remove(origem)
            return destino
        destino = os.path.join(pasta, f"{base} ({n}){ext}")
        n += 1
    shutil.move(origem, destino)
    return destino


def conferir_respostas(itens, desde, data_ref, pasta_mes):
    """Procura respostas com anexo na Caixa de Entrada e salva os arquivos em
    Extratos/AAAA-MM/<Fundo>/. Retorna ({indice: n_cotistas_recebidos}, n_salvos_agora, falhas) ou None."""
    desde = desde - timedelta(minutes=1)
    data_fmt = data_ref.strftime("%d/%m")
    esperados = [{e.lower() for e in lista_emails(it["emails"])} for it in itens]
    assuntos = [_assunto_base(montar_assunto(it, data_fmt)) for it in itens]
    repetidos = Counter(assuntos)

    os.makedirs(pasta_mes, exist_ok=True)
    manifesto_path = os.path.join(pasta_mes, "_salvos.json")
    try:
        with open(manifesto_path, encoding="utf-8") as f:
            manifesto = json.load(f)
    except Exception:
        manifesto = {}

    cobertos, sem_ident, falhas, novos = {}, {}, [], 0
    pythoncom.CoInitialize()
    try:
        ns = win32.Dispatch("outlook.application").GetNamespace("MAPI")
        for m in _recentes(ns.GetDefaultFolder(6), "ReceivedTime", desde):
            try:
                if m.Class != 43 or m.Attachments.Count == 0:
                    continue
                remetente, assunto = _remetente(m), _assunto_base(m.Subject)
            except Exception:
                continue

            # ignora resposta que cita outra data de referência (ex.: extrato de outro mês)
            datas = re.findall(r"\b\d{2}/\d{2}\b", assunto)
            if datas and data_fmt not in datas:
                continue

            # todos os itens do remetente; se houver mais de um, o conteúdo do anexo desempata
            candidatos = [k for k, e in enumerate(esperados) if remetente in e and "extrato" in assunto]
            if not candidatos:
                candidatos = [k for k, a in enumerate(assuntos) if a == assunto and repetidos[a] == 1]
            if not candidatos:
                continue

            for a in m.Attachments:
                if not _anexo_valido(a):
                    continue
                nome = _nome_seguro(a.FileName)
                chave = f"{m.EntryID}|{nome}|{a.Size}"
                tmp_salvo = None
                try:
                    if chave in manifesto and os.path.exists(manifesto[chave]):
                        caminho, ja_salvo = manifesto[chave], True
                    else:
                        # salva numa pasta temporária; só depois de identificar o fundo vai para o destino
                        tmp = os.path.join(pasta_mes, "_tmp")
                        os.makedirs(tmp, exist_ok=True)
                        tmp_salvo = caminho = os.path.join(tmp, f"{os.getpid()}_{nome}")
                        a.SaveAsFile(caminho)
                        ja_salvo = False

                    i = _escolher_item(itens, candidatos, nome, caminho)

                    if not ja_salvo:
                        fundo = _fundo_do_anexo(itens[i], nome, m, caminho)
                        destino = _mover_para_pasta(caminho, pasta_mes, fundo, nome)
                        tmp_salvo = None
                        manifesto[chave] = destino
                        caminho = destino
                        novos += 1
                    elif _em_pasta_sem_fundo(caminho):
                        # já salvo numa pasta de revisão: tenta de novo com as regras atuais
                        fundo = _fundo_identificado(itens[i], nome, m, caminho)
                        if fundo:
                            destino = _mover_para_pasta(caminho, pasta_mes, fundo, nome, evitar_duplicata=True)
                            manifesto[chave] = caminho = destino
                except Exception as e:
                    falhas.append(f"{nome}: {e}")
                    if tmp_salvo and os.path.exists(tmp_salvo):
                        try:
                            os.remove(tmp_salvo)
                        except OSError:
                            pass
                    continue

                # só conta o que foi realmente salvo; conta por cotista identificado no anexo
                pasta_fundo = os.path.basename(os.path.dirname(caminho))  # fundo em que o anexo foi guardado
                achados = _cotistas_cobertos(itens[i], nome, caminho, pasta_fundo if pasta_fundo in itens[i]["fundos"] else None)
                if achados:
                    cobertos.setdefault(i, set()).update(achados)
                else:
                    sem_ident[i] = sem_ident.get(i, 0) + 1
        with open(manifesto_path, "w", encoding="utf-8") as f:
            json.dump(manifesto, f, ensure_ascii=False, indent=1)

        recebidos = {}
        for i in set(cobertos) | set(sem_ident):
            n_cot = sum(len(c) for c in itens[i]["fundos"].values())
            recebidos[i] = max(1, min(n_cot, len(cobertos.get(i, ())) + sem_ident.get(i, 0)))
        return recebidos, novos, falhas
    except Exception as e:
        st.error(f"Não foi possível procurar as respostas no Outlook: {e}")
        return None
    finally:
        pythoncom.CoUninitialize()


def fazer_conferencia(rodada, pasta_mes):
    """Consulta o Outlook (envios e respostas) e guarda o resultado na sessão. True se deu certo."""
    status = conferir_envios(rodada["itens"], rodada["inicio"])
    resp = conferir_respostas(rodada["itens"], rodada["inicio"], rodada["data"], pasta_mes) if status is not None else None
    if status is None:
        return False
    st.session_state["conferencia"] = {
        "status": status,
        "recebidos": resp[0] if resp else {},
        "falhas": resp[2] if resp else [],
        "hora": datetime.now(),
    }
    if resp and resp[1]:
        st.toast(f"{resp[1]} arquivo(s) salvo(s) na pasta Extratos.")
    return True


UNIDADES_REENVIO = {"horas": 1, "dias": 24, "semanas": 168}


def reenviar_sem_resposta(rodada, pasta_mes, intervalo):
    """Confere o Outlook e reenvia (de fato envia) os e-mails já enviados que seguem sem extrato
    há mais que o intervalo. Retorna quantos foram reenviados, ou None se a conferência falhou."""
    if not fazer_conferencia(rodada, pasta_mes):
        return None
    conf = st.session_state["conferencia"]
    recebidos = dict(conf["recebidos"])
    for k, v in contar_na_pasta(rodada["itens"], pasta_mes).items():
        recebidos[k] = max(recebidos.get(k, 0), v)
    agora = datetime.now()
    reenvios = rodada.setdefault("reenvios", {})
    devidos = []
    for i, it in enumerate(rodada["itens"]):
        n_cot = sum(len(c) for c in it["fundos"].values())
        if conf["status"].get(i) != "Enviado" or recebidos.get(i, 0) >= n_cot:
            continue
        if agora - reenvios.get(i, rodada["inicio"]) >= intervalo:
            devidos.append(i)
    if not devidos:
        return 0
    n = abrir_emails_outlook([rodada["itens"][i] for i in devidos], rodada["data"], enviar=True)
    for i in devidos:
        reenvios[i] = agora
    try:
        salvar_rodada(rodada)
    except OSError:
        pass
    return n


def contar_na_pasta(itens, pasta_mes):
    """{indice: cotistas recebidos} a partir dos arquivos nas pastas dos fundos, inclusive os colocados à mão."""
    recebidos = {}
    for i, it in enumerate(itens):
        cobertos, sem_ident = set(), 0
        for fundo in it["fundos"]:
            pasta = os.path.join(pasta_mes, _nome_seguro(fundo)) if fundo else ""
            if not pasta or not os.path.isdir(pasta):
                continue
            for nome in sorted(os.listdir(pasta)):
                caminho = os.path.join(pasta, nome)
                if not os.path.isfile(caminho):
                    continue
                achados = _cotistas_cobertos(it, nome, caminho, fundo)
                if achados:
                    cobertos |= achados
                else:
                    sem_ident += 1
        total = len(cobertos) + sem_ident
        if total:
            n_cot = sum(len(c) for c in it["fundos"].values())
            recebidos[i] = max(1, min(n_cot, total))
    return recebidos


@st.cache_resource
def iniciar_vigia(pasta):
    """Vigia a pasta do mês (watchdog) e conta os eventos de arquivo; a tela compara o contador para se atualizar."""
    estado = {"n": 0}
    if Observer is None:
        return estado
    os.makedirs(pasta, exist_ok=True)

    class _Vigia(FileSystemEventHandler):
        def _tocou(self, ev):
            caminho = str(getattr(ev, "dest_path", "") or ev.src_path)
            if ev.is_directory or f"{os.sep}_tmp{os.sep}" in caminho or caminho.endswith("_salvos.json"):
                return
            estado["n"] += 1

        on_created = on_moved = on_deleted = _tocou

    try:
        obs = Observer()
        obs.daemon = True
        obs.schedule(_Vigia(), pasta, recursive=True)
        obs.start()
    except Exception:
        pass
    return estado


def _dir_rodadas():
    return os.path.join(PASTA_EXTRATOS, "_rodadas")


def salvar_rodada(rodada, encerrada=False):
    """Guarda a rodada em disco para retomar a conferência em outro dia."""
    os.makedirs(_dir_rodadas(), exist_ok=True)
    # data+hora de abertura no nome: refazer a rodada da mesma data não sobrescreve a anterior
    caminho = rodada.get("arquivo") or os.path.join(
        _dir_rodadas(), f"{rodada['data']:%Y-%m-%d}_{rodada['inicio']:%H%M%S}.json"
    )
    dados = {
        "inicio": rodada["inicio"].isoformat(),
        "data": rodada["data"].isoformat(),
        "encerrada": encerrada,
        "reenvios": {str(k): v.isoformat() for k, v in rodada.get("reenvios", {}).items()},
        "itens": [
            {
                "emails": it["emails"],
                "linhas": it["linhas"],
                "cnpjs": it["cnpjs"],
                "fundos": [[f, [list(c) for c in cots]] for f, cots in it["fundos"].items()],
            }
            for it in rodada["itens"]
        ],
    }
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=1)
    rodada["arquivo"] = caminho


def carregar_ultima_rodada():
    """Retorna a rodada aberta mais recente (ou None) e quantas outras rodadas abertas existem."""
    pasta = _dir_rodadas()
    if not os.path.isdir(pasta):
        return None, 0
    arquivos = sorted(
        (os.path.join(pasta, n) for n in os.listdir(pasta) if n.endswith(".json")),
        key=os.path.getmtime,
        reverse=True,
    )
    escolhida, outras = None, 0
    for caminho in arquivos:
        try:
            with open(caminho, encoding="utf-8") as f:
                d = json.load(f)
            if d.get("encerrada"):
                continue
            if escolhida is not None:
                outras += 1
                continue
            itens = [
                {
                    "emails": i["emails"],
                    "linhas": i["linhas"],
                    "cnpjs": i["cnpjs"],
                    "fundos": OrderedDict((fu, [tuple(c) for c in cots]) for fu, cots in i["fundos"]),
                }
                for i in d["itens"]
            ]
            for it in itens:
                it["invalidos"] = [e for e in lista_emails(it["emails"]) if not EMAIL_RE.match(e)]
            escolhida = {
                "inicio": datetime.fromisoformat(d["inicio"]),
                "data": datetime.fromisoformat(d["data"]).date(),
                "itens": itens,
                "arquivo": caminho,
                "reenvios": {int(k): datetime.fromisoformat(v) for k, v in d.get("reenvios", {}).items()},
            }
        except Exception:
            continue
    return escolhida, outras


TEMA = """[theme]
base = "light"
primaryColor = "#6A0F2B"
backgroundColor = "#F7F4F5"
secondaryBackgroundColor = "#FFFFFF"
textColor = "#26151B"
font = "sans serif"

[client]
toolbarMode = "minimal"
"""


def garantir_tema():
    """Cria .streamlit/config.toml (tema claro) na pasta de execução, se ainda não existir."""
    pasta = os.path.join(os.getcwd(), ".streamlit")
    caminho = os.path.join(pasta, "config.toml")
    if os.path.exists(caminho):
        return False
    try:
        os.makedirs(pasta, exist_ok=True)
        with open(caminho, "w", encoding="utf-8") as f:
            f.write(TEMA)
        return True
    except OSError:
        return False


def kpi(rotulo, valor, alerta=False, tipo=""):
    cls = "kpi warn" if alerta and valor > 0 else f"kpi {tipo}".strip()
    return f'<div class="{cls}"><div class="kpi-val">{valor}</div><div class="kpi-lbl">{rotulo}</div></div>'


# ==========================================
# INTERFACE
# ==========================================
# set_page_config e tema ficam no portal (home.py)
st.markdown(CSS, unsafe_allow_html=True)

data = st.session_state.get("data_ref")

import sys  # noqa: E402
from pathlib import Path  # noqa: E402

_raiz = str(Path(__file__).resolve().parents[2])
if _raiz not in sys.path:
    sys.path.insert(0, _raiz)

from tema import topo  # noqa: E402  (tema compartilhado, na raiz do portal)

topo("Solicitação de extratos", "Gere os e-mails de extratos direto no Outlook e envie", app="extratos")

if not os.path.exists(ARQUIVO):
    st.error(f"Não encontrei '{os.path.basename(ARQUIVO)}'. Coloque a planilha na mesma pasta do app.py e recarregue a página.")
    st.stop()

mtime = os.path.getmtime(ARQUIVO)
df, originais = carregar_planilha(ARQUIVO, mtime)

faltando = [rotulo for campo, rotulo in OBRIGATORIAS.items() if campo not in df.columns]
if faltando:
    st.error(f"Faltam colunas na planilha: {', '.join(faltando)}.")
    st.caption(f"Colunas encontradas: {', '.join(originais)}")
    st.stop()

itens = montar_itens(df)
validos = [i for i in itens if i["emails"]]
sem_email = len(itens) - len(validos)
linhas_total = sum(i["linhas"] for i in itens)

k1, k2, k3 = st.columns(3)
k1.markdown(kpi("Linhas na planilha", linhas_total), unsafe_allow_html=True)
k2.markdown(kpi("E-mails a enviar", len(validos)), unsafe_allow_html=True)
k3.markdown(kpi("Sem e-mail cadastrado", sem_email, alerta=True), unsafe_allow_html=True)
st.write("")

com_invalido = [i for i in validos if i.get("invalidos")]
if com_invalido:
    st.warning(
        f"{len(com_invalido)} destinatário(s) com e-mail de formato inválido: "
        + "; ".join(e for i in com_invalido for e in i["invalidos"][:3])
        + ". Corrija na planilha antes de enviar."
    )

esq, dir_ = st.columns([1, 2.3], gap="medium")

# ---------- Painel de envio ----------
with esq:
    with st.container(border=True, key="card_envio"):
        st.markdown(
            '<div class="card-title">Preparar envio</div>'
            '<div class="card-hint">A data aparece no assunto e no texto de cada e-mail. '
            "Destinatários repetidos são unificados: cada um recebe um único e-mail com todos os fundos e cotistas.</div>",
            unsafe_allow_html=True,
        )
        st.date_input("Data de referência", value=None, format="DD/MM/YYYY", key="data_ref")
        motivo = motivo_nao_util(data) if data else None
        if motivo:
            anterior = dia_util_anterior(data)
            st.warning(f"{data:%d/%m/%Y} cai em {motivo}. Confira se a data de referência está certa.")
            st.button(
                f"Usar o dia útil anterior ({anterior:%d/%m/%Y})",
                key="usar_util",
                use_container_width=True,
                on_click=lambda d=anterior: st.session_state.update(data_ref=d),
            )
        st.write("")

        if st.button(
            f"Abrir todos ({len(validos)})",
            type="primary",
            use_container_width=True,
            disabled=data is None or not validos,
        ):
            st.session_state["confirmar_todos"] = True

        if data is None:
            st.caption("Escolha a data para liberar os botões de envio.")
        elif validos:
            st.caption(f"{len(validos)} e-mails serão abertos no Outlook, um por destinatário. Você revisa e envia cada um.")

        if st.session_state.get("confirmar_todos") and not (data and validos):
            st.session_state["confirmar_todos"] = False

        if st.session_state.get("confirmar_todos"):
            aviso = f"O Outlook vai abrir {len(validos)} janelas de e-mail."
            if sem_email:
                aviso += f" {sem_email} sem e-mail cadastrado ficam de fora."
            st.warning(aviso)
            b1, b2 = st.columns(2)
            if b1.button("Sim, abrir", key="confirma", type="primary", use_container_width=True):
                st.session_state["confirmar_todos"] = False
                inicio = datetime.now()
                abertos = abrir_emails_outlook(validos, data)
                if abertos:
                    st.session_state["rodada"] = {"inicio": inicio, "data": data, "itens": validos}
                    salvar_rodada(st.session_state["rodada"])
                    st.session_state.pop("conferencia", None)
                    st.toast(f"{abertos} e-mails abertos no Outlook.")
            if b2.button("Cancelar", key="cancela", use_container_width=True):
                st.session_state["confirmar_todos"] = False
                st.rerun()

# ---------- Lista ----------
with dir_:
    with st.container(border=True, key="card_lista"):
        st.markdown('<div class="card-title">Fundos e cotistas</div>', unsafe_allow_html=True)
        busca = st.text_input(
            "Buscar",
            placeholder="Busque por fundo ou cotista",
            label_visibility="collapsed",
        ).strip().lower()

        registros = []
        for it in itens:
            for fundo, cots in it["fundos"].items():
                for cotista, _doc in (cots or [("", "")]):
                    registros.append({"fundo": fundo, "cotista": cotista, "ok": bool(it["emails"])})

        visiveis = [r for r in registros if not busca or busca in f"{r['fundo']} {r['cotista']}".lower()]

        if not visiveis:
            st.info("Nenhum resultado. Apague o texto da busca para ver a lista completa.")
        else:
            linhas, anterior = [], None
            for r in visiveis:
                status = (
                    '<span class="badge ok">E-mail cadastrado</span>'
                    if r["ok"]
                    else '<span class="badge sem">Sem e-mail</span>'
                )
                # o nome do fundo aparece só na primeira linha de cada grupo
                novo = r["fundo"] != anterior
                anterior = r["fundo"]
                nome_fundo = (html.escape(r["fundo"]) or "-") if novo else ""
                linhas.append(
                    f'<tr class="{"novo" if novo else "cont"}"><td class="fundo">{nome_fundo}</td>'
                    f'<td>{html.escape(r["cotista"]) or "-"}</td><td>{status}</td></tr>'
                )
            st.markdown(
                '<div class="tbl-wrap"><table class="tbl"><thead><tr>'
                "<th>Fundo</th><th>Cotista</th><th>Status</th></tr></thead><tbody>"
                + "".join(linhas)
                + "</tbody></table></div>",
                unsafe_allow_html=True,
            )

        st.caption(f"Planilha {os.path.basename(ARQUIVO)}, atualizada em {datetime.fromtimestamp(mtime):%d/%m/%Y %H:%M}.")

# ---------- Conferência de envios e respostas ----------
if "rodada" not in st.session_state and not st.session_state.get("rodada_carregada"):
    st.session_state["rodada_carregada"] = True
    _retomada, _outras = carregar_ultima_rodada()
    if _retomada:
        st.session_state["rodada"] = _retomada
        st.session_state["rodadas_outras"] = _outras

rodada = st.session_state.get("rodada")
if not rodada:
    st.write("")
    st.markdown(
        '<div class="vazio"><b>Conferência de envios e respostas</b>'
        "Ela aparece aqui depois que você abrir os e-mails no Outlook.</div>",
        unsafe_allow_html=True,
    )
if rodada:
    st.write("")
    pasta_mes = os.path.join(PASTA_EXTRATOS, rodada["data"].strftime("%Y-%m"))
    vigia = iniciar_vigia(pasta_mes)
    conf = st.session_state.get("conferencia")
    pend = []
    if conf:
        # soma ao que veio do Outlook os arquivos que estão nas pastas (inclusive os colocados à mão)
        chave_disco = (pasta_mes, vigia["n"], conf["hora"])
        if st.session_state.get("disco_chave") != chave_disco:
            st.session_state["disco"] = contar_na_pasta(rodada["itens"], pasta_mes)
            st.session_state["disco_chave"] = chave_disco
        efetivos = dict(conf["recebidos"])
        for k, v in st.session_state["disco"].items():
            efetivos[k] = max(efetivos.get(k, 0), v)
        conf = {**conf, "recebidos": efetivos}
        pend = [
            i for i, stt in conf["status"].items()
            if stt in ("Aguardando envio", "Erro de entrega") and i not in conf["recebidos"]
        ]

    with st.container(border=True, key="card_conf"):
        st.markdown(
            '<div class="card-title">Conferência de envios e respostas</div>'
            f'<div class="card-hint">Rodada de referência {rodada["data"]:%d/%m/%Y}, aberta em '
            f'{rodada["inicio"]:%d/%m às %H:%M}. Depois de enviar, clique em conferir: o app verifica os envios '
            "e salva os extratos recebidos por resposta.</div>",
            unsafe_allow_html=True,
        )
        b1, b2, b3, b4 = st.columns(4)
        if b1.button("Conferir envios e respostas", key="conferir", type="primary", use_container_width=True):
            with st.spinner("Consultando o Outlook..."):
                ok = fazer_conferencia(rodada, pasta_mes)
            if ok:
                st.rerun()

        if b2.button(f"Reabrir pendentes ({len(pend)})", key="reabrir", disabled=not pend, use_container_width=True):
            if abrir_emails_outlook([rodada["itens"][i] for i in pend], rodada["data"]):
                st.toast("E-mails pendentes abertos no Outlook.")

        if b3.button("Abrir pasta dos extratos", key="pasta", use_container_width=True):
            try:
                os.makedirs(pasta_mes, exist_ok=True)
                os.startfile(pasta_mes)
            except (AttributeError, OSError) as e:
                st.error(f"Não foi possível abrir a pasta: {e}")

        if b4.button("Encerrar rodada", key="encerrar", use_container_width=True):
            try:
                salvar_rodada(rodada, encerrada=True)
            except OSError:
                pass
            st.session_state.pop("rodada", None)
            st.session_state.pop("conferencia", None)
            st.rerun()

        a1, a2 = st.columns([2, 1])
        auto = a1.checkbox(
            "Conferir automaticamente",
            key="auto_conf",
            help="Com o app aberto, consulta o Outlook sozinho no intervalo escolhido e salva as respostas.",
        )
        minutos = a2.selectbox(
            "Intervalo", [5, 10, 15, 30], key="auto_min", disabled=not auto,
            format_func=lambda m: f"a cada {m} min", label_visibility="collapsed",
        )
        if auto:
            ultima = (st.session_state.get("conferencia") or {}).get("hora") or rodada["inicio"]
            st.caption(f"Próxima conferência automática por volta de {ultima + timedelta(minutes=minutos):%H:%M}.")

            def _auto():
                c = st.session_state.get("conferencia")
                desde = c["hora"] if c else rodada["inicio"]
                if datetime.now() - desde >= timedelta(minutes=minutos):
                    with st.spinner("Conferindo automaticamente no Outlook..."):
                        ok = fazer_conferencia(rodada, pasta_mes)
                    if ok:
                        st.rerun()

            st.fragment(run_every=timedelta(minutes=1))(_auto)()

        r1, r2, r3 = st.columns([2, 1, 1])
        reenviar = r1.checkbox(
            "Reenviar automaticamente quem não respondeu",
            key="auto_reenvio",
            help="Com o app aberto, a cada intervalo confere o Outlook e reenvia o e-mail a quem já recebeu "
            "mas ainda não mandou o extrato. Os e-mails são enviados direto, sem abrir para revisão.",
        )
        qtd = r2.number_input(
            "Quantidade", min_value=1, max_value=999, value=1, step=1, key="reenvio_qtd",
            disabled=not reenviar, label_visibility="collapsed",
        )
        unidade = r3.selectbox(
            "Unidade", list(UNIDADES_REENVIO), index=1, key="reenvio_unidade",
            disabled=not reenviar, label_visibility="collapsed",
        )
        if reenviar:
            intervalo_reenvio = timedelta(hours=qtd * UNIDADES_REENVIO[unidade])
            st.caption(f"Reenvia a cada {qtd} {unidade} quem ainda não enviou o extrato.")

            def _reenvio():
                feitos = reenviar_sem_resposta(rodada, pasta_mes, intervalo_reenvio)
                if feitos:
                    st.toast(f"{feitos} e-mail(s) reenviado(s) automaticamente.")
                if feitos is not None:
                    st.rerun()

            def _checar_reenvio():
                # cada e-mail tem o seu prazo; confere com frequência suficiente para respeitá-los
                ultima = st.session_state.get("reenvio_ultima_checagem", rodada["inicio"])
                if datetime.now() - ultima >= min(intervalo_reenvio, timedelta(minutes=30)):
                    st.session_state["reenvio_ultima_checagem"] = datetime.now()
                    with st.spinner("Reenviando e-mails sem resposta..."):
                        _reenvio()

            st.fragment(run_every=timedelta(minutes=1))(_checar_reenvio)()

        # vigia da pasta: atualiza a tela quando chega ou sai arquivo (também os colocados à mão)
        def _vigiar():
            k = f"vigia_{pasta_mes}"
            visto = st.session_state.setdefault(k, vigia["n"])
            if vigia["n"] != visto:
                st.session_state[k] = vigia["n"]
                st.rerun()

        st.fragment(run_every="3s")(_vigiar)()

        if st.session_state.get("rodadas_outras"):
            st.caption(
                f"Há outras {st.session_state['rodadas_outras']} rodada(s) antiga(s) abertas em Extratos/_rodadas. "
                "Esta é a mais recente; encerre as que não usa mais."
            )

        if conf and conf.get("falhas"):
            st.warning(
                f"{len(conf['falhas'])} anexo(s) não puderam ser salvos e não foram contados: "
                + "; ".join(conf["falhas"][:5])
            )

        if conf:
            classes = {"Enviado": "ok", "Erro de entrega": "sem", "Na caixa de saída": "pend", "Aguardando envio": "pend"}
            total = len(rodada["itens"])
            enviados = sum(1 for v in conf["status"].values() if v == "Enviado")
            def _falta(k, it_):
                return sum(len(c) for c in it_["fundos"].values()) - (conf["recebidos"].get(k) or 0)

            n_completos = sum(1 for k, it_ in enumerate(rodada["itens"]) if _falta(k, it_) <= 0)
            n_erros = sum(1 for v in conf["status"].values() if v == "Erro de entrega")
            n_pend = total - n_completos
            c1, c2, c3, c4 = st.columns(4)
            c1.markdown(kpi("E-mails enviados", f"{enviados} de {total}", tipo="ok" if enviados == total else ""), unsafe_allow_html=True)
            c2.markdown(kpi("Com extrato recebido", len(conf["recebidos"])), unsafe_allow_html=True)
            c3.markdown(kpi("Completos", n_completos, tipo="ok"), unsafe_allow_html=True)
            c4.markdown(kpi("Com pendência", n_pend, alerta=True), unsafe_allow_html=True)
            st.caption(
                f"Última conferência às {conf['hora']:%H:%M}."
                + (f" {n_erros} com erro de entrega." if n_erros else "")
            )
            filtro = st.radio(
                "Mostrar", ["Todos", "Pendentes", "Completos"], horizontal=True,
                key="filtro_conf", label_visibility="collapsed",
            )
            larguras = [3, 1, 1.7, 1.5, 1.5, 1.5]
            cab = st.columns(larguras)
            for c, t in zip(cab, ["Fundo", "Cotistas", "Envio", "Extrato", "Alerta", "Ação"]):
                c.markdown(f'<div class="colhead">{t}</div>', unsafe_allow_html=True)
            st.markdown('<div class="sep"></div>', unsafe_allow_html=True)
            with st.container(height=540, border=False):
                for i, it in enumerate(rodada["itens"]):
                    completo = _falta(i, it) <= 0
                    if (filtro == "Pendentes" and completo) or (filtro == "Completos" and not completo):
                        continue
                    stt = conf["status"].get(i, "Aguardando envio")
                    rec = conf["recebidos"].get(i)
                    extrato = (
                        f'<span class="badge ok">Recebido ({rec})</span>' if rec else '<span class="badge pend">Aguardando</span>'
                    )
                    fundos = "; ".join(f for f in it["fundos"] if f) or "-"
                    n_cot = sum(len(c) for c in it["fundos"].values())
                    falta = n_cot - (rec or 0)
                    if falta <= 0:
                        alerta = '<span class="badge ok">Completo</span>'
                    elif rec or stt == "Enviado":
                        alerta = f'<span class="badge sem">Faltam {falta}</span>'
                    else:
                        alerta = '<span class="ativo">-</span>'

                    cols = st.columns(larguras)
                    marca = '<span class="row-warn"></span>' if (stt == "Erro de entrega" or "Faltam" in alerta) else ""
                    cols[0].markdown(f'{marca}<div class="fundo">{html.escape(fundos)}</div>', unsafe_allow_html=True)
                    cols[1].markdown(str(n_cot))
                    cols[2].markdown(f'<span class="badge {classes[stt]}">{stt}</span>', unsafe_allow_html=True)
                    cols[3].markdown(extrato, unsafe_allow_html=True)
                    cols[4].markdown(alerta, unsafe_allow_html=True)
                    # botão individual: só para e-mails não enviados (ou com erro) e sem extrato recebido
                    if stt in ("Aguardando envio", "Erro de entrega") and not rec:
                        if cols[5].button("Reenviar", key=f"reenviar_{i}", use_container_width=True):
                            if abrir_emails_outlook([it], rodada["data"]):
                                st.toast("E-mail aberto no Outlook.")
                    else:
                        cols[5].markdown('<span class="ativo">-</span>', unsafe_allow_html=True)
                    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)
            st.caption(
                f"Arquivos salvos em {pasta_mes}, uma pasta por fundo. Em e-mails com mais de um fundo, o fundo é "
                "identificado pelo nome do arquivo, pelo conteúdo do PDF, pelo assunto ou pelo corpo da resposta; "
                "se nada identificar, o anexo vai para uma pasta com o nome do remetente. O extrato é reconhecido em "
                "respostas com anexo na Caixa de Entrada padrão, pelo remetente ou pelo assunto. Na caixa de saída "
                "não entra em 'Reabrir pendentes', para evitar e-mail duplicado."
            )
