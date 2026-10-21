"""Busca os dados das carteiras na API Britech e grava o Relatorio_Rentabilidade.xlsx + os extratos de conta corrente.

Uso na linha de comando (le os IDs de ids_cliente.xlsx):
    python teste.py

Uso pelo app (app.py): o app escolhe os IDs e o periodo e chama buscar_dados().
"""
import os
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

import pandas as pd
import requests

# BASE_URL = "https://tag.britech.com.br/WS"  # PROD (desativada para não sobrecarregar o pessoal de produção)
BASE_URL = "https://tag-pre-prod.britech.com.br/ws"  # PRÉ-PROD

def _ler_env(caminho: Path) -> None:
    """Carrega o .env da pasta do app (BRITECH_USER / BRITECH_PASS) sem sobrescrever variáveis já definidas."""
    try:
        linhas = caminho.read_text(encoding="utf-8").splitlines()
    except OSError:
        return
    for linha in linhas:
        linha = linha.strip()
        if linha and not linha.startswith("#") and "=" in linha:
            chave, valor = linha.split("=", 1)
            os.environ.setdefault(chave.strip(), valor.strip().strip('"').strip("'"))


_ler_env(Path(__file__).resolve().parent / ".env")
USUARIO = os.environ.get("BRITECH_USER", "")
SENHA = os.environ.get("BRITECH_PASS", "")

PASTA = Path(__file__).resolve().parent
ARQUIVO_IDS = PASTA / "ids_cliente.xlsx"
ARQUIVO_SAIDA = PASTA / "Relatorio_Rentabilidade.xlsx"
PASTA_EXTRATOS = PASTA / "extratos"
DATA_INICIO_PADRAO = "2026-07-31"
DATA_FIM_PADRAO = "2026-08-31"

# Historico de cotas para o Monitor de cotas (precisa de bem mais dias que a conferencia: as janelas de 252 e 126
# dias so comecam a classificar depois de metade da janela)
ARQUIVO_HISTORICO = PASTA / "Historico_Cotas.csv"
HISTORICO_DIAS_PADRAO = 730  # 2 anos antes da data final
COLUNAS_HISTORICO = ["Data", "IdCarteira", "CotaBruta", "PLFechamento", "QuantidadeFechamento"]

# Tempo limite das chamadas (segundos): (conectar, esperar resposta). Sem isso, uma API que trava
# deixa o programa esperando para sempre (no app, o botao "Buscar na API" ficaria preso).
TIMEOUT = (15, 120)
TIMEOUT_EXTRATO = (15, 180)  # o extrato e um arquivo e demora mais


class ErroAPI(RuntimeError):
    """A API nao respondeu a tempo ou nao foi possivel conecta-la. A mensagem ja e amigavel para mostrar ao usuario."""


class AtlasPasAPI:
    def __init__(self, base_url: str, usuario: str, senha: str):
        self.base_url = base_url.rstrip('/')
        self.session = requests.Session()
        self.session.auth = (usuario, senha)
        self.session.headers.update({
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        })

    def close(self) -> None:
        """Fecha as conexoes abertas da sessao."""
        self.session.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    def _executar_get(self, endpoint: str, params: Dict[str, Any]) -> Optional[List[Dict[str, Any]]]:
        url = f"{self.base_url}/{endpoint}"
        try:
            print(f"Consultando: {url} {params}")
            response = self.session.get(url, params=params, timeout=TIMEOUT)

            if response.status_code == 401:
                print("❌ ERRO: Autenticação falhou.")
                return None

            response.raise_for_status()
            return response.json()

        except (requests.exceptions.ConnectTimeout, requests.exceptions.ConnectionError) as e:
            raise ErroAPI(f"Não consegui conectar na API ({endpoint}). Confira a rede e o acesso. Detalhe: {e}") from None
        except requests.exceptions.Timeout:
            raise ErroAPI(f"A API não respondeu em {TIMEOUT[1]} segundos ({endpoint}). Tente de novo daqui a pouco "
                          "ou use um período menor.") from None
        except Exception as e:
            print(f"❌ Erro na requisição: {e}")
        return None

    def baixar_extrato_conta_corrente(self, id_cliente, data_inicio, data_fim, caminho, periodo=1):
        """Baixa o extrato de conta corrente em Excel (a API devolve o arquivo, nao JSON). periodo=1 e 'Mes'."""
        url = f"{self.base_url}/api/Relatorio/RelatorioExtratoContaCorrente"
        params = {"IdCliente": id_cliente, "DataInicio": data_inicio, "DataFim": data_fim,
                  "Periodo": periodo, "TipoArquivo": "Excel"}
        try:
            print(f"Consultando: {url} {params}")
            response = self.session.get(url, params=params, timeout=TIMEOUT_EXTRATO)
            response.raise_for_status()
            with open(caminho, "wb") as f:
                f.write(response.content)
            return True
        except Exception as e:
            print(f"❌ Erro ao baixar extrato: {e}")
        return False

    def buscar_rentabilidade_ativo_dia(self, id_cliente, data_inicio, data_fim):
        return self._executar_get("api/Rentabilidade/BuscaRentabAtivo_Dia",{"idCliente": id_cliente, "dataInicio": data_inicio, "dataFim": data_fim})

    def buscar_rentabilidade_classe(self, id_cliente, data_inicial, data_final):
        return self._executar_get("api/Rentabilidade/BuscaRentabClasse_PeriodoNaoExplodida",{"IdCliente": id_cliente, "DataInicial": data_inicial, "DataFinal": data_final})

    def buscar_rentabilidade_subclasse(self, id_cliente, data_inicial, data_final):
        return self._executar_get("api/Rentabilidade/BuscaRentabSubClasse_PeriodoNaoExplodida",{"IdCliente": id_cliente, "DataInicial": data_inicial, "DataFinal": data_final})

    def buscar_rentabilidade_estrategia(self, id_cliente_proprietario, data_inicio, data_fim):
        return self._executar_get("api/Rentabilidade/BuscaRentabEstrategia_PeriodoNaoExplodida",{"idClienteProprietario": id_cliente_proprietario, "dataInicio": data_inicio, "dataFim": data_fim})

    def buscar_historico_cota_dia(self, id_cliente, data_inicio, data_fim):
        """Cota, PL e quantidade de cotas por dia (mostra se houve entrada/saida de cotas)."""
        return self._executar_get("api/Fundo/BuscaHistoricoCotaDia",{"idCarteira": id_cliente, "dataInicio": data_inicio, "dataFim": data_fim})

    def buscar_operacoes_cotista(self, id_cliente, data_inicio, data_fim):
        """Aplicacoes e resgates de cotas da carteira no periodo (sem precisar do ID do cotista)."""
        return self._executar_get("api/Fundo/OperacaoCotistaAnalitico_PorCarteiraPeriodo",{"IdsCarteiras": id_cliente, "DataInicio": data_inicio, "DataFim": data_fim})


def ler_ids(arquivo=ARQUIVO_IDS) -> List[str]:
    """IDs da coluna id_cliente da planilha (sem repeticao). Lista vazia se a planilha nao existir."""
    arquivo = Path(arquivo)
    if not arquivo.exists():
        return []
    return (pd.read_excel(arquivo, dtype=str)["id_cliente"]
            .dropna().str.strip().drop_duplicates().tolist())


def montar_consultas(api: AtlasPasAPI) -> Dict[str, Callable]:
    """aba do Excel -> consulta da API. As duas ultimas explicam divergencias de PL (movimentacao de cotistas)."""
    return {
        "Ativo": api.buscar_rentabilidade_ativo_dia,
        "Classe": api.buscar_rentabilidade_classe,
        "SubClasse": api.buscar_rentabilidade_subclasse,
        "Estrategia": api.buscar_rentabilidade_estrategia,
        "HistCota": api.buscar_historico_cota_dia,
        "OpCotista": api.buscar_operacoes_cotista,
    }


def _executar_historico(clientes, data_inicio, data_fim, arquivo_saida: Path, api: AtlasPasAPI, progresso) -> dict:
    """Uma chamada por carteira (BuscaHistoricoCotaDia) e um CSV com Data, IdCarteira, CotaBruta, PL e quantidade de cotas."""
    quadros, sem_dados = [], []
    for i, cliente in enumerate(clientes):
        if progresso:
            progresso(i / len(clientes), f"Carteira {cliente}: histórico de cotas")
        dados = api.buscar_historico_cota_dia(cliente, data_inicio, data_fim)
        if dados:
            df = pd.DataFrame(dados)
            for c in COLUNAS_HISTORICO:
                if c not in df.columns:
                    df[c] = None
            quadros.append(df[COLUNAS_HISTORICO])
        else:
            print(f"⚠️ Sem histórico de cotas: carteira {cliente}")
            sem_dados.append(cliente)
    if not quadros:
        raise RuntimeError("A API não devolveu histórico de cotas para nenhuma carteira.")
    historico = pd.concat(quadros, ignore_index=True)
    historico["Data"] = pd.to_datetime(historico["Data"]).dt.strftime("%Y-%m-%d")
    arquivo_saida.parent.mkdir(parents=True, exist_ok=True)
    temporario = arquivo_saida.with_name(arquivo_saida.stem + ".tmp.csv")
    historico.to_csv(temporario, index=False)
    try:
        os.replace(temporario, arquivo_saida)
    except PermissionError:
        temporario.unlink(missing_ok=True)
        raise PermissionError(f"Não consegui gravar {arquivo_saida.name}. Feche o arquivo e tente de novo.")
    return {"arquivo": str(arquivo_saida), "linhas": len(historico),
            "carteiras": [c for c in clientes if c not in sem_dados], "sem_dados": sem_dados}


def buscar_historico_cotas(clientes, data_inicio, data_fim, arquivo_saida=ARQUIVO_HISTORICO,
                           api: Optional[AtlasPasAPI] = None,
                           progresso: Optional[Callable[[float, str], None]] = None) -> dict:
    """Busca so o historico de cotas (para o Monitor de cotas). Mesmas regras de erro e de fechamento de buscar_dados()."""
    clientes = list(dict.fromkeys(str(c).strip() for c in clientes if str(c).strip()))
    if not clientes:
        raise ValueError("Nenhuma carteira informada.")
    data_inicio, data_fim = str(data_inicio), str(data_fim)
    if data_inicio >= data_fim:
        raise ValueError("A data inicial precisa ser anterior à data final.")
    propria = api is None
    api = api or AtlasPasAPI(BASE_URL, USUARIO, SENHA)
    try:
        return _executar_historico(clientes, data_inicio, data_fim, Path(arquivo_saida), api, progresso)
    finally:
        if propria:
            api.close()


def buscar_dados(clientes, data_inicio, data_fim, arquivo_saida=ARQUIVO_SAIDA, pasta_extratos=PASTA_EXTRATOS,
                 api: Optional[AtlasPasAPI] = None,
                 progresso: Optional[Callable[[float, str], None]] = None,
                 historico_inicio=None, arquivo_historico=ARQUIVO_HISTORICO) -> dict:
    """Busca as carteiras na API e grava o Excel e os extratos.

    clientes          - lista de IDs
    data_inicio       - dia-base ('AAAA-MM-DD' ou date); data_fim - ultimo dia
    progresso         - funcao opcional (fracao de 0 a 1, texto) para mostrar andamento
    historico_inicio  - se informado, busca tambem o historico de cotas desde essa data (Monitor de cotas)
    Levanta ValueError (entrada invalida), RuntimeError/ErroAPI (API sem dados, sem resposta ou fora do ar) ou
    PermissionError (Excel aberto).
    Retorna {'arquivo', 'abas', 'clientes', 'sem_dados': {aba: [ids]}, 'extratos_falhos': [ids], 'historico': dict|None}.
    O arquivo antigo so e substituido se tudo der certo. A conexao com a API e fechada no fim, mesmo se der erro.
    Se so o historico falhar, a conferencia ja gravada e mantida e o erro vem em resultado['historico']['erro'].
    """
    clientes = list(dict.fromkeys(str(c).strip() for c in clientes if str(c).strip()))
    if not clientes:
        raise ValueError("Nenhuma carteira informada.")
    data_inicio, data_fim = str(data_inicio), str(data_fim)
    if data_inicio >= data_fim:
        raise ValueError("A data inicial precisa ser anterior à data final.")
    if historico_inicio:
        historico_inicio = str(historico_inicio)
        if historico_inicio >= data_fim:
            raise ValueError("O início do histórico do monitor precisa ser anterior à data final.")

    propria = api is None  # so fecha a sessao se foi criada aqui (quem passa a sua cuida dela)
    api = api or AtlasPasAPI(BASE_URL, USUARIO, SENHA)
    try:
        return _executar_busca(clientes, data_inicio, data_fim, Path(arquivo_saida), Path(pasta_extratos), api, progresso,
                               historico_inicio, Path(arquivo_historico))
    finally:
        if propria:
            api.close()


def _executar_busca(clientes, data_inicio, data_fim, arquivo_saida: Path, pasta_extratos: Path, api: AtlasPasAPI,
                    progresso, historico_inicio=None, arquivo_historico: Path = ARQUIVO_HISTORICO) -> dict:
    consultas = montar_consultas(api)
    total = len(clientes) * (len(consultas) + 1 + (1 if historico_inicio else 0))
    feitos = 0

    def avisar(texto: str) -> None:
        if progresso:
            progresso(min(feitos / total, 1.0), texto)

    resultados: Dict[str, List[pd.DataFrame]] = {aba: [] for aba in consultas}
    sem_dados: Dict[str, List[str]] = {aba: [] for aba in consultas}

    for cliente in clientes:
        print(f"\n=== Carteira {cliente} ===")
        for aba, consulta in consultas.items():
            avisar(f"Carteira {cliente}: {aba}")
            dados = consulta(cliente, data_inicio, data_fim)
            if dados:
                df = pd.DataFrame(dados)
                df.insert(0, "id_cliente_consulta", cliente)
                resultados[aba].append(df)
            else:
                # OpCotista vazio e normal: significa que nao houve aplicacao nem resgate de cotas
                print(f"⚠️ Sem dados: {aba} / carteira {cliente}")
                sem_dados[aba].append(cliente)
            feitos += 1

    # Extrato de conta corrente de cada carteira (usado pelo converter.py na ponte da diferenca)
    pasta_extratos.mkdir(parents=True, exist_ok=True)
    extratos_falhos = []
    for cliente in clientes:
        avisar(f"Carteira {cliente}: extrato de conta corrente")
        ok = api.baixar_extrato_conta_corrente(
            cliente, data_inicio, data_fim, pasta_extratos / f"ReportExtratoContaCorrente_{cliente}.xlsx")
        if not ok:
            extratos_falhos.append(cliente)
        feitos += 1

    abas = [aba for aba, dfs in resultados.items() if dfs]
    if "Ativo" not in abas:
        raise RuntimeError("A API não devolveu dados de ativos para nenhuma carteira. Confira os IDs, o período e o acesso.")

    # Grava num arquivo temporario e so troca no fim: se o Excel estiver aberto, o arquivo antigo fica intacto
    avisar("Gravando o Excel")
    temporario = arquivo_saida.with_name(arquivo_saida.stem + ".tmp.xlsx")
    with pd.ExcelWriter(temporario, engine="openpyxl") as escritor:
        for aba in abas:
            pd.concat(resultados[aba], ignore_index=True).to_excel(escritor, sheet_name=aba, index=False)
            print(f"✅ Aba '{aba}' gerada.")
    try:
        os.replace(temporario, arquivo_saida)
    except PermissionError:
        temporario.unlink(missing_ok=True)
        raise PermissionError(f"Não consegui gravar {arquivo_saida.name}. Feche o arquivo no Excel e tente de novo.")

    # Historico para o Monitor de cotas: depois da conferencia gravada, para nao perde-la se so esta etapa falhar
    historico = None
    if historico_inicio:
        base = feitos

        def progresso_historico(fracao: float, texto: str) -> None:
            if progresso:
                progresso(min((base + fracao * len(clientes)) / total, 1.0), texto)

        try:
            historico = _executar_historico(clientes, historico_inicio, data_fim, arquivo_historico, api, progresso_historico)
        except (RuntimeError, ValueError, PermissionError) as e:
            historico = {"erro": str(e)}

    if progresso:
        progresso(1.0, "Concluído")
    return {"arquivo": str(arquivo_saida), "abas": abas, "clientes": clientes, "sem_dados": sem_dados,
            "extratos_falhos": extratos_falhos, "historico": historico}


def main() -> None:
    clientes = ler_ids()
    print(f"{len(clientes)} carteiras lidas de {ARQUIVO_IDS}")
    historico_inicio = (date.fromisoformat(DATA_FIM_PADRAO) - timedelta(days=HISTORICO_DIAS_PADRAO)).isoformat()
    try:
        resultado = buscar_dados(clientes, DATA_INICIO_PADRAO, DATA_FIM_PADRAO, historico_inicio=historico_inicio)
    except (ValueError, RuntimeError, PermissionError) as e:  # mensagem clara em vez de um traceback longo
        print(f"\n❌ {e}")
        raise SystemExit(1)
    print(f"\n🚀 Arquivo '{resultado['arquivo']}' salvo.")
    h = resultado["historico"]
    if h and "erro" in h:
        print(f"⚠️ Histórico de cotas (monitor) não gravado: {h['erro']}")
    elif h:
        print(f"📈 Histórico de cotas: {h['linhas']} linhas em '{h['arquivo']}'.")


if __name__ == "__main__":
    main()
