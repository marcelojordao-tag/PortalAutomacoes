# Portal de Automações | TAG Investimentos

Portal web interno que reúne, em um só lugar, as ferramentas de automação usadas pela área de operações. Cada ferramenta é uma página do portal, com visual padronizado (identidade TAG) e acesso por um menu lateral.

## Objetivo

Reduzir trabalho manual e erros em rotinas operacionais, centralizando as automações em uma interface única, simples de usar e de ampliar. Hoje o portal cobre duas rotinas: a solicitação e conferência de extratos de cotistas, e a conciliação de boletas com controle de erros.

## Ferramentas

### Solicitação de extratos (`portal/apps/extratos`)
Automatiza o ciclo de pedido e recebimento de extratos de cotistas.
- Lê a planilha `Template_Fundos_Ativos.xlsx` (fundos, cotistas e destinatários) e monta os itens a enviar.
- Gera os e-mails (assunto e corpo em HTML) e os abre no **Outlook**, preservando a assinatura. A data de referência considera feriados nacionais e o dia útil anterior.
- Confere no Outlook o que foi enviado e o que foi respondido.
- Identifica os anexos recebidos (por nome do arquivo, remetente e conteúdo do PDF) e os organiza nas pastas do mês.
- Salva o estado de cada rodada para retomá-la depois.

### Mesa de Boletas (`portal/apps/boletador`)
Reúne três abas:
- **Conferência de Boletas:** importa a base da TAG diretamente do boletador (`importador.py`, via Playwright) ou aceita a planilha enviada manualmente, e a concilia com a base da XP. Mostra o que está conciliado, divergente ou ausente, e permite baixar o resultado em Excel.
- **Controle de Liquidação:** aba reservada, ainda em construção ("Em breve").
- **Controle de Erros:** quadro kanban de problemas e gargalos, com prioridade, categoria, causa raiz, conexões entre problemas, histórico de ocorrências, dashboard e grafo. Os dados ficam em `controle_erros.json`, com backup diário em `backups_erros/`.

## Estrutura

```
portal/
├── home.py             # página inicial e navegação
├── tema.py             # tema visual compartilhado (topo, cores, selos)
├── logo_tag.png
├── iniciar.bat         # cria o ambiente e inicia o portal
├── requirements.txt
└── apps/
    ├── extratos/       # Solicitação de extratos
    └── boletador/      # Mesa de Boletas (+ Padrao_Topo_Header.md, padrão do cabeçalho)
```

## Como executar

Pré-requisitos: **Windows**, **Python 3.12** e **Outlook** instalado e aberto (para a ferramenta de extratos). Na Mesa de Boletas, a importação automática precisa de acesso ao boletador da TAG.

Dê dois cliques em `portal/iniciar.bat`. Na primeira vez, ele cria o ambiente virtual, instala as dependências e o Chromium do Playwright. Depois, inicia o portal no navegador.

Ou, manualmente, dentro de `portal/`:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
streamlit run home.py
```

## Bibliotecas utilizadas

Dependências em `portal/requirements.txt`:

| Biblioteca | Uso |
|---|---|
| `streamlit` | Interface web do portal e navegação entre as ferramentas |
| `pandas` | Leitura e tratamento das planilhas e conciliações |
| `numpy` | Cálculos numéricos de apoio |
| `plotly` | Gráficos do dashboard de erros |
| `openpyxl` | Leitura e escrita de arquivos `.xlsx` |
| `xlrd` | Leitura de arquivos `.xls` |
| `pypdf` | Leitura do texto de PDFs para identificar os extratos recebidos |
| `watchdog` | Monitoramento das pastas de extratos |
| `pywin32` | Integração com o Outlook (somente Windows) |
| `playwright` | Navegador automatizado para importar a base do boletador |

## Como adicionar uma nova ferramenta

1. Crie uma pasta em `portal/apps/<nome>/` com o `app.py` da ferramenta.
2. Em `portal/home.py`, crie a `st.Page` e acrescente uma entrada na lista `FERRAMENTAS`. O card na home e o item no menu são gerados a partir dela.
3. Use `tema.py` para manter o visual igual ao das demais ferramentas.

## Segurança

Não coloque senhas, chaves de API ou credenciais no código. Use variáveis de ambiente (`.env`, já ignorado pelo `.gitignore`). Esta base contém dados operacionais (planilha de fundos, controle de erros). Trate o repositório como interno.

## Contato

- **Responsável:** Marcelo Vinicius Jordão Almeida, marcelo.jordao@taginvest.com.br
