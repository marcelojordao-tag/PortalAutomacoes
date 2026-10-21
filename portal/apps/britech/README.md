# Rentabilidade e Cotas (Britech)

Duas abas sobre as mesmas carteiras:

- **Conferência de Rentabilidade:** o PL final bate com a rentabilidade da cota? Quando não bate, mostra de onde vem a diferença (aplicações e resgates de cotas).
- **Monitor de Cotas:** o retorno diário da cota está Acima, Abaixo ou Dentro do normal da carteira (banda de média ± desvios, em 4 janelas). Gera gráficos e Excel.

## Como usar

1. Copie `.env.example` para `.env` e preencha `BRITECH_USER` e `BRITECH_PASS` (o `.env` não vai para o git).
2. No portal, abra **Rentabilidade e Cotas**, escolha os IDs (lista em `ids_cliente.xlsx`) e o período, e clique em **Buscar na API**.

## Arquivos

- `app.py`: tela (Streamlit).
- `teste.py`: busca na API Britech (pré-produção) e grava `Relatorio_Rentabilidade.xlsx`, `extratos/` e `Historico_Cotas.csv`. Também roda sozinho com `python teste.py`.
- `converter.py`: conferência de rentabilidade.
- `monitor.py`: lógica do monitor de cotas.
- `tema_britech.py`: estilo próprio desta ferramenta (o cabeçalho vem do tema do portal).
- `ids_cliente.xlsx`: IDs das carteiras.

Os arquivos gerados na busca ficam fora do git.
