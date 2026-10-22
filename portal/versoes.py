"""Versão de cada ferramenta do portal. Cada app evolui sozinho: ao alterar um app, atualize só a linha dele.

Formato: MAIOR.MENOR.CORRECAO (ex.: 1.2.0) e a data da alteração (AAAA-MM-DD).
"""

VERSOES = {
    "portal":        {"versao": "1.0.0", "data": "2026-10-22"},  # home.py, tema.py, iniciar.py
    "extratos":      {"versao": "1.0.0", "data": "2026-10-22"},
    "boletas":       {"versao": "1.0.0", "data": "2026-10-22"},
    "rentabilidade": {"versao": "1.0.0", "data": "2026-10-22"},
    "bradesco":      {"versao": "1.0.0", "data": "2026-10-22"},
}


def rotulo(chave: str) -> str:
    """Texto curto para exibir na tela, ex.: 'v1.0.0 · 22/10/2026'."""
    v = VERSOES.get(chave)
    if not v:
        return ""
    ano, mes, dia = v["data"].split("-")
    return f"v{v['versao']} · {dia}/{mes}/{ano}"
