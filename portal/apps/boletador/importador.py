"""Importação da base da TAG (boletador) via Playwright.

Roda como processo separado (`python importador.py`), chamado pelo app. Assim o Playwright
síncrono não convive com o event loop do Streamlit, e o processo abre sem carregar
streamlit/pandas/plotly.

Códigos de saída: 0 ok | 2 tempo esgotado no login | 3 erro ou navegador fechado.
O caminho do arquivo baixado é impresso na saída com o prefixo MARCADOR_ARQUIVO.
"""
import sys
from pathlib import Path

URL_BOLETADOR = "https://boletador-hom.taginvest.com.br/"
PERFIL_NAVEGADOR = Path.home() / ".boletador_perfil"  # guarda o login entre execuções
TEMPO_MAX_LOGIN = 300  # segundos esperando o usuário logar
PASTA_DESTINO = Path(__file__).resolve().parent / "downloads_boletador"
MARCADOR_ARQUIVO = "ARQUIVO::"  # o subprocesso imprime o caminho do arquivo com esse prefixo


def importar_base() -> None:
    """Abre o boletador, espera o login, navega até Ordens e baixa a base."""
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        opcoes = dict(user_data_dir=str(PERFIL_NAVEGADOR), headless=False, no_viewport=True)
        try:
            ctx = p.chromium.launch_persistent_context(channel="chrome", **opcoes)
        except Exception:
            ctx = p.chromium.launch_persistent_context(**opcoes)  # Chromium do Playwright

        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto(URL_BOLETADOR)

            # O login termina quando o ícone de menu aparece na tela
            icone_menu = page.locator('i.material-icons[title="Menu"]')
            try:
                icone_menu.wait_for(state="visible", timeout=TEMPO_MAX_LOGIN * 1000)
            except PlaywrightTimeoutError:
                sys.exit(2)

            # Navega para Cadastros > Ordens
            ordens = page.locator('a:text-is("Ordens"):visible').first

            if not ordens.is_visible():  # só abre "Cadastros" se o submenu ainda estiver fechado
                page.locator("#itemCadastros").click()
            ordens.click()

            # Dá 3 segundos para o Angular carregar completamente a tabela e os componentes visuais
            page.wait_for_timeout(3000)

            # 1. Pega o PRIMEIRO botão visível
            download = page.locator('button[title="Exportar"]').first
            download.wait_for(state="visible", timeout=15000)

            print("Botão encontrado! Tentando clicar via JavaScript...")

            # 2. Prepara o Playwright para o download
            with page.expect_download(timeout=120000) as download_info:
                # 3. O clique via JavaScript (ignora qualquer bloqueio visual da tela)
                download.evaluate("botao => botao.click()")

            print("Clique funcionou! Aguardando o download terminar...")
            arquivo = download_info.value

            # 4. Espera a tela de loading sumir
            spinner = page.locator("#cover-spin")
            try:
                spinner.wait_for(state="hidden", timeout=120000)
            except Exception:
                pass  # se der erro aqui, o loading já sumiu

            # 5. Salva o arquivo na pasta de destino
            PASTA_DESTINO.mkdir(parents=True, exist_ok=True)
            caminho_completo = PASTA_DESTINO / arquivo.suggested_filename
            arquivo.save_as(caminho_completo)

            # 6. Devolve o caminho para o app principal (lido pelo stdout)
            print(f"{MARCADOR_ARQUIVO}{caminho_completo}", flush=True)

        except SystemExit:
            raise
        except Exception as e:
            print(f"Erro durante a execução do Playwright: {e}", file=sys.stderr, flush=True)
            sys.exit(3)
        finally:
            try:
                ctx.close()  # fecha o navegador em qualquer cenário
            except Exception:
                pass


if __name__ == "__main__":
    importar_base()
