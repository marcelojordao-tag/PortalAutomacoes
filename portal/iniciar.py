"""Inicia o Portal de Automações.

Valida se as bibliotecas estão instaladas; se não, instala para o usuário atual
(pip --user), sem criar venv. Depois sobe o Streamlit.
"""
import glob
import importlib.util
import os
import subprocess
import sys

PASTA = os.path.dirname(os.path.abspath(__file__))
REQUIREMENTS = os.path.join(PASTA, "requirements.txt")

# módulo importado -> pacote do requirements.txt
MODULOS = [
    "streamlit", "pandas", "numpy", "plotly", "altair", "openpyxl",
    "xlrd", "pypdf", "watchdog", "requests", "win32api", "playwright",
]


def faltando():
    return [m for m in MODULOS if importlib.util.find_spec(m) is None]


def executar(*args):
    return subprocess.call([sys.executable, *args], cwd=PASTA)


def chromium_instalado():
    base = os.environ.get("PLAYWRIGHT_BROWSERS_PATH") or os.path.join(
        os.environ.get("LOCALAPPDATA", ""), "ms-playwright"
    )
    return bool(glob.glob(os.path.join(base, "chromium-*")))


def main():
    ausentes = faltando()
    if ausentes:
        print(f"Bibliotecas ausentes ({', '.join(ausentes)}). Instalando...")
        if executar("-m", "pip", "install", "--user", "-r", REQUIREMENTS) != 0:
            return erro()
        if faltando():
            print("Instalação concluída, mas ainda faltam módulos:", faltando())
            return erro()

    if not chromium_instalado():
        print("Instalando Chromium do Playwright...")
        if executar("-m", "playwright", "install", "chromium") != 0:
            return erro()

    return executar("-m", "streamlit", "run", "home.py", "--browser.gatherUsageStats", "false")


def erro():
    print("\nFalha na instalação. Verifique a internet e rode novamente.")
    return 1


if __name__ == "__main__":
    codigo = main()
    if codigo:
        input("\nPressione Enter para sair...")
    sys.exit(codigo)
