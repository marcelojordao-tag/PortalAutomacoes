@echo off
cd /d %~dp0

where python >nul 2>&1
if errorlevel 1 (
    echo Python nao encontrado. Instale o Python 3 e marque "Add to PATH".
    pause
    exit /b 1
)

python iniciar.py
pause
