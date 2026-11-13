@echo off
chcp 65001 >nul
cd /d "%~dp0"
python plano_alimentar\gerar_plano.py %*
if errorlevel 1 pause
