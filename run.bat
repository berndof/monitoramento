@echo off
:: Navega para o diretório onde o script .bat está
cd /d %~dp0

:: Executa o script principal usando 'uv run'
uv run main.py


