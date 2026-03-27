@echo off
title CIMA Planning v3.0
cd /d "%~dp0"
echo.
echo  ████████████████████████████████████████████████████
echo     CIMA Planning v3.0  —  Iniciando servidor...
echo  ████████████████████████████████████████████████████
echo.

REM Crear carpetas si no existen
IF NOT EXIST "assets\"   mkdir "assets"
IF NOT EXIST "static\"   mkdir "static"
IF NOT EXIST "frontend\" mkdir "frontend"
IF NOT EXIST "backend\reports\" mkdir "backend\reports"

REM Verificar Python
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo  [ERROR] Python no encontrado. Instala Python 3.11+ y agregalo al PATH.
    pause
    exit /b 1
)

echo  Instalando / verificando dependencias...
cd backend
pip install -r requirements.txt -q --disable-pip-version-check

echo.
echo  Servidor iniciando en: http://localhost:8000
echo  Presiona Ctrl+C para detener.
echo.

REM Abrir navegador despues de 2 segundos
start /b cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:8000"

REM Iniciar el servidor (bloqueante)
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload

pause
