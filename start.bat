@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

title GESA - Gestor de Evaluaciones de Suficiencia Académica

set "PYTHON_URL=https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe"
set "PYTHON_INSTALLER=%~dp0python_installer.exe"
set "REQUIREMENTS=%~dp0requirements.txt"
set "APP=%~dp0desktop_app.py"

rem ── 1. RUTA RÁPIDA: Si Python y las dependencias ya están listas, iniciar de inmediato (< 1s) ──
python -c "import PyQt6, docx, win32com" >nul 2>&1
if not errorlevel 1 (
    python "%APP%"
    exit /b 0
)

py -3 -c "import PyQt6, docx, win32com" >nul 2>&1
if not errorlevel 1 (
    py -3 "%APP%"
    exit /b 0
)

rem ── 2. CONFIGURACIÓN INICIAL / INSTALACIÓN (solo si falta algún componente) ──
cls
echo =======================================================
echo    GESA - Gestor de Evaluaciones de Suficiencia
echo =======================================================
echo.

echo [1/3] Verificando entorno Python...

set "PY_CMD="
python --version >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=python"
) else (
    py -3 --version >nul 2>&1
    if not errorlevel 1 (
        set "PY_CMD=py -3"
    ) else (
        if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
            set "PATH=%LocalAppData%\Programs\Python\Python312;%LocalAppData%\Programs\Python\Python312\Scripts;%PATH%"
            set "PY_CMD=python"
        )
    )
)

if defined PY_CMD (
    echo [OK] Python detectado correctamente.
    goto install_deps
)

echo.
echo [!] Python 3 no se encuentra instalado en este equipo.
echo [>] Descargando e instalando Python 3.12 automaticamente (proceso de unica vez)...
echo.

powershell -Command "Write-Host 'Descargando instalador de Python...'; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%PYTHON_URL%', '%PYTHON_INSTALLER%')"

if not exist "%PYTHON_INSTALLER%" (
    echo.
    echo [ERROR] No se pudo descargar el instalador de Python.
    echo Por favor verifica tu conexion a Internet e intenta nuevamente.
    pause
    exit /b 1
)

echo [>] Instalando Python 3.12 (espera unos segundos)...
start /wait "" "%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1 SimpleInstall=1
del "%PYTHON_INSTALLER%" 2>nul

set "PATH=%LocalAppData%\Programs\Python\Python312;%LocalAppData%\Programs\Python\Python312\Scripts;%PATH%"

python --version >nul 2>&1
if errorlevel 1 (
    py -3 --version >nul 2>&1
    if not errorlevel 1 (
        set "PY_CMD=py -3"
    ) else (
        echo [ERROR] No se pudo activar Python automaticamente. Reinicia la aplicacion.
        pause
        exit /b 1
    )
) else (
    set "PY_CMD=python"
)

echo [OK] Python instalado y configurado correctamente.

:install_deps
echo.
echo [2/3] Instalando librerias necesarias (PyQt6, python-docx, pywin32)...
!PY_CMD! -m pip install -r "%REQUIREMENTS%" --quiet --no-warn-script-location
!PY_CMD! -c "import win32com.client" 2>nul
if errorlevel 1 (
    !PY_CMD! -m pip install pywin32 --quiet --no-warn-script-location
)

echo [OK] Librerias instaladas.

echo.
echo [3/3] Registrando accesos directos...
powershell -ExecutionPolicy Bypass -File "%~dp0create_shortcut.ps1" >nul 2>&1

echo.
echo [>] Iniciando GESA...
echo.

!PY_CMD! "%APP%"

endlocal
