@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

title GESA - Gestor de Evaluaciones de Suficiencia Académica

set "PYTHON_URL=https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe"
set "PYTHON_INSTALLER=%~dp0python_installer.exe"
set "REQUIREMENTS=%~dp0requirements.txt"
set "APP=%~dp0desktop_app.py"

cls
echo.
echo ========================================
echo   GESTOR DE EVALUACIONES DE
echo   SUFICIENCIA ACADÉMICA
echo ========================================
echo v1.0 ^| Generación automática de exámenes
echo.

echo [1/6] Comprobando actualizaciones desde GitHub...
git --version >nul 2>&1
if not errorlevel 1 (
    if exist "%~dp0.git" (
        git pull origin main --quiet 2>nul
        echo [OK] Código sincronizado con GitHub
    ) else (
        echo [OK] Inicio directo
    )
) else (
    echo [OK] Inicio directo
)

echo.
echo [2/6] Verificando Python...
python --version >nul 2>&1
if not errorlevel 1 (
    echo [OK] Python detectado
    goto :step2
)

echo [!] Python 3.10+ no encontrado
echo.
echo Descargando Python (64-bit)...
powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%PYTHON_URL%', '%PYTHON_INSTALLER%')}"
if not exist "%PYTHON_INSTALLER%" (
    echo [FAIL] No se pudo descargar Python.
    pause
    exit /b 1
)
echo Instalando Python (espera unos segundos)...
start /wait "" "%PYTHON_INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_launcher=0
del "%PYTHON_INSTALLER%" 2>nul
echo [OK] Python instalado.
echo.
echo IMPORTANTE: Cierra y abre el programa de nuevo
echo para que los cambios de PATH tengan efecto.
pause
exit /b 0

:step2
echo.
echo [3/6] Actualizando pip...
python -m pip install --upgrade pip -q
echo [OK] pip actualizado

echo.
echo [4/6] Instalando dependencias...
python -c "import win32com.client" 2>nul
if errorlevel 1 (
    python -m pip install pywin32 -q
    python Scripts\pywin32_postinstall.py -install 2>nul
    echo pywin32 instalado
)
python -m pip install -r "%REQUIREMENTS%" -q
echo [OK] Paquetes listos

echo.
echo [5/6] Verificando Microsoft Word...
reg query "HKLM\SOFTWARE\Microsoft\Office" 2>nul >nul
if not errorlevel 1 (
    echo [OK] Microsoft Word detectado
) else (
    reg query "HKCU\SOFTWARE\Microsoft\Office" 2>nul >nul
    if not errorlevel 1 (
        echo [OK] Microsoft Word detectado
    ) else (
        echo [!] No se detectó Microsoft Word
        echo La generación de exámenes requiere Word.
        choice /c sn /n /m "Continuar? (S/N): "
        if errorlevel 2 exit /b
    )
)

echo.
echo [6/6] Arrancando aplicación...
echo.
echo ========================================
echo        TODO LISTO
echo ========================================
echo.

python "%APP%"
if errorlevel 1 (
    echo.
    echo [FAIL] La aplicación se cerró con error
    pause
)
endlocal
