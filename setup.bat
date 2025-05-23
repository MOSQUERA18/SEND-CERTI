@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

:: Obtener la ruta actual y verificar si está en "C:\send-message"
set "ruta_base=C:\send-message"
echo %CD% | findstr /I /C:"%ruta_base%" >nul
if %errorlevel% neq 0 (
    echo ERROR: La carpeta debe estar alojada en "%ruta_base%" para ejecutar correctamente los iconos.
    pause
    exit /b
)

:: Crear entorno virtual (opcional)
if not exist venv (
    python -m venv venv
)

call venv\Scripts\activate

:: Instalar dependencias
pip install --upgrade pip
pip install pandas selenium webdriver-manager tk openpyxl python-dotenv pyinstaller

:: Crear el ejecutable con pyinstaller
pyinstaller --onefile --noconsole --icon="backend/iconos/envio_wasap.ico" main.py

:: Mensaje final
echo.
echo ==============================================
echo Proceso finalizado. El ejecutable está en "dist\main.exe"
echo ==============================================
pause
