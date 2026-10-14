@echo off
REM ============================================================
REM  Genera dist\InventarioTI.exe (un solo archivo, sin consola).
REM  Ejecutar EN WINDOWS con Python 3.10 o superior instalado.
REM  Doble clic o desde una terminal: build_exe.bat
REM ============================================================
setlocal
cd /d "%~dp0"

if not exist venv\Scripts\activate.bat (
    echo Creando entorno virtual...
    where py >nul 2>nul
    if errorlevel 1 ( python -m venv venv ) else ( py -3 -m venv venv )
    if errorlevel 1 goto :error
)

call venv\Scripts\activate.bat
if errorlevel 1 goto :error

echo Instalando dependencias...
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 goto :error

echo Generando el .exe...
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name InventarioTI ^
  --add-data "assets;assets" ^
  --collect-all customtkinter ^
  --hidden-import win32timezone ^
  --hidden-import PIL._tkinter_finder ^
  app.py
if errorlevel 1 goto :error

echo.
echo ============================================================
echo  Listo: dist\InventarioTI.exe
echo  Copialo a una carpeta donde se pueda escribir: junto al .exe
echo  se crean data\ (base de datos), actas\ y reportes\.
echo ============================================================
pause
exit /b 0

:error
echo.
echo *** Algo fallo. Revisa el mensaje de arriba. ***
pause
exit /b 1
