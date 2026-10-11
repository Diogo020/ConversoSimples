@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    py -3 -m venv .venv
    if errorlevel 1 (
        echo Nao foi possivel criar o ambiente Python.
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :error
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :error
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name ClaraConvert conversor.py
if errorlevel 1 goto :error

echo.
echo Executavel criado em: dist\ClaraConvert.exe
pause
exit /b 0

:error
echo.
echo O processo de compilacao falhou. Veja a mensagem acima.
pause
exit /b 1
