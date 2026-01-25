@echo off
echo ============================================
echo   Build Executable Windows
echo ============================================
echo.

cd /d "%~dp0"

echo Verification de Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERREUR] Python n'est pas installe
    echo Installez Python depuis https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Installation des dependances...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Lancement du build...
python build_executable.py

pause
