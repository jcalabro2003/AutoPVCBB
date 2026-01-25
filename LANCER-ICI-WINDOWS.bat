@echo off
setlocal enabledelayedexpansion

echo ============================================
echo   Convertisseur DocX vers LaTeX/PDF
echo ============================================
echo Repertoire de travail: %CD%
echo.

cd /d "%~dp0"

REM Verifier si Python est installe
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ATTENTION] Python n'est pas installe ou n'est pas dans le PATH
    echo.
    echo Installation automatique de Python en cours...
    echo.
    
    call :install_python
    
    if !errorlevel! neq 0 (
        echo.
        echo [ERREUR] L'installation de Python a echoue
        pause
        exit /b 1
    )
    
    echo.
    echo Python a ete installe avec succes!
    echo Redemarrage du script...
    echo.
    timeout /t 3 >nul
    
    REM Actualiser les variables d'environnement
    call :refresh_env
    
    REM Verifier a nouveau Python
    python --version >nul 2>&1
    if !errorlevel! neq 0 (
        echo [ERREUR] Python est installe mais pas encore disponible dans le PATH
        echo Veuillez fermer cette fenetre et relancer le script
        pause
        exit /b 1
    )
)

echo Python detecte: 
python --version
echo.

REM Verifier et creer l'environnement virtuel si necessaire
if not exist "venv" (
    echo Creation de l'environnement virtuel...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [ERREUR] Impossible de creer l'environnement virtuel
        pause
        exit /b 1
    )
    echo.
)

REM Activer l'environnement virtuel
echo Activation de l'environnement virtuel...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo [ERREUR] Impossible d'activer l'environnement virtuel
    pause
    exit /b 1
)

REM Installer/Mettre a jour les dependances
echo.
echo Verification des dependances...
pip install -q --upgrade pip

if not exist "requirements.txt" (
    echo [ATTENTION] Fichier requirements.txt introuvable
    echo Creation d'un requirements.txt minimal...
    (
        echo python-docx>=0.8.11
        echo cohere>=5.0
        echo tkinterdnd2>=0.3.0
    ) > requirements.txt
)

echo Installation des dependances depuis requirements.txt...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERREUR] Echec de l'installation des dependances
    echo Verifiez le contenu de requirements.txt
    echo.
    echo Contenu actuel de requirements.txt:
    type requirements.txt
    echo.
    pause
    exit /b 1
)

REM Verifier si LaTeX est installe
echo Verification de LaTeX...
pdflatex --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ATTENTION] LaTeX n'est pas installe
    echo.
    echo Installation automatique de MiKTeX en cours...
    echo Cette operation peut prendre 10-20 minutes...
    echo.
    
    call :install_latex
    
    if !errorlevel! neq 0 (
        echo.
        echo [ATTENTION] L'installation de LaTeX a echoue
        echo L'application fonctionnera mais sans generation de PDF
        echo.
        timeout /t 5
    ) else (
        echo.
        echo MiKTeX a ete installe avec succes!
        echo.
        
        REM Actualiser les variables d'environnement
        call :refresh_env
        
        pdflatex --version >nul 2>&1
        if !errorlevel! neq 0 (
            echo [ATTENTION] MiKTeX est installe mais pas encore disponible
            echo Vous devrez peut-etre redemarrer votre ordinateur
            echo.
        )
    )
) else (
    echo LaTeX detecte: 
    pdflatex --version | findstr "pdfTeX"
    echo.
)

REM Lancer l'application
echo.
echo ============================================
echo   Lancement de l'application...
echo ============================================
echo.
python app.py

REM Pause si erreur
if %errorlevel% neq 0 (
    echo.
    echo [ERREUR] L'application s'est terminee avec une erreur.
    pause
)

REM Desactiver l'environnement virtuel
deactivate
exit /b 0

REM ============================================
REM Fonction: Installation de Python
REM ============================================
:install_python
    echo Telechargement de Python 3.12...
    echo.
    
    REM Determiner l'architecture du systeme
    if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
        set PYTHON_ARCH=amd64
        set PYTHON_URL=https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe
    ) else (
        set PYTHON_ARCH=win32
        set PYTHON_URL=https://www.python.org/ftp/python/3.12.0/python-3.12.0.exe
    )
    
    set PYTHON_INSTALLER=%TEMP%\python_installer.exe
    
    REM Telecharger Python avec PowerShell
    echo Telechargement depuis: %PYTHON_URL%
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_INSTALLER%'}"
    
    if not exist "%PYTHON_INSTALLER%" (
        echo [ERREUR] Echec du telechargement de Python
        exit /b 1
    )
    
    echo.
    echo Installation de Python...
    echo Cette operation peut prendre quelques minutes...
    echo.
    
    REM Installer Python en mode silencieux avec ajout au PATH
    "%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
    
    REM Attendre la fin de l'installation
    timeout /t 10 >nul
    
    REM Nettoyer
    if exist "%PYTHON_INSTALLER%" del "%PYTHON_INSTALLER%"
    
    exit /b 0

REM ============================================
REM Fonction: Installation de LaTeX (MiKTeX)
REM ============================================
:install_latex
    echo Telechargement de MiKTeX (installation minimale)...
    echo.
    
    REM URL de MiKTeX portable/basic (plus leger)
    set MIKTEX_URL=https://miktex.org/download/ctan/systems/win32/miktex/setup/windows-x64/basic-miktex-24.1-x64.exe
    set MIKTEX_INSTALLER=%TEMP%\miktex_installer.exe
    
    REM Telecharger MiKTeX avec PowerShell
    echo Telechargement depuis miktex.org...
    echo ATTENTION: Le fichier fait environ 250 MB, soyez patient...
    echo.
    
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $ProgressPreference = 'SilentlyContinue'; Write-Host 'Telechargement en cours...'; Invoke-WebRequest -Uri '%MIKTEX_URL%' -OutFile '%MIKTEX_INSTALLER%'; Write-Host 'Telechargement termine!'}"
    
    if not exist "%MIKTEX_INSTALLER%" (
        echo [ERREUR] Echec du telechargement de MiKTeX
        exit /b 1
    )
    
    echo.
    echo Installation de MiKTeX...
    echo Cette operation peut prendre 10-15 minutes...
    echo Une fenetre d'installation peut s'ouvrir brievement.
    echo.
    
    REM Installer MiKTeX en mode shared (accessible a tous les utilisateurs)
    REM Options: --unattended = silencieux, --shared = installation partagee
    "%MIKTEX_INSTALLER%" --unattended --shared --package-set=basic
    
    if %errorlevel% neq 0 (
        echo.
        echo Tentative d'installation en mode utilisateur local...
        "%MIKTEX_INSTALLER%" --unattended --package-set=basic
    )
    
    REM Attendre la fin de l'installation
    timeout /t 15 >nul
    
    REM Configurer MiKTeX pour installer automatiquement les packages manquants
    echo Configuration de MiKTeX...
    
    REM Chercher mpm.exe (MiKTeX Package Manager)
    set "MIKTEX_BIN="
    if exist "%ProgramFiles%\MiKTeX\miktex\bin\x64\mpm.exe" (
        set "MIKTEX_BIN=%ProgramFiles%\MiKTeX\miktex\bin\x64"
    )
    if exist "%LOCALAPPDATA%\Programs\MiKTeX\miktex\bin\x64\mpm.exe" (
        set "MIKTEX_BIN=%LOCALAPPDATA%\Programs\MiKTeX\miktex\bin\x64"
    )
    
    if not "!MIKTEX_BIN!"=="" (
        echo Configuration de l'installation automatique des packages...
        "!MIKTEX_BIN!\initexmf.exe" --set-config-value [MPM]AutoInstall=1 >nul 2>&1
        
        echo Installation des packages essentiels...
        "!MIKTEX_BIN!\mpm.exe" --install=geometry >nul 2>&1
        "!MIKTEX_BIN!\mpm.exe" --install=fancyhdr >nul 2>&1
        "!MIKTEX_BIN!\mpm.exe" --install=multicol >nul 2>&1
        "!MIKTEX_BIN!\mpm.exe" --install=graphics >nul 2>&1
        "!MIKTEX_BIN!\mpm.exe" --install=float >nul 2>&1
        "!MIKTEX_BIN!\mpm.exe" --install=varwidth >nul 2>&1
        "!MIKTEX_BIN!\mpm.exe" --install=eurosym >nul 2>&1
    )
    
    REM Nettoyer
    if exist "%MIKTEX_INSTALLER%" del "%MIKTEX_INSTALLER%"
    
    echo.
    echo Installation de MiKTeX terminee!
    exit /b 0

REM ============================================
REM Fonction: Actualiser les variables d'environnement
REM ============================================
:refresh_env
    REM Relire les variables d'environnement PATH depuis le registre
    for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v PATH 2^>nul') do set "USER_PATH=%%b"
    for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH 2^>nul') do set "SYSTEM_PATH=%%b"
    
    set "PATH=%SYSTEM_PATH%;%USER_PATH%"
    exit /b 0