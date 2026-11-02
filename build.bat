@echo off
setlocal
set MODE=onedir
cd /d "%~dp0"

rem ============================================================
rem  Construction de resilia2.exe (a lancer sur un poste Windows)
rem  Prerequis : Python 3.10+ (lanceur "py") avec acces a pypi.org
rem  Resultat  : dist\resilia2.exe
rem
rem  MODE=onefile : un seul fichier (par defaut)
rem  MODE=onedir  : dossier dist\resilia2\ (demarrage plus rapide,
rem                 utile si l'antivirus ralentit la version onefile)
rem ============================================================
if "%MODE%"=="" set MODE=onefile

where py >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python ^(commande "py"^) introuvable.
    exit /b 1
)

if not exist .venv (
    echo [BUILD] Creation de l'environnement virtuel...
    py -3 -m venv .venv || exit /b 1
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip || exit /b 1
pip install -r requirements.txt || exit /b 1

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [BUILD] PyInstaller (%MODE%)...
pyinstaller --noconfirm --clean --%MODE% --windowed ^
    --name resilia2 ^
    --icon resilia.ico ^
    --add-data "web;web" ^
    --add-data "resilia.ico;." ^
    --collect-all playwright ^
    --hidden-import win32timezone ^
    resilia2.py || exit /b 1

echo.
if "%MODE%"=="onefile" (
    echo [OK] dist\resilia2.exe
) else (
    echo [OK] dist\resilia2\resilia2.exe ^(copier tout le dossier^)
)
endlocal
