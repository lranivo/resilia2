@echo off
REM Compile resilia2.iss en un seul exe (Inno Setup 6 requis sur CE poste
REM de développement uniquement : https://jrsoftware.org/isdl.php)
setlocal
cd /d "%~dp0"

if not exist "dist\resilia2\resilia2.exe" (
  echo [ERREUR] dist\resilia2\resilia2.exe introuvable. Lancez d'abord PyInstaller.
  pause & exit /b 1
)
if not exist "resilia.ico" (
  echo [ERREUR] resilia.ico introuvable a la racine du projet.
  pause & exit /b 1
)

set "ISCC="
for %%P in (
  "%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe"
  "%ProgramFiles%\Inno Setup 7\ISCC.exe"
  "%LOCALAPPDATA%\Programs\Inno Setup 7\ISCC.exe"
) do if exist %%P set "ISCC=%%~P"

if not defined ISCC (
  echo [ERREUR] Inno Setup 6 introuvable. Installez-le puis relancez.
  pause & exit /b 1
)

"%ISCC%" resilia2.iss
if errorlevel 1 ( echo [ERREUR] La compilation a echoue. & pause & exit /b 1 )

echo.
echo [OK] Installateur cree dans le dossier "installer".
pause
