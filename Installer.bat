@echo off
REM ============================================================
REM  Installation du programme excel-to-quadra (sans droits admin)
REM  A lancer UNE SEULE FOIS, par double-clic.
REM ============================================================
cd /d "%~dp0"

echo.
echo --- Verification de Python ---
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo Installer Python depuis https://www.python.org/downloads/
    echo en cochant la case "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)
python --version

echo.
echo --- Installation du programme et de ses dependances ---
python -m pip install --user . 
if errorlevel 1 (
    echo.
    echo [ERREUR] L'installation a echoue. Voir le message ci-dessus
    echo ou contacter Sebastien Grison.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Installation terminee avec succes.
echo  Vous pouvez maintenant utiliser Lancer.bat
echo ============================================================
echo.
pause
