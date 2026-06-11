@echo off
REM ============================================================
REM  Generation des fichiers d'ecritures Quadra
REM  Double-cliquer sur ce fichier a chaque situation.
REM ============================================================

REM --- Nom du fichier de configuration (a adapter si besoin) ---
set CONFIG=config\situation.local.yaml

cd /d "%~dp0"

if not exist "%CONFIG%" (
    echo [ERREUR] Fichier de configuration introuvable : %CONFIG%
    echo Verifier le nom dans la ligne "set CONFIG=" de ce fichier,
    echo ou demander le fichier de configuration a Sebastien Grison.
    echo.
    pause
    exit /b 1
)

echo.
echo --- Generation en cours (config : %CONFIG%) ---
echo.
python -m excel_to_quadra.cli --config "%CONFIG%"
if errorlevel 1 (
    echo.
    echo ************************************************************
    echo  ATTENTION : un probleme a ete detecte (voir messages
    echo  ci-dessus). Ne pas importer les dossiers en anomalie.
    echo ************************************************************
) else (
    echo.
    echo ------------------------------------------------------------
    echo  Termine. Les fichiers sont dans le dossier "sortie".
    echo ------------------------------------------------------------
)
echo.
echo Cette fenetre peut etre fermee apres lecture du rapport.
pause
