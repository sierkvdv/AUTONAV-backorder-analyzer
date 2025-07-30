@echo off
echo ========================================
echo Navision Backorder Analyzer
echo ========================================
echo.

REM Controleer of Python geïnstalleerd is
python --version >nul 2>&1
if errorlevel 1 (
    echo FOUT: Python is niet geïnstalleerd of niet gevonden in PATH
    echo Installeer Python 3.8+ en probeer opnieuw
    pause
    exit /b 1
)

echo Python gevonden, starten van analyse...
echo.

REM Voer het script uit
python backorder_analyzer.py

echo.
echo ========================================
echo Analyse voltooid!
echo ========================================
echo.
echo Output bestand: Output/Backorder_Analyse.xlsx
echo Log bestand: backorder_analyzer.log
echo.
pause