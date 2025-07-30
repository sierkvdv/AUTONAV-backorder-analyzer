# Navision Backorder Analyzer - PowerShell Launcher
# ================================================

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Navision Backorder Analyzer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Controleer of Python geïnstalleerd is
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✓ Python gevonden: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "✗ FOUT: Python is niet geïnstalleerd of niet gevonden in PATH" -ForegroundColor Red
    Write-Host "Installeer Python 3.8+ en probeer opnieuw" -ForegroundColor Yellow
    Read-Host "Druk op Enter om af te sluiten"
    exit 1
}

# Controleer of het hoofdscript bestaat
if (-not (Test-Path "backorder_analyzer.py")) {
    Write-Host "✗ FOUT: backorder_analyzer.py niet gevonden!" -ForegroundColor Red
    Write-Host "Zorg dat je dit script in dezelfde map uitvoert als backorder_analyzer.py" -ForegroundColor Yellow
    Read-Host "Druk op Enter om af te sluiten"
    exit 1
}

Write-Host "✓ Script gevonden, starten van analyse..." -ForegroundColor Green
Write-Host ""

# Voer het script uit
try {
    python backorder_analyzer.py
    $exitCode = $LASTEXITCODE
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    if ($exitCode -eq 0) {
        Write-Host "✓ Analyse succesvol voltooid!" -ForegroundColor Green
    } else {
        Write-Host "⚠ Analyse voltooid met waarschuwingen" -ForegroundColor Yellow
    }
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Toon output bestanden
    if (Test-Path "Output/Backorder_Analyse.xlsx") {
        Write-Host "✓ Output bestand: Output/Backorder_Analyse.xlsx" -ForegroundColor Green
    } else {
        Write-Host "✗ Output bestand niet gevonden" -ForegroundColor Red
    }
    
    if (Test-Path "backorder_analyzer.log") {
        Write-Host "✓ Log bestand: backorder_analyzer.log" -ForegroundColor Green
    } else {
        Write-Host "✗ Log bestand niet gevonden" -ForegroundColor Red
    }
    
} catch {
    Write-Host "✗ FOUT bij uitvoeren van script: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Read-Host "Druk op Enter om af te sluiten"