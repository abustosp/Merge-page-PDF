$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$distPath = Join-Path $projectRoot "Ejecutable"
$workPath = Join-Path $projectRoot "temp_build"

# Detectar PyInstaller
$pyinstaller = "pyinstaller"
if (Test-Path "$projectRoot\venv\Scripts\pyinstaller.exe") {
    $pyinstaller = "$projectRoot\venv\Scripts\pyinstaller.exe"
}

if ($pyinstaller -eq "pyinstaller" -and -not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
    throw "Error: pyinstaller no encontrado. Activa el venv o instala pyinstaller."
}

Write-Host "=== Compilando Consolidar PDF ===" -ForegroundColor Cyan

$iconPath = Join-Path $projectRoot "bin\ABP-blanco-en-fondo-negro.ico"

& $pyinstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --distpath "$distPath" `
  --workpath "$workPath" `
  --specpath "$workPath" `
  --name "ConsolidarPDF" `
  --icon "$iconPath" `
  ".\gui_consolidar.py"

if (-not $?) {
    throw "Error durante la compilación con PyInstaller"
}

Write-Host "=== Copiando archivos adicionales ===" -ForegroundColor Cyan

$distBinPath = Join-Path $distPath "bin"
if (-not (Test-Path $distBinPath)) {
    New-Item -ItemType Directory -Force -Path $distBinPath | Out-Null
}
Copy-Item "$projectRoot\bin\*" $distBinPath -Recurse -Force

Write-Host "=== Build completado ===" -ForegroundColor Green
Write-Host "Ejecutable en: $distPath\ConsolidarPDF.exe"
