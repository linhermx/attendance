param(
  [switch]$Release
)

$ErrorActionPreference = "Stop"

Set-Location (Split-Path $PSScriptRoot -Parent)

$iconPath = ".\src\attendance\resources\attendance.ico"
$iconArgs = @()
if (Test-Path $iconPath) {
  $iconArgs = @("--add-data", "src\attendance\resources\attendance.ico;resources", "--icon", $iconPath)
}

if (!(Test-Path ".\venv\Scripts\python.exe")) {
  python -m venv venv
}

. .\venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install -r requirements-dev.txt

python -c "import numpy, pandas; print(f'Runtime de datos verificado: NumPy {numpy.__version__}, pandas {pandas.__version__}')"
if ($LASTEXITCODE -ne 0) {
  throw "El entorno de build no puede importar NumPy y pandas. Recrea venv antes de empaquetar."
}

$distDir = ".\dist"
$buildDir = ".\build\attendance_windows"
$appDir = Join-Path $distDir "attendance_windows"
$zipPath = Join-Path $distDir "attendance_windows.zip"
$aliasZipPath = Join-Path $distDir "attendance.zip"
$legacyExePath = Join-Path $distDir "attendance_windows.exe"

if (!(Test-Path $distDir)) {
  New-Item -ItemType Directory -Path $distDir | Out-Null
}
if (Test-Path $buildDir) { Remove-Item $buildDir -Recurse -Force }
if (Test-Path $appDir) { Remove-Item $appDir -Recurse -Force }
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
if (Test-Path $aliasZipPath) { Remove-Item $aliasZipPath -Force }
if (Test-Path $legacyExePath) { Remove-Item $legacyExePath -Force }

$pyiArgs = @(
  "--noconfirm",
  "--clean",
  "--onedir",
  "--windowed",
  "--paths", ".\src",
  "--hidden-import", "attendance",
  "--hidden-import", "attendance.gui",
  "--hidden-import", "attendance.core",
  "--name", "attendance_windows"
)

if ($iconArgs.Count -gt 0) {
  $pyiArgs += $iconArgs
}

$pyiArgs += "attendance_gui.py"

python -m PyInstaller @pyiArgs

$appExePath = Join-Path $appDir "attendance_windows.exe"
if (!(Test-Path $appExePath)) {
  throw "No se genero el ejecutable: dist\attendance_windows\attendance_windows.exe"
}

$smokeProcess = Start-Process `
  -FilePath $appExePath `
  -ArgumentList "--smoke-test" `
  -WindowStyle Hidden `
  -PassThru
if (!$smokeProcess.WaitForExit(15000)) {
  Stop-Process -Id $smokeProcess.Id -Force -ErrorAction SilentlyContinue
  throw "El ejecutable empaquetado no completo la autoprueba en 15 segundos."
}
if ($smokeProcess.ExitCode -ne 0) {
  throw "El ejecutable empaquetado fallo la autoprueba con codigo $($smokeProcess.ExitCode)."
}
Write-Host "Autoprueba del ejecutable completada correctamente."

Compress-Archive -Path $appDir -DestinationPath $zipPath -Force

Write-Host "Paquete onedir generado en: dist\attendance_windows\"
Write-Host "ZIP de actualizacion generado en: dist\attendance_windows.zip"

if ($Release) {
  Copy-Item $zipPath -Destination $aliasZipPath -Force
  Write-Host "Alias listo: dist\attendance.zip"
}
