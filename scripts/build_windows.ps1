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

if (!(Test-Path (Join-Path $appDir "attendance_windows.exe"))) {
  throw "No se genero el ejecutable: dist\attendance_windows\attendance_windows.exe"
}

Compress-Archive -Path $appDir -DestinationPath $zipPath -Force

Write-Host "Paquete onedir generado en: dist\attendance_windows\"
Write-Host "ZIP de actualizacion generado en: dist\attendance_windows.zip"

if ($Release) {
  Copy-Item $zipPath -Destination $aliasZipPath -Force
  Write-Host "Alias listo: dist\attendance.zip"
}
