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

if (Test-Path ".\build") { Remove-Item ".\build" -Recurse -Force }
if (Test-Path ".\dist") { Remove-Item ".\dist" -Recurse -Force }

$pyiArgs = @(
  "--noconfirm",
  "--clean",
  "--onefile",
  "--windowed",
  "--paths", ".\src",
  "--hidden-import", "attendance",
  "--hidden-import", "attendance.gui",
  "--hidden-import", "attendance.core",
  "--name", "attendance_windows",
  "attendance_gui.py"
)

if ($iconArgs.Count -gt 0) {
  $pyiArgs = $pyiArgs[0..6] + $iconArgs + $pyiArgs[7..($pyiArgs.Count - 1)]
}

pyinstaller @pyiArgs

if (!(Test-Path ".\dist\attendance_windows.exe")) {
  throw "No se genero el ejecutable: dist\attendance_windows.exe"
}

Write-Host "EXE generado en: dist\attendance_windows.exe"

if ($Release) {
  Copy-Item ".\dist\attendance_windows.exe" -Destination ".\dist\attendance.exe" -Force
  Write-Host "Alias listo: dist\attendance.exe"
}
