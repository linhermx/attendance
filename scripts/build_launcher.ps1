$ErrorActionPreference = "Stop"

Set-Location (Split-Path $PSScriptRoot -Parent)

$iconPath = ".\src\attendance\resources\attendance.ico"
$iconArgs = @()
if (Test-Path $iconPath) {
  $iconArgs = @("--icon", $iconPath)
}

if (!(Test-Path ".\venv\Scripts\python.exe")) {
  python -m venv venv
}

. .\venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install -r requirements-dev.txt

if (Test-Path ".\dist\launcher") { Remove-Item ".\dist\launcher" -Recurse -Force }

$pyiArgs = @(
  "--noconfirm",
  "--clean",
  "--onefile",
  "--windowed",
  "--name", "attendance_launcher",
  "attendance_launcher.py"
)

if ($iconArgs.Count -gt 0) {
  $pyiArgs = $pyiArgs[0..3] + $iconArgs + $pyiArgs[4..($pyiArgs.Count - 1)]
}

pyinstaller @pyiArgs

Write-Host "Launcher EXE generado en: dist\attendance_launcher.exe"
