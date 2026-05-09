$ErrorActionPreference = "Stop"

Set-Location (Split-Path $PSScriptRoot -Parent)

$iconPath = ".\src\attendance\resources\attendance.ico"
$iconArgs = @()
if (Test-Path $iconPath) {
  $iconArgs = @("--icon", $iconPath)
}

$brandAssetCandidates = @(
  ".\src\attendance\resources\attendance_brand.png",
  ".\src\attendance\resources\attendance_brand.gif"
)
$brandDataArgs = @()
foreach ($brandAsset in $brandAssetCandidates) {
  if (Test-Path $brandAsset) {
    $brandDataArgs = @("--add-data", "${brandAsset};resources")
    break
  }
}

if (!(Test-Path ".\venv\Scripts\python.exe")) {
  python -m venv venv
}

. .\venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install -r requirements-dev.txt

$distDir = ".\dist"
$launcherDir = Join-Path $distDir "attendance_launcher"
$portableZipPath = Join-Path $distDir "attendance_launcher_portable.zip"
$bundledAssetsDir = Join-Path $launcherDir "bundled_assets"
$bundledAppZip = Join-Path $bundledAssetsDir "attendance_windows.zip"
$bundledMetadata = Join-Path $bundledAssetsDir "attendance_release.json"
$appZipPath = Join-Path $distDir "attendance_windows.zip"
$versionFile = ".\src\attendance\version.py"
$legacyExePath = Join-Path $distDir "attendance_launcher.exe"

if (!(Test-Path $appZipPath)) {
  & ".\scripts\build_windows.ps1"
}

if (!(Test-Path $distDir)) {
  New-Item -ItemType Directory -Path $distDir | Out-Null
}
if (Test-Path ".\build\attendance_launcher") { Remove-Item ".\build\attendance_launcher" -Recurse -Force }
if (Test-Path $launcherDir) { Remove-Item $launcherDir -Recurse -Force }
if (Test-Path $portableZipPath) { Remove-Item $portableZipPath -Force }
if (Test-Path $legacyExePath) { Remove-Item $legacyExePath -Force }

$pyiArgs = @(
  "--noconfirm",
  "--clean",
  "--onedir",
  "--windowed",
  "--name", "attendance_launcher"
)

if ($iconArgs.Count -gt 0) {
  $pyiArgs += $iconArgs
}

if ($brandDataArgs.Count -gt 0) {
  $pyiArgs += $brandDataArgs
}

$pyiArgs += "attendance_launcher.py"

python -m PyInstaller @pyiArgs

if (!(Test-Path (Join-Path $launcherDir "attendance_launcher.exe"))) {
  throw "No se genero el ejecutable: dist\attendance_launcher\attendance_launcher.exe"
}

$versionMatch = Select-String -Path $versionFile -Pattern '__version__ = "([^"]+)"'
if (!$versionMatch.Matches.Count) {
  throw "No se pudo leer la version desde src\\attendance\\version.py"
}
$version = $versionMatch.Matches[0].Groups[1].Value

New-Item -ItemType Directory -Path $bundledAssetsDir -Force | Out-Null
Copy-Item $appZipPath -Destination $bundledAppZip -Force

$metadata = @{
  tag = "v$version"
  asset_name = "attendance_windows.zip"
} | ConvertTo-Json -Depth 3
[System.IO.File]::WriteAllText(
  $bundledMetadata,
  $metadata,
  (New-Object System.Text.UTF8Encoding($false))
)

Compress-Archive -Path $launcherDir -DestinationPath $portableZipPath -Force

Write-Host "Launcher onedir generado en: dist\attendance_launcher\"
Write-Host "ZIP portable generado en: dist\attendance_launcher_portable.zip"
