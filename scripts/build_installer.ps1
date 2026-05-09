$ErrorActionPreference = "Stop"

Set-Location (Split-Path $PSScriptRoot -Parent)

function Get-IsccPath {
  $command = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
  if ($command) {
    return $command.Source
  }

  $candidates = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
  )

  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path $candidate)) {
      return $candidate
    }
  }

  throw "No se encontro ISCC.exe. Instala Inno Setup 6 o agregalo al PATH."
}

$iscc = Get-IsccPath

& ".\scripts\build_windows.ps1"
& ".\scripts\build_launcher.ps1"

$versionMatch = Select-String -Path ".\src\attendance\version.py" -Pattern '__version__ = "([^"]+)"'
if (!$versionMatch.Matches.Count) {
  throw "No se pudo leer la version desde src\\attendance\\version.py"
}
$version = $versionMatch.Matches[0].Groups[1].Value

$distInstallerDir = ".\dist\installer"
if (!(Test-Path $distInstallerDir)) {
  New-Item -ItemType Directory -Path $distInstallerDir | Out-Null
}

$issPath = ".\installer\attendance_setup.iss"
$launcherSourceDir = (Resolve-Path ".\dist\attendance_launcher").Path
$outputDir = (Resolve-Path $distInstallerDir).Path
$setupIcon = (Resolve-Path ".\src\attendance\resources\attendance.ico").Path

& $iscc `
  "/DMyAppVersion=$version" `
  "/DLauncherSourceDir=$launcherSourceDir" `
  "/DOutputDir=$outputDir" `
  "/DSetupIconFilePath=$setupIcon" `
  $issPath

Write-Host "Instalador generado en: dist\installer\attendance_setup.exe"
