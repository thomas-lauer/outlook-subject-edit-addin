param(
    [string]$ManifestPath = ".\manifest.xml",
    [string]$AddinFolderName = "OutlookSubjectEditAddin"
)

Write-Host "=== Outlook Add-In Sideload Installation ===" -ForegroundColor Cyan

# 1. Manifest prüfen
if (-not (Test-Path $ManifestPath)) {
    Write-Host "Manifest nicht gefunden: $ManifestPath" -ForegroundColor Red
    exit 1
}

$ManifestFullPath = (Resolve-Path $ManifestPath).Path
Write-Host "Verwende Manifest: $ManifestFullPath"

# 2. Zielpfad für Sideload-Add-Ins unter Windows (Office 365, Office 2019+)
#    Das ist ein Standardpfad, den Outlook für Developer-Addins überwacht.
$wefBase = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
$targetFolder = Join-Path $wefBase $AddinFolderName

Write-Host "Zielordner: $targetFolder"

if (-not (Test-Path $targetFolder)) {
    Write-Host "Zielordner existiert nicht. Erstelle Ordner..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
}

# 3. Manifest kopieren
$targetManifest = Join-Path $targetFolder "manifest.xml"

Copy-Item -Path $ManifestFullPath -Destination $targetManifest -Force

Write-Host "Manifest wurde nach:" -ForegroundColor Green
Write-Host "  $targetManifest" -ForegroundColor Green

Write-Host ""
Write-Host "Bitte Outlook komplett schließen und neu starten," -ForegroundColor Cyan
Write-Host "dann sollte das Add-In in einer geöffneten E-Mail unter den Add-Ins erscheinen." -ForegroundColor Cyan
