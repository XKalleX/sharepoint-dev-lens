#Requires -Version 5.1
<#
.SYNOPSIS
    Builds SharePoint Dev Lens extension packages for Chrome and Edge.
.DESCRIPTION
    Copies src/ to dist/chrome/ and dist/edge/, then places the correct
    manifest.json in each folder. No npm or build tools required.
.EXAMPLE
    .\build.ps1
#>

$ErrorActionPreference = 'Stop'
$root    = $PSScriptRoot
$src     = Join-Path $root 'src'
$distCh  = Join-Path $root 'dist\chrome'
$distEd  = Join-Path $root 'dist\edge'
$mfCh    = Join-Path $root 'manifests\manifest.chrome.json'
$mfEd    = Join-Path $root 'manifests\manifest.edge.json'

Write-Host ""
Write-Host " SharePoint Dev Lens - Build"
Write-Host " ===========================`n"

# 1. Icons
$iconSource = Join-Path $src 'icons\icon_source.png'
$icon16     = Join-Path $src 'icons\icon16.png'

if (Test-Path $iconSource) {
    Write-Host " [1/3] icon_source.png gefunden – skaliere auf 16/48/128 px..."
    & (Join-Path $root 'scripts\resize-icons.ps1')
    if ($LASTEXITCODE -ne 0) { exit 1 }
} elseif (-not (Test-Path $icon16)) {
    Write-Host " [1/3] Kein icon_source.png – generiere Standard-Icons..."
    node (Join-Path $root 'scripts\generate-icons.js')
} else {
    Write-Host " [1/3] Icons vorhanden, ueberspringe Generierung."
}

# 2. Chrome
Write-Host " [2/3] Building dist\chrome\ ..."
if (Test-Path $distCh) { Remove-Item $distCh -Recurse -Force }
New-Item -ItemType Directory -Path $distCh -Force | Out-Null
Copy-Item -Path (Join-Path $src '*') -Destination $distCh -Recurse -Force
Copy-Item -Path $mfCh -Destination (Join-Path $distCh 'manifest.json') -Force

# 3. Edge
Write-Host " [3/3] Building dist\edge\ ..."
if (Test-Path $distEd) { Remove-Item $distEd -Recurse -Force }
New-Item -ItemType Directory -Path $distEd -Force | Out-Null
Copy-Item -Path (Join-Path $src '*') -Destination $distEd -Recurse -Force
Copy-Item -Path $mfEd -Destination (Join-Path $distEd 'manifest.json') -Force

Write-Host ""
Write-Host " Done!" -ForegroundColor Green
Write-Host ""
Write-Host " Chrome  ->  dist\chrome\"
Write-Host " Edge    ->  dist\edge\"
Write-Host ""
Write-Host " To load in Chrome:  chrome://extensions/  -> Developer mode -> Load unpacked -> select dist\chrome"
Write-Host " To load in Edge:    edge://extensions/    -> Developer mode -> Load unpacked -> select dist\edge"
Write-Host ""
