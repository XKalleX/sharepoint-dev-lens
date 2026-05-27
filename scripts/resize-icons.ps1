#Requires -Version 5.1
<#
.SYNOPSIS
    Resizes src/icons/icon_source.png (256x256) to icon16.png, icon48.png, icon128.png
    using high-quality bicubic resampling. Output is lossless PNG.
.DESCRIPTION
    No npm, no external tools – uses only the built-in System.Drawing assembly.
    Source file: src/icons/icon_source.png  (must be 256×256 or any square PNG)
.EXAMPLE
    .\scripts\resize-icons.ps1
#>

Add-Type -AssemblyName System.Drawing

$root   = $PSScriptRoot | Split-Path -Parent
$src    = Join-Path $root 'src\icons\icon_source.png'
$outDir = Join-Path $root 'src\icons'

if (-not (Test-Path $src)) {
    Write-Host " [ERROR] Source not found: $src" -ForegroundColor Red
    Write-Host "         Bitte kopiere dein 256x256-PNG nach src\icons\icon_source.png" -ForegroundColor Yellow
    exit 1
}

$sizes = @(16, 48, 128)

Write-Host ""
Write-Host " SharePoint Dev Lens – Icon Resize" -ForegroundColor Cyan
Write-Host " ==================================`n"

$sourceImage = [System.Drawing.Bitmap]::FromFile($src)
Write-Host " Quelle: $($sourceImage.Width)x$($sourceImage.Height) px – $src"
Write-Host ""

foreach ($size in $sizes) {
    $dest = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($dest)

    # Highest-quality downscale settings
    $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality

    # Transparent background first
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.DrawImage($sourceImage, 0, 0, $size, $size)
    $g.Dispose()

    $outPath = Join-Path $outDir "icon${size}.png"

    # Save as PNG (lossless, supports transparency)
    $dest.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $dest.Dispose()

    Write-Host " ✓  icon${size}.png  ($outPath)" -ForegroundColor Green
}

$sourceImage.Dispose()

Write-Host ""
Write-Host " Fertig! Icons in src\icons\ aktualisiert." -ForegroundColor Green
Write-Host ""
