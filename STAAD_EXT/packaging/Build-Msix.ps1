[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9.-]{3,50}$')]
    [string] $IdentityName,
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $Publisher,
    [Parameter(Mandatory)]
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string] $Version,
    [string] $ApplicationDirectory = (Join-Path $PSScriptRoot '..\dist\STAAD_EXT'),
    [string] $OutputPath = (Join-Path $PSScriptRoot '..\dist\STAAD_EXT-store.msix')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Find-MakeAppx {
    $fromPath = Get-Command MakeAppx.exe -ErrorAction SilentlyContinue
    if ($null -ne $fromPath) {
        return $fromPath.Source
    }
    $kitsRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
    if (Test-Path -LiteralPath $kitsRoot) {
        $candidate = Get-ChildItem -LiteralPath $kitsRoot -Recurse -Filter MakeAppx.exe |
            Where-Object { $_.DirectoryName -like '*\x64' } |
            Sort-Object -Property FullName -Descending |
            Select-Object -First 1
        if ($null -ne $candidate) {
            return $candidate.FullName
        }
    }
    throw 'MakeAppx.exe was not found. Install the Windows 10 or 11 SDK.'
}

function New-Logo {
    param(
        [Parameter(Mandatory)]
        [string] $Path,
        [Parameter(Mandatory)]
        [int] $Width,
        [Parameter(Mandatory)]
        [int] $Height
    )
    Add-Type -AssemblyName System.Drawing
    $bitmap = [System.Drawing.Bitmap]::new($Width, $Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([System.Drawing.Color]::FromArgb(19, 78, 74))
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $fontSize = [Math]::Max(10, [Math]::Floor([Math]::Min($Width, $Height) * 0.34))
        $font = [System.Drawing.Font]::new('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $brush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
        $format = [System.Drawing.StringFormat]::new()
        try {
            $format.Alignment = [System.Drawing.StringAlignment]::Center
            $format.LineAlignment = [System.Drawing.StringAlignment]::Center
            $rectangle = [System.Drawing.RectangleF]::new(0, 0, $Width, $Height)
            $graphics.DrawString('SE', $font, $brush, $rectangle, $format)
        }
        finally {
            $format.Dispose()
            $brush.Dispose()
            $font.Dispose()
        }
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

$versionParts = $Version.Split('.')
if ($versionParts.Count -ne 4 -or @($versionParts | Where-Object { [int64] $_ -gt 65535 }).Count -gt 0) {
    throw 'Each MSIX version component must be between 0 and 65535.'
}
$resolvedApplicationDirectory = (Resolve-Path -LiteralPath $ApplicationDirectory).Path
$executable = Join-Path $resolvedApplicationDirectory 'STAAD_EXT.exe'
if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
    throw "PyInstaller output was not found at '$executable'. Build STAAD_EXT.spec first."
}
$outputDirectory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$resolvedOutputDirectory = (Resolve-Path -LiteralPath $outputDirectory).Path
$resolvedOutputPath = Join-Path $resolvedOutputDirectory (Split-Path -Leaf $OutputPath)
$stagingDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('staad-ext-msix-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stagingDirectory | Out-Null

try {
    Copy-Item -LiteralPath $resolvedApplicationDirectory -Destination (Join-Path $stagingDirectory 'STAAD_EXT') -Recurse
    $assetsDirectory = Join-Path $stagingDirectory 'Assets'
    New-Item -ItemType Directory -Path $assetsDirectory | Out-Null
    New-Logo -Path (Join-Path $assetsDirectory 'Square44x44Logo.png') -Width 44 -Height 44
    New-Logo -Path (Join-Path $assetsDirectory 'Square150x150Logo.png') -Width 150 -Height 150
    New-Logo -Path (Join-Path $assetsDirectory 'StoreLogo.png') -Width 50 -Height 50
    $manifestTemplate = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'AppxManifest.template.xml') -Raw
    $manifest = $manifestTemplate.Replace('{{IDENTITY_NAME}}', [System.Security.SecurityElement]::Escape($IdentityName)).Replace('{{PUBLISHER}}', [System.Security.SecurityElement]::Escape($Publisher)).Replace('{{VERSION}}', $Version)
    Set-Content -LiteralPath (Join-Path $stagingDirectory 'AppxManifest.xml') -Value $manifest -Encoding utf8
    $makeAppx = Find-MakeAppx
    & $makeAppx pack /d $stagingDirectory /p $resolvedOutputPath /o
    if ($LASTEXITCODE -ne 0) {
        throw "MakeAppx.exe failed with exit code $LASTEXITCODE."
    }
    Write-Host 'Created unsigned Microsoft Store submission package:'
    Write-Host $resolvedOutputPath
}
finally {
    if (Test-Path -LiteralPath $stagingDirectory) {
        Remove-Item -LiteralPath $stagingDirectory -Recurse -Force
    }
}
