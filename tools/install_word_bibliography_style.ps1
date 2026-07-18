#requires -Version 5.1

[CmdletBinding()]
param([switch]$Force)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'word\Yibin-GB-T-7714-Numeric.xsl'
$styleDirectory = Join-Path $env:APPDATA 'Microsoft\Bibliography\Style'
$target = Join-Path $styleDirectory (Split-Path -Leaf $source)

if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    throw "Bundled Word bibliography style not found: $source"
}

$gbCandidates = @(
    (Join-Path $styleDirectory 'GB.XSL'),
    (Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\Bibliography\Style\GB.XSL')
)
if (-not ($gbCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })) {
    throw 'Microsoft Word GB.XSL was not found. Install the Word bibliography components first.'
}

New-Item -ItemType Directory -Force $styleDirectory | Out-Null
if (Test-Path -LiteralPath $target -PathType Leaf) {
    $same = (Get-FileHash -Algorithm SHA256 -LiteralPath $source).Hash -eq
        (Get-FileHash -Algorithm SHA256 -LiteralPath $target).Hash
    if ($same) {
        Write-Host "Word bibliography style is already installed: $target"
        return
    }
    if (-not $Force) {
        throw "A different style already exists at $target. Re-run with -Force to replace it."
    }
}

Copy-Item -LiteralPath $source -Destination $target -Force
Write-Host "Installed Word bibliography style: $target"
Write-Host 'Restart Microsoft Word before selecting the new style.'
