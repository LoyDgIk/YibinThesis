#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$TectonicPath,
    [string]$BiberPath,
    [string]$PandocPath,
    [string]$PythonPath,
    [switch]$SkipPythonEnvironment
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$ToolRoot = Join-Path $ProjectRoot '.tools'

function Resolve-ToolSource {
    param(
        [string]$ExplicitPath,
        [string]$EnvironmentVariable,
        [string[]]$CommandNames
    )

    foreach ($candidate in @($ExplicitPath, [Environment]::GetEnvironmentVariable($EnvironmentVariable))) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    foreach ($commandName in $CommandNames) {
        $command = Get-Command $commandName -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command) {
            return $command.Source
        }
    }
    return $null
}

function Install-LocalTool {
    param(
        [string]$Name,
        [string]$Source,
        [string]$RelativeTarget
    )

    $target = Join-Path $ProjectRoot $RelativeTarget
    if (Test-Path -LiteralPath $target -PathType Leaf) {
        Write-Host "[OK] $Name already exists: $target"
        return
    }
    if ([string]::IsNullOrWhiteSpace($Source)) {
        throw "$Name was not found. Supply its explicit path or the matching YIBINTHESIS_* environment variable."
    }
    New-Item -ItemType Directory -Force (Split-Path -Parent $target) | Out-Null
    Copy-Item -LiteralPath $Source -Destination $target -Force
    Write-Host "[COPIED] $Name -> $target"
}

New-Item -ItemType Directory -Force $ToolRoot | Out-Null

$tectonic = Resolve-ToolSource -ExplicitPath $TectonicPath -EnvironmentVariable 'YIBINTHESIS_TECTONIC' -CommandNames @('tectonic')
$biber = Resolve-ToolSource -ExplicitPath $BiberPath -EnvironmentVariable 'YIBINTHESIS_BIBER' -CommandNames @('biber')
$pandoc = Resolve-ToolSource -ExplicitPath $PandocPath -EnvironmentVariable 'YIBINTHESIS_PANDOC' -CommandNames @('pandoc')

Install-LocalTool -Name 'Tectonic 0.16.9' -Source $tectonic -RelativeTarget '.tools\tectonic\tectonic.exe'
Install-LocalTool -Name 'Biber 2.17' -Source $biber -RelativeTarget '.tools\biber-2.17\biber.exe'
Install-LocalTool -Name 'Pandoc 3.x' -Source $pandoc -RelativeTarget '.tools\pandoc-3.9.0.2\pandoc.exe'

if (-not $SkipPythonEnvironment) {
    $venvPython = Join-Path $ToolRoot 'venv\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
        $python = Resolve-ToolSource -ExplicitPath $PythonPath -EnvironmentVariable 'YIBINTHESIS_PYTHON' -CommandNames @('python', 'python3')
        if ([string]::IsNullOrWhiteSpace($python)) {
            throw 'Python was not found; pass -PythonPath or use -SkipPythonEnvironment.'
        }
        & $python -m venv (Join-Path $ToolRoot 'venv')
        if ($LASTEXITCODE -ne 0) {
            throw "Python virtual environment creation failed with exit code $LASTEXITCODE."
        }
    }
    & $venvPython -m pip install -r (Join-Path $ProjectRoot 'requirements-word.txt')
    if ($LASTEXITCODE -ne 0) {
        throw "Word dependency installation failed with exit code $LASTEXITCODE."
    }
    Write-Host "[OK] Python Word environment: $venvPython"
}

Write-Host 'Project-local toolchain setup completed.'
Write-Host 'Run: .\build.ps1 doctor'
