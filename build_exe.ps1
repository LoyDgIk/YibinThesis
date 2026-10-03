#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$Python = 'python',
    [string]$OutputDirectory = 'dist',
    [switch]$Clean,
    [switch]$OneFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Root = [System.IO.Path]::GetFullPath($PSScriptRoot)
$workDirectory = Join-Path $Root '.pyinstaller'
if ($Clean -and (Test-Path -LiteralPath $workDirectory)) {
    Remove-Item -LiteralPath $workDirectory -Recurse -Force
}

$dataFiles = @(
    @('build.ps1', '.'),
    @('yibinthesis.project.schema.json', '.'),
    @('yibinthesis.cls', '.'),
    @('yibinthesis-proposal.sty', '.'),
    @('yibinthesis-literature-review.sty', '.'),
    @('assets', 'assets'),
    @('tests', 'tests'),
    @('word', 'word'),
    @('lib/build_word.py', 'lib'),
    @('lib/word_core.py', 'lib'),
    @('lib/word_oxml.py', 'lib'),
    @('lib/word_profiles.py', 'lib'),
    @('lib/build_reference_docx.py', 'lib'),
    @('lib/prepare_signature_assets.py', 'lib'),
    @('lib/audit_format.py', 'lib')
)

$pyinstallerArgs = @(
    '-m', 'PyInstaller',
    '--noconfirm',
    '--clean',
    '--name', 'yibinthesis',
    '--distpath', (Join-Path $Root $OutputDirectory),
    '--workpath', (Join-Path $workDirectory 'build'),
    '--specpath', $workDirectory,
    (Join-Path $Root 'lib\yibinthesis_cli\__main__.py')
)
if ($OneFile) {
    $pyinstallerArgs += '--onefile'
}
else {
    $pyinstallerArgs += '--onedir'
}
foreach ($entry in $dataFiles) {
    $source = Join-Path $Root $entry[0]
    if (-not (Test-Path -LiteralPath $source)) {
        throw "PyInstaller resource not found: $source"
    }
    $pyinstallerArgs += @('--add-data', "$source;$($entry[1])")
}

Write-Host "Building yibinthesis.exe with $Python ..."
& $Python $pyinstallerArgs
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE."
}
$executable = if ($OneFile) {
    Join-Path (Join-Path $Root $OutputDirectory) 'yibinthesis.exe'
}
else {
    Join-Path (Join-Path (Join-Path $Root $OutputDirectory) 'yibinthesis') 'yibinthesis.exe'
}
Write-Host ("Executable generated: " + $executable)
