#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('pdf', 'word', 'all', 'check', 'doctor', 'clean')]
    [string]$Command = 'all',

    [Parameter()]
    [string]$Main = 'main.tex'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ProjectRoot = [System.IO.Path]::GetFullPath($PSScriptRoot)
$BuildRoot = Join-Path $ProjectRoot 'build'
$PdfBuildRoot = Join-Path $BuildRoot 'pdf'
$WordBuildRoot = Join-Path $BuildRoot 'word'

function Get-FullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$BasePath
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

function Test-PathWithin {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Candidate,

        [Parameter(Mandatory = $true)]
        [string]$Parent,

        [switch]$AllowEqual
    )

    $candidatePath = [System.IO.Path]::GetFullPath($Candidate).TrimEnd('\', '/')
    $parentPath = [System.IO.Path]::GetFullPath($Parent).TrimEnd('\', '/')
    if ($candidatePath.Equals($parentPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $AllowEqual.IsPresent
    }

    $prefix = $parentPath + [System.IO.Path]::DirectorySeparatorChar
    return $candidatePath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Assert-SafeBuildPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot,

        [switch]$AllowRoot
    )

    if (-not (Test-PathWithin -Candidate $Path -Parent $AllowedRoot -AllowEqual:$AllowRoot)) {
        throw "Refusing unsafe build path outside '$AllowedRoot': $Path"
    }
}

function ConvertTo-SafeName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $safe = $Value -replace '[^\p{L}\p{Nd}._-]+', '-'
    $safe = $safe.Trim(' ', '.', '-', '_')
    if ([string]::IsNullOrWhiteSpace($safe) -or $safe -eq '.' -or $safe -eq '..') {
        $safe = 'document'
    }
    if ($safe.Length -gt 80) {
        $safe = $safe.Substring(0, 80).TrimEnd('.', '-', '_')
    }
    return $safe
}

function Get-ShortHash {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
        $hash = $sha256.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash).Replace('-', '').Substring(0, 8).ToLowerInvariant())
    }
    finally {
        $sha256.Dispose()
    }
}

function Resolve-MainFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $path = Get-FullPath -Path $Value -BasePath $ProjectRoot
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "LaTeX entry point not found: $path"
    }
    if (-not [System.IO.Path]::GetExtension($path).Equals('.tex', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "LaTeX entry point must be a .tex file: $path"
    }
    return (Resolve-Path -LiteralPath $path).Path
}

function Get-BuildLayout {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MainPath
    )

    $mainFullPath = [System.IO.Path]::GetFullPath($MainPath)
    $mainDirectory = Split-Path -Parent $mainFullPath
    $jobName = [System.IO.Path]::GetFileNameWithoutExtension($mainFullPath)
    $safeJobName = ConvertTo-SafeName -Value $jobName
    $insideProject = Test-PathWithin -Candidate $mainFullPath -Parent $ProjectRoot

    if ($insideProject) {
        $relativeMain = $mainFullPath.Substring($ProjectRoot.TrimEnd('\', '/').Length).TrimStart('\', '/')
        $relativeDirectory = Split-Path -Parent $relativeMain
        if ([string]::IsNullOrWhiteSpace($relativeDirectory) -or $relativeDirectory -eq '.') {
            $pdfOutputDirectory = $PdfBuildRoot
        }
        else {
            $pdfOutputDirectory = Join-Path $PdfBuildRoot $relativeDirectory
        }

        if ($jobName -eq 'main' -and -not [string]::IsNullOrWhiteSpace($relativeDirectory)) {
            $wordBaseName = ConvertTo-SafeName -Value (Split-Path -Leaf $relativeDirectory)
        }
        else {
            $wordBaseName = $safeJobName
        }
        $compilerInput = $relativeMain
    }
    else {
        $parentName = ConvertTo-SafeName -Value (Split-Path -Leaf $mainDirectory)
        $hash = Get-ShortHash -Value $mainFullPath.ToLowerInvariant()
        $externalName = ConvertTo-SafeName -Value "$parentName-$safeJobName-$hash"
        $pdfOutputDirectory = Join-Path (Join-Path $PdfBuildRoot 'external') $externalName
        $wordBaseName = $externalName
        $compilerInput = $mainFullPath
    }

    $pdfOutputDirectory = [System.IO.Path]::GetFullPath($pdfOutputDirectory)
    $wordOutput = [System.IO.Path]::GetFullPath((Join-Path $WordBuildRoot "$wordBaseName.docx"))
    Assert-SafeBuildPath -Path $pdfOutputDirectory -AllowedRoot $PdfBuildRoot -AllowRoot
    Assert-SafeBuildPath -Path $wordOutput -AllowedRoot $WordBuildRoot

    return [pscustomobject]@{
        MainPath = $mainFullPath
        CompilerInput = $compilerInput
        JobName = $jobName
        PdfOutputDirectory = $pdfOutputDirectory
        PdfOutput = Join-Path $pdfOutputDirectory "$jobName.pdf"
        WordBaseName = $wordBaseName
        WordOutput = $wordOutput
        IsProjectFile = $insideProject
    }
}

function Get-KnownTectonicPaths {
    $paths = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        $latexCache = Join-Path $env:USERPROFILE '.codex\plugins\cache\openai-bundled\latex'
        if (Test-Path -LiteralPath $latexCache -PathType Container) {
            Get-ChildItem -LiteralPath $latexCache -Directory -ErrorAction SilentlyContinue |
                Sort-Object Name -Descending |
                ForEach-Object {
                    $paths.Add((Join-Path $_.FullName 'bin\tectonic.exe'))
                }
        }
    }
    return $paths.ToArray()
}

function Find-Executable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [string]$EnvironmentVariable,

        [string[]]$CommandNames = @(),
        [string[]]$KnownPaths = @()
    )

    $override = [System.Environment]::GetEnvironmentVariable($EnvironmentVariable)
    if (-not [string]::IsNullOrWhiteSpace($override)) {
        $override = $override.Trim().Trim('"')
        if (Test-Path -LiteralPath $override -PathType Leaf) {
            return [pscustomobject]@{
                Name = $DisplayName
                Path = (Resolve-Path -LiteralPath $override).Path
                Source = "env:$EnvironmentVariable"
            }
        }
        $overrideCommand = Get-Command $override -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $overrideCommand) {
            return [pscustomobject]@{
                Name = $DisplayName
                Path = $overrideCommand.Source
                Source = "env:$EnvironmentVariable"
            }
        }
        return [pscustomobject]@{
            Name = $DisplayName
            Path = $null
            Source = "invalid env:$EnvironmentVariable ($override)"
        }
    }

    foreach ($knownPath in $KnownPaths) {
        if (-not [string]::IsNullOrWhiteSpace($knownPath) -and (Test-Path -LiteralPath $knownPath -PathType Leaf)) {
            return [pscustomobject]@{
                Name = $DisplayName
                Path = (Resolve-Path -LiteralPath $knownPath).Path
                Source = 'known local path'
            }
        }
    }

    foreach ($commandName in $CommandNames) {
        $command = Get-Command $commandName -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command) {
            return [pscustomobject]@{
                Name = $DisplayName
                Path = $command.Source
                Source = 'PATH'
            }
        }
    }

    return [pscustomobject]@{
        Name = $DisplayName
        Path = $null
        Source = 'not found'
    }
}

function Find-Python {
    $python = Find-Executable -DisplayName 'Python' -EnvironmentVariable 'YIBINTHESIS_PYTHON' -CommandNames @('python', 'python3', 'py') -KnownPaths @(
        (Join-Path $ProjectRoot '.tools\venv\Scripts\python.exe'),
        (Join-Path $ProjectRoot '.tools\python\python.exe')
    )
    if ([string]::IsNullOrWhiteSpace([string]$python.Path)) {
        return [pscustomobject]@{
            Name = 'Python'
            Path = $null
            Source = $python.Source
            PrefixArguments = @()
        }
    }

    $prefix = @()
    if ([System.IO.Path]::GetFileName($python.Path) -ieq 'py.exe') {
        $prefix = @('-3')
    }
    return [pscustomobject]@{
        Name = 'Python'
        Path = $python.Path
        Source = $python.Source
        PrefixArguments = $prefix
    }
}

function Get-ToolSet {
    $localToolRoot = Join-Path $ProjectRoot '.tools'
    $tectonicPaths = @(
        (Join-Path $localToolRoot 'tectonic\tectonic.exe'),
        (Join-Path $localToolRoot 'tectonic.exe')
    ) + @(Get-KnownTectonicPaths)
    $tectonic = Find-Executable -DisplayName 'Tectonic' -EnvironmentVariable 'YIBINTHESIS_TECTONIC' -CommandNames @('tectonic') -KnownPaths $tectonicPaths
    $biber = Find-Executable -DisplayName 'Biber' -EnvironmentVariable 'YIBINTHESIS_BIBER' -CommandNames @('biber') -KnownPaths @(
        (Join-Path $localToolRoot 'biber-2.17\biber.exe'),
        (Join-Path $localToolRoot 'biber\biber.exe')
    )
    $pandoc = Find-Executable -DisplayName 'Pandoc' -EnvironmentVariable 'YIBINTHESIS_PANDOC' -CommandNames @('pandoc') -KnownPaths @(
        (Join-Path $localToolRoot 'pandoc-3.9.0.2\pandoc.exe'),
        (Join-Path $localToolRoot 'pandoc\pandoc.exe')
    )

    return [pscustomobject]@{
        Tectonic = $tectonic
        Biber = $biber
        Pandoc = $pandoc
        Latexmk = Find-Executable -DisplayName 'latexmk' -EnvironmentVariable 'YIBINTHESIS_LATEXMK' -CommandNames @('latexmk')
        XeLaTeX = Find-Executable -DisplayName 'XeLaTeX' -EnvironmentVariable 'YIBINTHESIS_XELATEX' -CommandNames @('xelatex')
        Python = Find-Python
    }
}

function Invoke-NativeTool {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$Arguments = @(),

        [Parameter(Mandatory = $true)]
        [string]$WorkingDirectory,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    Write-Host "[$Description] $FilePath $($Arguments -join ' ')"
    Push-Location -LiteralPath $WorkingDirectory
    try {
        & $FilePath @Arguments
        $exitCode = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }
    if ($exitCode -ne 0) {
        throw "$Description failed with exit code $exitCode."
    }
}

function Test-MainNeedsBiber {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MainPath
    )

    $source = Get-Content -LiteralPath $MainPath -Raw -Encoding UTF8
    if ($source -match '\\documentclass\s*\[[^\]]*\bnobibliography\b[^\]]*\]') {
        return $false
    }
    return ($source -match '\\addbibresource\b' -or $source -match '\\print(?:yibin)?bibliography\b')
}

function Get-BiberVersion {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Biber
    )

    $line = Get-VersionLine -Tool $Biber -Match 'biber version'
    if ($line -match 'biber version:\s*(?<version>\d+\.\d+(?:\.\d+)?)') {
        return $Matches['version']
    }
    return $null
}

function Assert-CompatibleBiber {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Biber
    )

    if ([string]::IsNullOrWhiteSpace([string]$Biber.Path)) {
        throw 'Biber is required by this entry point but was not found.'
    }
    $version = Get-BiberVersion -Biber $Biber
    if ($version -ne '2.17') {
        $shown = if ([string]::IsNullOrWhiteSpace($version)) { 'unknown' } else { $version }
        throw "Biber 2.17 is required for the current BCF 3.8 workflow; detected $shown. Biber 2.21 is incompatible with this project."
    }
}

function Invoke-TectonicBuild {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Layout,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Tools
    )

    if ([string]::IsNullOrWhiteSpace([string]$Tools.Tectonic.Path)) {
        throw 'Tectonic was not found. Run .\build.ps1 doctor for dependency details.'
    }

    $needsBiber = Test-MainNeedsBiber -MainPath $Layout.MainPath
    if ($needsBiber) {
        Assert-CompatibleBiber -Biber $Tools.Biber
    }
    $tectonicVersion = Get-VersionLine -Tool $Tools.Tectonic -Match '^Tectonic '
    if ($tectonicVersion -notmatch '^Tectonic 0\.16\.9(?:\s|$)') {
        Write-Warning "This template is verified with Tectonic 0.16.9; detected: $tectonicVersion"
    }

    New-Item -ItemType Directory -Path $Layout.PdfOutputDirectory -Force | Out-Null
    if (Test-Path -LiteralPath $Layout.PdfOutput -PathType Leaf) {
        Remove-Item -LiteralPath $Layout.PdfOutput -Force
    }

    $commonArguments = @(
        '-X', 'compile',
        '-Z', "search-path=$ProjectRoot",
        '-Z', "search-path=$(Split-Path -Parent $Layout.MainPath)",
        '--outdir', $Layout.PdfOutputDirectory,
        '--keep-intermediates',
        '--keep-logs',
        '--synctex'
    )
    Invoke-NativeTool -FilePath $Tools.Tectonic.Path -Arguments ($commonArguments + @('--pass', 'tex', $Layout.CompilerInput)) -WorkingDirectory $ProjectRoot -Description 'tectonic pass 1'

    if ($needsBiber) {
        if ($Layout.JobName -match '[^\x00-\x7F]') {
            throw 'Biber 2.17 cannot reliably process a non-ASCII TeX job name. Rename the entry .tex file with an ASCII filename.'
        }
        $relativeOutputDirectory = $Layout.PdfOutputDirectory.Substring(
            $ProjectRoot.TrimEnd('\', '/').Length
        ).TrimStart('\', '/')
        if ($relativeOutputDirectory -match '[^\x00-\x7F]') {
            throw 'Biber 2.17 cannot reliably process a non-ASCII output path. Use an ASCII relative directory for the entry point.'
        }
        $biberArguments = @(
            '--input-directory', $relativeOutputDirectory,
            '--output-directory', $relativeOutputDirectory,
            $Layout.JobName
        )
        Invoke-NativeTool -FilePath $Tools.Biber.Path -Arguments $biberArguments -WorkingDirectory $ProjectRoot -Description 'biber 2.17'

        $bbl = Join-Path $Layout.PdfOutputDirectory "$($Layout.JobName).bbl"
        if (-not (Test-Path -LiteralPath $bbl -PathType Leaf) -or (Get-Item -LiteralPath $bbl).Length -eq 0) {
            throw "Biber exited successfully but produced no usable bibliography: $bbl"
        }
    }

    # The final default Tectonic pass performs PDF conversion and may invoke
    # Biber once more. Keep the verified 2.17 executable available on PATH.
    $oldPath = $env:PATH
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$Tools.Biber.Path)) {
            $biberDirectory = Split-Path -Parent $Tools.Biber.Path
            $env:PATH = $biberDirectory + [System.IO.Path]::PathSeparator + $env:PATH
        }
        $finalArguments = $commonArguments + @('--reruns', '2', $Layout.CompilerInput)
        Invoke-NativeTool -FilePath $Tools.Tectonic.Path -Arguments $finalArguments -WorkingDirectory $ProjectRoot -Description 'tectonic final passes'
    }
    finally {
        $env:PATH = $oldPath
    }
}

function Invoke-LatexmkBuild {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Layout,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Tools
    )

    New-Item -ItemType Directory -Path $Layout.PdfOutputDirectory -Force | Out-Null
    $oldPath = $env:PATH
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$Tools.Biber.Path)) {
            $biberDirectory = Split-Path -Parent $Tools.Biber.Path
            $env:PATH = $biberDirectory + [System.IO.Path]::PathSeparator + $env:PATH
        }
        $arguments = @(
            '-xelatex',
            '-interaction=nonstopmode',
            '-halt-on-error',
            '-file-line-error',
            '-synctex=1',
            "-outdir=$($Layout.PdfOutputDirectory)",
            $Layout.CompilerInput
        )
        Invoke-NativeTool -FilePath $Tools.Latexmk.Path -Arguments $arguments -WorkingDirectory $ProjectRoot -Description 'latexmk'
    }
    finally {
        $env:PATH = $oldPath
    }
}

function Build-Pdf {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Layout,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Tools
    )

    $canUseLatexmk = -not [string]::IsNullOrWhiteSpace([string]$Tools.Latexmk.Path) -and
        -not [string]::IsNullOrWhiteSpace([string]$Tools.XeLaTeX.Path)
    if ((Test-MainNeedsBiber -MainPath $Layout.MainPath) -and -not $canUseLatexmk) {
        Assert-CompatibleBiber -Biber $Tools.Biber
    }
    if ($canUseLatexmk) {
        Invoke-LatexmkBuild -Layout $Layout -Tools $Tools
    }
    else {
        Invoke-TectonicBuild -Layout $Layout -Tools $Tools
    }

    if (-not (Test-Path -LiteralPath $Layout.PdfOutput -PathType Leaf)) {
        throw "PDF compiler exited successfully but output was not found: $($Layout.PdfOutput)"
    }
    Write-Host "PDF generated: $($Layout.PdfOutput)"
}

function Build-Word {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Layout,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Tools
    )

    if ([string]::IsNullOrWhiteSpace([string]$Tools.Python.Path)) {
        throw 'Python was not found. Set YIBINTHESIS_PYTHON or add Python to PATH.'
    }
    if ([string]::IsNullOrWhiteSpace([string]$Tools.Pandoc.Path)) {
        throw 'Pandoc was not found. Set YIBINTHESIS_PANDOC or add Pandoc to PATH.'
    }

    $builder = Join-Path $ProjectRoot 'tools\build_word.py'
    if (-not (Test-Path -LiteralPath $builder -PathType Leaf)) {
        throw "Word builder not found: $builder"
    }
    New-Item -ItemType Directory -Path $WordBuildRoot -Force | Out-Null
    $arguments = @()
    $arguments += $Tools.Python.PrefixArguments
    $arguments += @(
        $builder,
        '--main', $Layout.MainPath,
        '--output', $Layout.WordOutput,
        '--pandoc', $Tools.Pandoc.Path
    )
    $oldPythonUtf8 = $env:PYTHONUTF8
    try {
        $env:PYTHONUTF8 = '1'
        Invoke-NativeTool -FilePath $Tools.Python.Path -Arguments $arguments -WorkingDirectory $ProjectRoot -Description 'Word build'
    }
    finally {
        $env:PYTHONUTF8 = $oldPythonUtf8
    }
    if (-not (Test-Path -LiteralPath $Layout.WordOutput -PathType Leaf)) {
        throw "Word builder exited successfully but output was not found: $($Layout.WordOutput)"
    }
    Write-Host "Word generated: $($Layout.WordOutput)"
}

function Invoke-Checks {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Tools
    )

    if ([string]::IsNullOrWhiteSpace([string]$Tools.Python.Path)) {
        throw 'Python was not found. Set YIBINTHESIS_PYTHON or add Python to PATH.'
    }

    $auditor = Join-Path $ProjectRoot 'tests\audit_format.py'
    $reference = Join-Path $ProjectRoot 'word\reference.docx'
    if (-not (Test-Path -LiteralPath $auditor -PathType Leaf)) {
        throw "Format auditor not found: $auditor"
    }

    $auditArguments = @()
    $auditArguments += $Tools.Python.PrefixArguments
    $auditArguments += @($auditor, '--reference', $reference)

    $oldPythonUtf8 = $env:PYTHONUTF8
    try {
        $env:PYTHONUTF8 = '1'
        Invoke-NativeTool -FilePath $Tools.Python.Path -Arguments $auditArguments -WorkingDirectory $ProjectRoot -Description 'format audit'
    }
    finally {
        $env:PYTHONUTF8 = $oldPythonUtf8
    }

    foreach ($smokeEntry in @('tests\smoke.tex', 'tests\smoke-science.tex')) {
        $smokePath = Resolve-MainFile -Value $smokeEntry
        $smokeLayout = Get-BuildLayout -MainPath $smokePath
        Build-Pdf -Layout $smokeLayout -Tools $Tools
    }
    Write-Host 'Check result: PASS (format contract + humanities/science smoke builds)'
}

function Get-VersionLine {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Tool,

        [string[]]$Arguments = @('--version'),
        [string]$Match = ''
    )

    if ([string]::IsNullOrWhiteSpace([string]$Tool.Path)) {
        return $null
    }
    $oldErrorActionPreference = $ErrorActionPreference
    try {
        # Windows PowerShell can promote redirected native stderr records when
        # the script-wide preference is Stop (Biber emits harmless warnings).
        $ErrorActionPreference = 'Continue'
        $commandArguments = @()
        if ($Tool.PSObject.Properties.Name -contains 'PrefixArguments') {
            $commandArguments += $Tool.PrefixArguments
        }
        $commandArguments += $Arguments
        $lines = & $Tool.Path @commandArguments 2>&1 | ForEach-Object { $_.ToString().Trim() } | Where-Object { $_ }
        if ($LASTEXITCODE -ne 0) {
            return 'version check failed'
        }
        if (-not [string]::IsNullOrWhiteSpace($Match)) {
            $matched = $lines | Where-Object { $_ -match $Match } | Select-Object -First 1
            if ($null -ne $matched) {
                return $matched
            }
        }
        return ($lines | Select-Object -First 1)
    }
    catch {
        return "version check failed: $($_.Exception.Message)"
    }
    finally {
        $ErrorActionPreference = $oldErrorActionPreference
    }
}

function Show-ToolStatus {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Tool,

        [string]$Version,
        [bool]$Required = $true
    )

    $role = if ($Required) { 'required' } else { 'optional' }
    if ([string]::IsNullOrWhiteSpace([string]$Tool.Path)) {
        Write-Host ("  [MISSING] {0} ({1}) - {2}" -f $Tool.Name, $role, $Tool.Source)
        return $false
    }
    $versionText = if ([string]::IsNullOrWhiteSpace($Version)) { '' } else { " | $Version" }
    Write-Host ("  [OK]      {0} ({1}) - {2} [{3}]{4}" -f $Tool.Name, $role, $Tool.Path, $Tool.Source, $versionText)
    return $true
}

function Test-PythonDocx {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Python
    )

    if ([string]::IsNullOrWhiteSpace([string]$Python.Path)) {
        return $false
    }
    $arguments = @()
    $arguments += $Python.PrefixArguments
    $arguments += @('-c', 'import docx')
    try {
        & $Python.Path @arguments 2>$null | Out-Null
        return ($LASTEXITCODE -eq 0)
    }
    catch {
        return $false
    }
}

function Invoke-Doctor {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Layout,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Tools
    )

    Write-Host "YibinThesis dependency doctor"
    Write-Host "  Project root: $ProjectRoot"
    Write-Host "  Main file:    $($Layout.MainPath)"
    Write-Host "  PDF output:   $($Layout.PdfOutput)"
    Write-Host "  Word output:  $($Layout.WordOutput)"
    Write-Host ''

    $needsBiber = Test-MainNeedsBiber -MainPath $Layout.MainPath
    $latexmkOk = Show-ToolStatus -Tool $Tools.Latexmk -Version (Get-VersionLine -Tool $Tools.Latexmk -Match 'Latexmk') -Required:$false
    $xelatexOk = Show-ToolStatus -Tool $Tools.XeLaTeX -Version (Get-VersionLine -Tool $Tools.XeLaTeX -Match 'XeTeX|XeLaTeX') -Required:$false
    $tectonicOk = Show-ToolStatus -Tool $Tools.Tectonic -Version (Get-VersionLine -Tool $Tools.Tectonic -Match '^Tectonic') -Required:$false
    $biberVersionLine = Get-VersionLine -Tool $Tools.Biber -Match 'biber version'
    $biberOk = Show-ToolStatus -Tool $Tools.Biber -Version $biberVersionLine -Required:$needsBiber
    $biberVersion = if ($biberOk) { Get-BiberVersion -Biber $Tools.Biber } else { $null }
    $biberCompatible = (-not $needsBiber) -or ($biberVersion -eq '2.17')
    if ($needsBiber -and $biberOk -and -not $biberCompatible) {
        Write-Host "  [INCOMPATIBLE] Biber $biberVersion; Tectonic 0.16.9 requires Biber 2.17 for BCF 3.8."
    }
    $pandocOk = Show-ToolStatus -Tool $Tools.Pandoc -Version (Get-VersionLine -Tool $Tools.Pandoc -Match '^pandoc ') -Required:$true
    $pythonOk = Show-ToolStatus -Tool $Tools.Python -Version (Get-VersionLine -Tool $Tools.Python -Match '^Python ') -Required:$true

    $pythonDocxOk = Test-PythonDocx -Python $Tools.Python
    if ($pythonDocxOk) {
        Write-Host '  [OK]      Python module: python-docx'
    }
    else {
        Write-Host '  [MISSING] Python module: python-docx (install requirements-word.txt)'
    }

    $wordBuilder = Join-Path $ProjectRoot 'tools\build_word.py'
    $referenceDoc = Join-Path $ProjectRoot 'word\reference.docx'
    $wordFilesOk = (Test-Path -LiteralPath $wordBuilder -PathType Leaf) -and (Test-Path -LiteralPath $referenceDoc -PathType Leaf)
    if ($wordFilesOk) {
        Write-Host '  [OK]      Word builder and reference.docx'
    }
    else {
        Write-Host '  [MISSING] tools/build_word.py or word/reference.docx'
    }

    $pdfEngineOk = ($latexmkOk -and $xelatexOk) -or $tectonicOk
    $pdfReady = $pdfEngineOk -and ((-not $needsBiber) -or ($biberOk -and $biberCompatible))
    $wordReady = $pandocOk -and $pythonOk -and $pythonDocxOk -and $wordFilesOk

    Write-Host ''
    Write-Host ("  PDF toolchain:  {0}" -f $(if ($pdfReady) { 'READY' } else { 'NOT READY' }))
    Write-Host ("  Word toolchain: {0}" -f $(if ($wordReady) { 'READY' } else { 'NOT READY' }))
    Write-Host '  Overrides: YIBINTHESIS_TECTONIC, YIBINTHESIS_BIBER, YIBINTHESIS_PANDOC,'
    Write-Host '             YIBINTHESIS_PYTHON, YIBINTHESIS_LATEXMK, YIBINTHESIS_XELATEX'

    if ($pdfReady -and $wordReady) {
        Write-Host 'Doctor result: READY (exit code 0)'
        return 0
    }
    Write-Host 'Doctor result: NOT READY (exit code 2)'
    return 2
}

function Remove-EmptyPdfDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StartDirectory
    )

    $current = [System.IO.Path]::GetFullPath($StartDirectory)
    while ((Test-PathWithin -Candidate $current -Parent $PdfBuildRoot) -and (Test-Path -LiteralPath $current -PathType Container)) {
        $entries = @(Get-ChildItem -LiteralPath $current -Force -ErrorAction SilentlyContinue)
        if ($entries.Count -ne 0) {
            break
        }
        Assert-SafeBuildPath -Path $current -AllowedRoot $PdfBuildRoot
        Remove-Item -LiteralPath $current -Force
        $current = Split-Path -Parent $current
    }
}

function Clean-Build {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Layout
    )

    Assert-SafeBuildPath -Path $Layout.PdfOutputDirectory -AllowedRoot $PdfBuildRoot -AllowRoot
    Assert-SafeBuildPath -Path $Layout.WordOutput -AllowedRoot $WordBuildRoot
    $removed = New-Object System.Collections.Generic.List[string]

    if (Test-Path -LiteralPath $Layout.PdfOutputDirectory -PathType Container) {
        $jobPrefix = $Layout.JobName + '.'
        $jobDashPrefix = $Layout.JobName + '-'
        $mintedName = '_minted-' + $Layout.JobName
        Get-ChildItem -LiteralPath $Layout.PdfOutputDirectory -Force -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name.StartsWith($jobPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
                $_.Name.StartsWith($jobDashPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
                $_.Name.Equals($mintedName, [System.StringComparison]::OrdinalIgnoreCase)
            } |
            ForEach-Object {
                Assert-SafeBuildPath -Path $_.FullName -AllowedRoot $PdfBuildRoot
                if ($_.PSIsContainer) {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force
                }
                else {
                    Remove-Item -LiteralPath $_.FullName -Force
                }
                $removed.Add($_.FullName)
            }

        # Tectonic may retain a bibliography copy beside the root PDF. Remove
        # only resource basenames explicitly declared by the selected entry.
        $source = Get-Content -LiteralPath $Layout.MainPath -Raw -Encoding UTF8
        $resourceMatches = [regex]::Matches(
            $source,
            '\\addbibresource(?:\s*\[[^\]]*\])?\s*\{(?<path>[^}]+)\}'
        )
        foreach ($match in $resourceMatches) {
            $resourceName = [System.IO.Path]::GetFileName($match.Groups['path'].Value.Trim())
            if ([string]::IsNullOrWhiteSpace($resourceName)) {
                continue
            }
            $resourceCopy = Join-Path $Layout.PdfOutputDirectory $resourceName
            Assert-SafeBuildPath -Path $resourceCopy -AllowedRoot $PdfBuildRoot
            if (Test-Path -LiteralPath $resourceCopy -PathType Leaf) {
                Remove-Item -LiteralPath $resourceCopy -Force
                $removed.Add($resourceCopy)
            }
        }

        if (-not $Layout.PdfOutputDirectory.Equals($PdfBuildRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
            Remove-EmptyPdfDirectories -StartDirectory $Layout.PdfOutputDirectory
        }
    }

    $wordStem = [System.IO.Path]::Combine(
        (Split-Path -Parent $Layout.WordOutput),
        [System.IO.Path]::GetFileNameWithoutExtension($Layout.WordOutput)
    )
    foreach ($candidate in @(
        $Layout.WordOutput,
        "$wordStem.pandoc.md",
        "$wordStem.sections.docx"
    )) {
        Assert-SafeBuildPath -Path $candidate -AllowedRoot $WordBuildRoot
        if (Test-Path -LiteralPath $candidate) {
            Remove-Item -LiteralPath $candidate -Force
            $removed.Add($candidate)
        }
    }

    if ($removed.Count -eq 0) {
        Write-Host 'Nothing to clean for the selected entry point.'
    }
    else {
        Write-Host "Removed $($removed.Count) build artifact(s):"
        $removed | ForEach-Object { Write-Host "  $_" }
    }
}

try {
    $mainPath = Resolve-MainFile -Value $Main
    $layout = Get-BuildLayout -MainPath $mainPath

    if ($Command -eq 'clean') {
        Clean-Build -Layout $layout
        exit 0
    }

    $tools = Get-ToolSet
    switch ($Command) {
        'doctor' {
            exit (Invoke-Doctor -Layout $layout -Tools $tools)
        }
        'pdf' {
            Build-Pdf -Layout $layout -Tools $tools
        }
        'word' {
            Build-Word -Layout $layout -Tools $tools
        }
        'check' {
            Invoke-Checks -Tools $tools
        }
        'all' {
            Build-Pdf -Layout $layout -Tools $tools
            Build-Word -Layout $layout -Tools $tools
        }
    }
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
