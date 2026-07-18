#requires -Version 5.1

[CmdletBinding()]
param([switch]$CheckOnly)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$definitions = @(
    [pscustomobject]@{ Name = '图'; Position = 1 },
    [pscustomobject]@{ Name = '表'; Position = 0 },
    [pscustomobject]@{ Name = '公式'; Position = 0 }
)

$word = $null
$ownsWord = $false
$ownedProcessIds = New-Object System.Collections.Generic.List[int]
$wordProcessIdsBefore = @()
$automationStart = Get-Date
$temporaryDocument = $null
$documents = $null
$captionLabels = $null
$normalTemplate = $null
$labels = New-Object System.Collections.Generic.List[object]
$existingByName = @{}
$created = New-Object System.Collections.Generic.List[string]
$changed = New-Object System.Collections.Generic.List[string]

try {
    $wordProcessIdsBefore = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    if ($wordProcessIdsBefore.Count -gt 0) {
        throw (
            '安装 Word 题注标签前请先关闭所有 Word 窗口和后台 WINWORD.EXE，' +
            "当前进程：$($wordProcessIdsBefore -join ', ')。"
        )
    }

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $ownsWord = $true
    foreach ($process in @(Get-Process WINWORD -ErrorAction SilentlyContinue)) {
        if ($wordProcessIdsBefore -notcontains $process.Id) {
            $ownedProcessIds.Add([int]$process.Id)
        }
    }

    $documents = $word.Documents
    $captionLabels = $word.CaptionLabels
    $normalTemplate = $word.NormalTemplate

    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($definition in $definitions) {
        try {
            $label = $captionLabels.Item([string]$definition.Name)
            $labels.Add($label)
            $existingByName[[string]$definition.Name] = $label
        }
        catch {
            $missing.Add([string]$definition.Name)
        }
    }

    if ($CheckOnly) {
        if ($missing.Count -gt 0) {
            throw "Word 缺少宜宾论文题注标签：$($missing -join '、')。"
        }
        Write-Host 'Word caption labels are registered: 图, 表, 公式.'
        return
    }

    # Word refuses CaptionLabels.Add when the automation instance has no active
    # document window.  Create a disposable blank document only for the global
    # label registration; never save or alter a user's open document.
    if ($documents.Count -eq 0) {
        $temporaryDocument = $documents.Add()
        $temporaryDocument.Activate()
    }

    $needsPropertyUpdate = $false
    foreach ($definition in $definitions) {
        $name = [string]$definition.Name
        if (-not $existingByName.ContainsKey($name)) {
            continue
        }
        $label = $existingByName[$name]
        if ([int]$label.Position -ne [int]$definition.Position -or [int]$label.NumberStyle -ne 0) {
            $needsPropertyUpdate = $true
        }
    }

    if (($missing.Count -gt 0 -or $needsPropertyUpdate) -and -not $normalTemplate.Saved) {
        throw (
            'Normal.dotm 当前包含其他未保存更改。请先在 Word 中保存或放弃这些更改，' +
            '再重新运行题注标签安装脚本。'
        )
    }

    if ($missing.Count -gt 0 -or $needsPropertyUpdate) {
        $normalPath = [string]$normalTemplate.FullName
        if (-not [string]::IsNullOrWhiteSpace($normalPath) -and (Test-Path -LiteralPath $normalPath -PathType Leaf)) {
            $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
            $backup = "$normalPath.yibinthesis-caption-labels-$timestamp.bak"
            Copy-Item -LiteralPath $normalPath -Destination $backup
            Write-Host "Backed up Word global template: $backup"
        }
    }

    foreach ($definition in $definitions) {
        $name = [string]$definition.Name
        if ($existingByName.ContainsKey($name)) {
            $label = $existingByName[$name]
        }
        else {
            $label = $captionLabels.Add($name)
            $labels.Add($label)
            $existingByName[$name] = $label
            $created.Add($name)
        }

        if ([int]$label.Position -ne [int]$definition.Position) {
            $label.Position = [int]$definition.Position
            $changed.Add("$($definition.Name):position")
        }
        if ([int]$label.NumberStyle -ne 0) {
            $label.NumberStyle = 0
            $changed.Add("$($definition.Name):number-style")
        }
    }

    if ($created.Count -gt 0 -or $changed.Count -gt 0) {
        $normalTemplate.Save()
    }

    foreach ($definition in $definitions) {
        $verification = $null
        try {
            $verification = $captionLabels.Item([string]$definition.Name)
            if ([int]$verification.Position -ne [int]$definition.Position) {
                throw "题注标签 '$($definition.Name)' 的默认位置未保存。"
            }
        }
        finally {
            if ($null -ne $verification) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($verification)
            }
        }
    }

    if ($created.Count -eq 0 -and $changed.Count -eq 0) {
        Write-Host 'Word caption labels were already configured: 图, 表, 公式.'
    }
    else {
        Write-Host "Registered Word caption labels: 图, 表, 公式."
    }
    Write-Host 'Reopen Word before using Insert Caption.'
}
finally {
    foreach ($label in $labels) {
        if ($null -ne $label) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($label) } catch {}
        }
    }
    if ($null -ne $temporaryDocument) {
        try { $temporaryDocument.Close($false) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($temporaryDocument) } catch {}
    }
    foreach ($comObject in @($normalTemplate, $captionLabels, $documents)) {
        if ($null -ne $comObject) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
        }
    }
    if ($null -ne $word) {
        if ($ownsWord) {
            try { $word.Quit() } catch {}
        }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($ownsWord) {
        $cleanupIds = New-Object System.Collections.Generic.HashSet[int]
        foreach ($processId in $ownedProcessIds) {
            [void]$cleanupIds.Add([int]$processId)
        }
        foreach ($process in @(Get-Process WINWORD -ErrorAction SilentlyContinue)) {
            if ($wordProcessIdsBefore -notcontains $process.Id -and
                $process.StartTime -ge $automationStart.AddSeconds(-1)) {
                [void]$cleanupIds.Add([int]$process.Id)
            }
        }
        foreach ($processId in $cleanupIds) {
            for ($attempt = 0; $attempt -lt 20; $attempt++) {
                if ($null -eq (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
                    break
                }
                Start-Sleep -Milliseconds 100
            }
            if ($null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
                Stop-Process -Id $processId -Force
            }
        }
        # Word may spawn a replacement background process shortly after Quit.
        # Sweep only processes created by this installer, never pre-existing
        # user Word windows.
        for ($sweep = 0; $sweep -lt 30; $sweep++) {
            $lateProcesses = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object {
                $wordProcessIdsBefore -notcontains $_.Id -and
                $_.StartTime -ge $automationStart.AddSeconds(-1)
            })
            foreach ($process in $lateProcesses) {
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            }
            Start-Sleep -Milliseconds 100
        }
    }
}
