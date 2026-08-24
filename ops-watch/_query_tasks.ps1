# _query_tasks.ps1 - read-only helper for ops_watch.py
# Reads target task names from config.json and outputs scheduler info as JSON (UTF-8).
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Continue'

$cfgPath = Join-Path $PSScriptRoot 'config.json'
$cfg = Get-Content -Raw -Path $cfgPath -Encoding UTF8 | ConvertFrom-Json

$out = @()
foreach ($t in $cfg.targets) {
    $task = $null
    try { $task = Get-ScheduledTask -TaskName $t.task -ErrorAction Stop } catch {}
    if ($null -eq $task) {
        $out += [pscustomobject]@{ name = $t.task; found = $false }
        continue
    }
    $i = $task | Get-ScheduledTaskInfo
    $lastRun = ''
    if ($i.LastRunTime -and $i.LastRunTime -ge (Get-Date '2000-01-01')) {
        $lastRun = $i.LastRunTime.ToString('yyyy-MM-dd HH:mm:ss')
    }
    $nextRun = ''
    if ($i.NextRunTime) { $nextRun = $i.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss') }
    $out += [pscustomobject]@{
        name        = $t.task
        found       = $true
        state       = [string]$task.State
        last_result = [int64]$i.LastTaskResult
        last_run    = $lastRun
        next_run    = $nextRun
    }
}
ConvertTo-Json -InputObject $out -Depth 3
