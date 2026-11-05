# Budget guard V1.0. A run is created only after the controller's preflight.
[CmdletBinding()]
param([string]$OutputDirectory = './lab-results')
$ErrorActionPreference = 'Stop'
if (-not $env:GITHUB_ACTIONS -or -not $env:GH_TOKEN) { throw 'Missing audit context; fail closed' }
$headers = @{ Authorization = ('Bearer ' + $env:GH_TOKEN); Accept = 'application/vnd.github+json'; 'X-GitHub-Api-Version' = '2022-11-28' }
$now = [DateTimeOffset]::UtcNow
$cutoff = $now.AddMinutes(-62) # Extra margin around the user's 60-minute window.
$base = 'https://api.github.com/repos/' + $env:GITHUB_REPOSITORY
$runs = Invoke-RestMethod -Uri ($base + '/actions/runs?per_page=100') -Headers $headers
$recent = @($runs.workflow_runs | Where-Object { [DateTimeOffset]$_.created_at -ge $cutoff })
if ($recent.Count -gt 10) { throw 'Hourly workflow budget exceeded. Do not execute experiments.' }
$jobCount = 0
$items = @()
foreach ($run in $recent) {
    $jobs = Invoke-RestMethod -Uri ($base + '/actions/runs/' + $run.id + '/jobs?per_page=100') -Headers $headers
    $count = @($jobs.jobs).Count
    $jobCount += $count
    if ($run.run_attempt -gt 1) { throw 'Rerun requires separate audit; fail closed' }
    $items += [ordered]@{ id=$run.id; created_at=$run.created_at; jobs=$count; attempt=$run.run_attempt }
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$audit = [ordered]@{ checked_at=$now.ToString('o'); window_minutes=62; maximum=10; workflow_runs=$recent.Count; jobs=$jobCount; runs=$items; allowed=($jobCount -le 10) }
$audit | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $OutputDirectory 'actions-budget-audit.json') -Encoding UTF8
Write-Output ('Rolling budget: ' + $recent.Count + ' workflow runs, ' + $jobCount + ' jobs, maximum 10')
if ($jobCount -gt 10) { throw 'Hourly job budget exceeded. Do not execute experiments.' }
