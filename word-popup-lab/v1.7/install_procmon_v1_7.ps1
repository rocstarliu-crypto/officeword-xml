# Official Microsoft diagnostics, disposable hosted runner only.
$ErrorActionPreference = 'Stop'
if ($env:GITHUB_ACTIONS -ne 'true' -or $env:RUNNER_ENVIRONMENT -ne 'github-hosted') { throw 'Hosted runner only' }
$work = Join-Path $env:RUNNER_TEMP 'word-lab-procmon'
New-Item -ItemType Directory -Path $work -Force | Out-Null
$zip = Join-Path $work 'ProcessMonitor.zip'
Invoke-WebRequest 'https://download.sysinternals.com/files/ProcessMonitor.zip' -OutFile $zip -TimeoutSec 120
Expand-Archive $zip -DestinationPath $work -Force
$exe = Join-Path $work 'Procmon64.exe'
$s = Get-AuthenticodeSignature $exe
if ($s.Status -ne 'Valid' -or $s.SignerCertificate.Subject -notmatch 'Microsoft Corporation') { throw 'Invalid Microsoft signature' }
'WORDLAB_PROCMON=' + $exe | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
Write-Output ('Official Process Monitor signature: ' + $s.Status)
