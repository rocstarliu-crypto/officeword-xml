# Disposable GitHub-hosted Windows runner only. No activation credentials.
[CmdletBinding()]
param([string]$OutputDirectory = (Join-Path $PWD 'lab-results'), [ValidateSet('Current','16.0.20228.20276')][string]$OfficeVersion = 'Current')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($env:GITHUB_ACTIONS -ne 'true' -or $env:RUNNER_ENVIRONMENT -ne 'github-hosted') {
    throw 'Installation is limited to disposable GitHub-hosted runners.'
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$work = Join-Path $env:RUNNER_TEMP 'word-popup-lab-odt-v2-0'
New-Item -ItemType Directory -Path $work -Force | Out-Null
$uri = 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_20430-20146.exe'
$installer = Join-Path $work 'odt.exe'
Invoke-WebRequest -Uri $uri -OutFile $installer -UseBasicParsing -TimeoutSec 120
$signature = Get-AuthenticodeSignature -FilePath $installer
if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'Microsoft Corporation') {
    throw 'The downloaded Office Deployment Tool does not have a valid Microsoft signature.'
}
$extract = Start-Process -FilePath $installer -ArgumentList @('/quiet', "/extract:$work") -Wait -PassThru
if ($extract.ExitCode -ne 0) { throw "ODT extraction failed: $($extract.ExitCode)" }
$xml = @'
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365HomePremRetail">
      <Language ID="zh-cn" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Excel" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="PowerPoint" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="Bing" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="PinIconsToTaskbar" Value="FALSE" />
  <Updates Enabled="FALSE" />
</Configuration>
'@
if ($OfficeVersion -ne 'Current') { $xml = $xml.Replace('Channel="Current"', ('Channel="Current" Version="' + $OfficeVersion + '"')) }
$config = Join-Path $work 'configuration.xml'
$xml | Set-Content -LiteralPath $config -Encoding UTF8
Write-Output ('Requested Office version: ' + $OfficeVersion)
$process = Start-Process -FilePath (Join-Path $work 'setup.exe') -ArgumentList @('/configure', $config) -PassThru
if (-not $process.WaitForExit(720000)) {
    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    throw 'Official Word installation exceeded the 12-minute time limit.'
}
$process.Refresh()
if ($process.ExitCode -ne 0) { throw "Official Word installation failed: $($process.ExitCode)" }

# Microsoft-documented policy values: 2 disables each connected experience class.
# Only the disposable runner profile is modified.
$privacy = 'HKCU:\Software\Policies\Microsoft\office\16.0\common\privacy'
New-Item -Path $privacy -Force | Out-Null
foreach ($name in @('disconnectedstate', 'usercontentdisabled', 'downloadcontentdisabled', 'controllerconnectedservicesenabled')) {
    New-ItemProperty -Path $privacy -Name $name -Value 2 -PropertyType DWord -Force | Out-Null
}
$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$inventory = [ordered]@{
    experimentVersion='2.0'; source='Official Microsoft Office Deployment Tool';
    requestedProduct='O365HomePremRetail'; installedProduct=$c2r.ProductReleaseIds;
    version=$c2r.VersionToReport; platform=$c2r.Platform;
    authentication='No Microsoft account supplied; activation and matching account are NOT verified';
    connectedExperiencePolicy='All four policy values set to 2 in disposable runner only'
}
$inventory | ConvertTo-Json | Tee-Object -FilePath (Join-Path $OutputDirectory 'installation.json')
