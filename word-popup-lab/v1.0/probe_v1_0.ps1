# Word read-only selection popup lab V1.0
# Flow: inventory -> desktop check -> Word check -> explicit prerequisite result.
# This probe does not install Office, change privacy settings, or open user files.
[CmdletBinding()]
param([string]$OutputDirectory = (Join-Path $PSScriptRoot 'results'))

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class WordLabDesktopV1 {
  [DllImport("user32.dll", SetLastError=true)]
  public static extern IntPtr OpenInputDesktop(uint flags, bool inherit, uint access);
  [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool GetUserObjectInformation(IntPtr handle, int index,
    StringBuilder text, int length, out int needed);
  [DllImport("user32.dll")] public static extern bool CloseDesktop(IntPtr handle);
}
'@

$desktop = [ordered]@{
    processSessionId = [Diagnostics.Process]::GetCurrentProcess().SessionId
    userInteractive = [Environment]::UserInteractive
    inputDesktopName = $null
    openInputDesktopError = 0
}
$handle = [WordLabDesktopV1]::OpenInputDesktop(0, $false, 1)
if ($handle -ne [IntPtr]::Zero) {
    try {
        $buffer = New-Object Text.StringBuilder 256
        $needed = 0
        if ([WordLabDesktopV1]::GetUserObjectInformation($handle, 2, $buffer, 512, [ref]$needed)) {
            $desktop.inputDesktopName = $buffer.ToString()
        } else { $desktop.openInputDesktopError = [Runtime.InteropServices.Marshal]::GetLastWin32Error() }
    } finally { [void][WordLabDesktopV1]::CloseDesktop($handle) }
} else { $desktop.openInputDesktopError = [Runtime.InteropServices.Marshal]::GetLastWin32Error() }

$candidates = New-Object 'System.Collections.Generic.List[string]'
foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
    if ($root) {
        foreach ($relative in @('Microsoft Office\root\Office16\WINWORD.EXE', 'Microsoft Office\Office16\WINWORD.EXE')) {
            $candidate = Join-Path $root $relative
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $candidates.Add($candidate) }
        }
    }
}
$registration = Get-ItemProperty -LiteralPath 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE' -ErrorAction SilentlyContinue
if ($null -ne $registration) {
    $registeredPath = $registration.'(default)'
    if ($registeredPath -and (Test-Path -LiteralPath $registeredPath -PathType Leaf)) { $candidates.Add($registeredPath) }
}
$wordExecutables = @($candidates | Select-Object -Unique | ForEach-Object {
    $file = Get-Item -LiteralPath $_
    [ordered]@{ path = $file.FullName; fileVersion = $file.VersionInfo.FileVersion; productVersion = $file.VersionInfo.ProductVersion }
})
$comRegistrationPresent = $null -ne [Type]::GetTypeFromProgID('Word.Application')
$blockers = New-Object 'System.Collections.Generic.List[string]'
if ($wordExecutables.Count -eq 0 -and -not $comRegistrationPresent) { $blockers.Add('WORD_NOT_INSTALLED') }
if ($desktop.processSessionId -eq 0 -or $desktop.inputDesktopName -ne 'Default') { $blockers.Add('INTERACTIVE_DESKTOP_NOT_CONFIRMED') }
$blockers.Add('MATCHING_LICENSE_AND_SIGNED_IN_PROFILE_NOT_VERIFIED')

$result = [ordered]@{
    labVersion = '1.0'
    timestampUtc = [DateTime]::UtcNow.ToString('o')
    operatingSystem = [Environment]::OSVersion.VersionString
    runnerImage = $env:ImageOS
    runnerImageVersion = $env:ImageVersion
    desktop = $desktop
    wordExecutables = $wordExecutables
    wordComRegistered = $comRegistrationPresent
    status = 'PRECONDITIONS_NOT_MET'
    reproduced = $false
    fixed = $false
    blockers = @($blockers)
    note = 'Environment probe only. A successful Windows job is not evidence of reproducing or fixing the Word popup.'
}
$json = $result | ConvertTo-Json -Depth 8
$json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'environment.json') -Encoding UTF8
Write-Output $json
$lines = @(
    '# Word popup lab V1.0: environment probe',
    '',
    ('- OS: ' + $result.operatingSystem),
    ('- Word installed: ' + ($wordExecutables.Count -gt 0 -or $comRegistrationPresent)),
    ('- Process session: ' + $desktop.processSessionId),
    ('- Input desktop: ' + $desktop.inputDesktopName),
    '- Popup reproduced: NO',
    '- Fix verified: NO',
    ('- Blockers: ' + ($blockers -join ', ')),
    '',
    'Office licensing and the matching signed-in Word profile must be supplied through an authorized interactive Windows environment.'
)
$lines | Set-Content -LiteralPath (Join-Path $OutputDirectory 'RESULT.md') -Encoding UTF8
if ($env:GITHUB_STEP_SUMMARY) { $lines | Add-Content -LiteralPath $env:GITHUB_STEP_SUMMARY -Encoding UTF8 }
exit 2
