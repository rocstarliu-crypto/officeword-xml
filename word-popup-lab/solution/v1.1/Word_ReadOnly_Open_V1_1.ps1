# 老板可读区流程图：选择原文件 → 核对只读 → Word 只读打开 → 核对原路径与页面视图 → 写日志
# V1.1：便携式打开工具。源文件、Office 隐私和 OneDrive 设置不写入。
[CmdletBinding()]
param([string]$FilePath = '', [string]$ResultJson = '')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$logDirectory = Join-Path $env:LOCALAPPDATA 'EIAWordReadonlyOpen\V1_1\logs'
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null

if (-not ('EIAWordReadonlyNativeV11' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class EIAWordReadonlyNativeV11 {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
    public delegate bool EnumWindowCallback(IntPtr handle, IntPtr extra);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowCallback callback, IntPtr extra);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder name, int count);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr handle, int command);
    public static IntPtr FindWordWindow(uint processId) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr h, IntPtr unused) {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p != processId) return true;
            var name = new StringBuilder(128); GetClassName(h, name, name.Capacity);
            if (name.ToString() == "OpusApp") { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
'@
}

function Open-ReadonlyOriginal([string]$SelectedPath) {
    $word = $null; $document = $null; $blank = $null; $original = $null
    $ownedWordApp = $false
    $createdProcessId = 0
    $knownWordIds = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $entry = [ordered]@{ version = '1.1'; time = (Get-Date).ToString('o'); success = $false; method = 'Documents.Open(ReadOnly=True)' }
    try {
        if (-not (Test-Path -LiteralPath $SelectedPath -PathType Leaf)) { throw '没有找到该文件。' }
        $file = Get-Item -LiteralPath $SelectedPath
        $fullPath = $file.FullName
        if ($file.Extension.ToLowerInvariant() -notin @('.docx','.doc','.docm','.dotx','.dot','.dotm','.rtf')) { throw '请选择 Word 文档。' }
        $entry.file = $fullPath
        $entry.fileReadonlyBefore = [bool]($file.Attributes -band [IO.FileAttributes]::ReadOnly)
        if (-not $entry.fileReadonlyBefore) { throw '该文件没有设置只读。请先在文件属性中设置只读，再使用本工具。' }
        $entry.stage = 'FindNativeWord'
        $paths = @("$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE", "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE", "$env:ProgramFiles\Microsoft Office\Office16\WINWORD.EXE", "${env:ProgramFiles(x86)}\Microsoft Office\Office16\WINWORD.EXE")
        foreach ($registry in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE')) {
            if (Test-Path -LiteralPath $registry) { $paths += (Get-Item -LiteralPath $registry).GetValue('') }
        }
        $executable = @($paths | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1)
        if ($executable.Count -eq 0) { throw '没有找到已安装的 Microsoft Word。' }
        $entry.stage = 'StartNativeWord'
        $nativeProcess = Start-Process -FilePath $executable[0] -ArgumentList ('/n "' + $fullPath + '"') -PassThru
        $createdProcessId = $nativeProcess.Id
        $entry.nativeProcessId = $createdProcessId
        $entry.stage = 'BindOwnedWord'
        $ready = $false
        $lastBindError = ''
        for ($attempt = 0; $attempt -lt 60; $attempt++) {
            $handle = [EIAWordReadonlyNativeV11]::FindWordWindow([uint32]$createdProcessId)
            if ($handle -ne [IntPtr]::Zero) {
                [void][EIAWordReadonlyNativeV11]::ShowWindow($handle, 6)
                Start-Sleep -Milliseconds 300
                try {
                    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
                    if ([long]$word.ActiveWindow.Hwnd -ne [long]$handle) {
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word); $word = $null
                        $boundDocument = [Runtime.InteropServices.Marshal]::BindToMoniker($fullPath)
                        $word = $boundDocument.Application
                    }
                    [uint32]$boundProcessId = 0
                    [void][EIAWordReadonlyNativeV11]::GetWindowThreadProcessId([IntPtr]$word.ActiveWindow.Hwnd, [ref]$boundProcessId)
                    if ([int]$boundProcessId -ne $createdProcessId) { throw '请先关闭这份文档原来打开的窗口，再使用本工具。' }
                    if (-not [string]::Equals([IO.Path]::GetFullPath([string]$word.ActiveDocument.FullName), $fullPath, [StringComparison]::OrdinalIgnoreCase)) { throw 'Word 原文件尚未准备就绪。' }
                    $ownedWordApp = $true; $ready = $true; break
                } catch { $lastBindError = $_.Exception.Message; $word = $null }
            }
            Start-Sleep -Milliseconds 500
        }
        if (-not $ready) { throw ('未连接到本工具新打开的 Word：' + $lastBindError) }
        $entry.ownedNativeInstanceConfirmed = $true
        $entry.stage = 'PrepareReadonlyReopen'
        $original = $word.ActiveDocument
        $blank = $word.Documents.Add()
        $original.Close(0); $original = $null
        $entry.stage = 'OpenExplicitReadonly'
        $securityBefore = $word.AutomationSecurity
        try {
            # 明确只读再打开时临时禁用宏；首次原生打开遵循 Word 原有安全设置。
            $word.AutomationSecurity = 3
            $document = $word.Documents.Open($fullPath, $false, $true, $false)
        } finally {
            $word.AutomationSecurity = $securityBefore
        }
        $entry.stage = 'ClosePlaceholder'
        $blank.Close(0); $blank = $null
        $entry.stage = 'VerifyOpenState'
        $word.ActiveWindow.View.Type = 3
        $word.ActiveWindow.WindowState = 1
        $entry.documentReadonly = [bool]$document.ReadOnly
        $entry.sameOriginalPath = [string]::Equals([IO.Path]::GetFullPath([string]$document.FullName), $fullPath, [StringComparison]::OrdinalIgnoreCase)
        $entry.fileReadonlyAfter = [bool]((Get-Item -LiteralPath $fullPath).Attributes -band [IO.FileAttributes]::ReadOnly)
        $entry.viewType = [int]$word.ActiveWindow.View.Type
        $entry.protectedViewCount = [int]$word.ProtectedViewWindows.Count
        $entry.wordVersion = [string]$word.Version
        $entry.wordBuild = [string]$word.Build
        $entry.wordProcessId = $createdProcessId
        $entry.wordWindowHandle = [long]$word.ActiveWindow.Hwnd
        if (-not ($entry.documentReadonly -and $entry.sameOriginalPath -and $entry.fileReadonlyAfter -and $entry.viewType -eq 3 -and $entry.protectedViewCount -eq 0)) {
            throw 'Word 返回的打开状态不符合要求，已停止；请使用“在受保护的视图中打开”。'
        }
        $entry.stage = 'Complete'
        $entry.success = $true
        # 保留可见的 Word 窗口。不保存文档，也不关闭用户其他 Word 窗口。
    } catch {
        $entry.error = $_.Exception.Message
        $entry.errorPosition = $_.InvocationInfo.PositionMessage
        if ($null -ne $document) { try { $document.Close(0) } catch {} }
        if ($null -ne $blank) { try { $blank.Close(0) } catch {} }
        if ($ownedWordApp -and $null -ne $word -and $createdProcessId -gt 0 -and $createdProcessId -notin $knownWordIds) {
            try { $word.Quit(0) } catch {}
        }
    } finally {
        $logPath = Join-Path $logDirectory ((Get-Date -Format 'yyyyMMdd_HHmmss_fff') + '.json')
        $entry.logPath = $logPath
        $json = $entry | ConvertTo-Json -Depth 5
        $json | Set-Content -LiteralPath $logPath -Encoding UTF8
        if ($ResultJson) { $json | Set-Content -LiteralPath $ResultJson -Encoding UTF8 }
        foreach ($item in @($document, $blank, $original, $word)) {
            if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)
            }
        }
    }
    if (-not $entry.success) { throw $entry.error }
    Write-Host '已用原文件在普通页面视图中只读打开。请在 Word 中选字、复制。' -ForegroundColor Green
    Write-Host ('日志：' + $entry.logPath)
}

if ($FilePath) {
    try { Open-ReadonlyOriginal $FilePath; exit 0 }
    catch { Write-Error $_.Exception.Message; exit 1 }
}

Add-Type -AssemblyName System.Windows.Forms
while ($true) {
    Write-Host ''
    Write-Host 'Word 只读打开工具 V1.1'
    Write-Host '1. 选择只读 Word 文档并打开（直接回车）'
    Write-Host '0. 查看程序工作流程'
    Write-Host '2. 查看日志文件夹'
    Write-Host '9. 退出工具'
    $choice = Read-Host '请选择'
    if ($choice -eq '9') { break }
    if ($choice -eq '0') { Write-Host '选文件 → 确认只读 → 原路径只读打开 → 核对页面视图 → 留存日志'; continue }
    if ($choice -eq '2') { Invoke-Item -LiteralPath $logDirectory; continue }
    if ($choice -notin @('', '1')) { continue }
    $picker = New-Object System.Windows.Forms.OpenFileDialog
    $picker.Title = '选择已设置只读的 Word 文档'
    $picker.Filter = 'Word 文档|*.docx;*.doc;*.docm;*.dotx;*.dot;*.dotm;*.rtf'
    $picker.Multiselect = $false
    try {
        if ($picker.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
            try { Open-ReadonlyOriginal $picker.FileName }
            catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, '打开未完成') }
        }
    } finally { $picker.Dispose() }
}
