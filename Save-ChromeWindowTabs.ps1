[CmdletBinding(DefaultParameterSetName = 'Export')]
param(
    [Parameter(ParameterSetName = 'Restore', Mandatory = $true)]
    [string]$RestoreFromFile
)

function Restart-InSTAIfNeeded {
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -eq [System.Threading.ApartmentState]::STA) {
        return
    }

    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', '"' + $PSCommandPath + '"')
    if ($PSBoundParameters.ContainsKey('RestoreFromFile')) {
        $argList += @('-RestoreFromFile', '"' + $RestoreFromFile + '"')
    }

    Start-Process -FilePath 'powershell.exe' -ArgumentList ($argList -join ' ') -Wait
    exit $LASTEXITCODE
}

function Add-NativeWindowInterop {
    Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class NativeWindowInterop
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@
}

function Get-WindowText {
    param([IntPtr]$Handle)

    $length = [NativeWindowInterop]::GetWindowTextLength($Handle)
    if ($length -le 0) {
        return ''
    }

    $builder = New-Object System.Text.StringBuilder ($length + 1)
    [void][NativeWindowInterop]::GetWindowText($Handle, $builder, $builder.Capacity)
    return $builder.ToString()
}

function Get-WindowClassName {
    param([IntPtr]$Handle)

    $builder = New-Object System.Text.StringBuilder 256
    [void][NativeWindowInterop]::GetClassName($Handle, $builder, $builder.Capacity)
    return $builder.ToString()
}

function Get-DetectedWindowName {
    param([string]$WindowTitle)

    if ([string]::IsNullOrWhiteSpace($WindowTitle)) {
        return 'Unnamed'
    }

    $baseTitle = $WindowTitle.Trim()
    foreach ($suffix in @(' - Google Chrome', ' - Chromium', ' - Google Chrome (Developer Build)')) {
        if ($baseTitle.EndsWith($suffix)) {
            $baseTitle = $baseTitle.Substring(0, $baseTitle.Length - $suffix.Length)
            break
        }
    }

    if ($baseTitle.EndsWith(' - Incognito')) {
        $baseTitle = $baseTitle.Substring(0, $baseTitle.Length - ' - Incognito'.Length)
    }

    $segments = $baseTitle -split ' - '
    $detected = if ($segments.Count -ge 2) { $segments[$segments.Count - 1] } else { $baseTitle }

    if ([string]::IsNullOrWhiteSpace($detected)) {
        return 'Unnamed'
    }

    return $detected.Trim()
}

function Get-ChromeWindows {
    $windows = New-Object System.Collections.Generic.List[object]

    $callback = [NativeWindowInterop+EnumWindowsProc]{
        param([IntPtr]$hWnd, [IntPtr]$lParam)

        if (-not [NativeWindowInterop]::IsWindowVisible($hWnd)) {
            return $true
        }

        $className = Get-WindowClassName -Handle $hWnd
        if (-not $className.StartsWith('Chrome_WidgetWin_')) {
            return $true
        }

        [uint32]$windowProcessId = 0
        [void][NativeWindowInterop]::GetWindowThreadProcessId($hWnd, [ref]$windowProcessId)
        if ($windowProcessId -eq 0) {
            return $true
        }

        $proc = Get-Process -Id $windowProcessId -ErrorAction SilentlyContinue
        if (-not $proc -or $proc.ProcessName -ne 'chrome') {
            return $true
        }

        $title = Get-WindowText -Handle $hWnd
        if ([string]::IsNullOrWhiteSpace($title)) {
            return $true
        }

        $memoryMb = [math]::Round($proc.WorkingSet64 / 1MB, 2)
        $windowName = Get-DetectedWindowName -WindowTitle $title

        $windows.Add([PSCustomObject]@{
            Handle = $hWnd
            ProcessId = [int]$windowProcessId
            WindowTitle = $title
            WindowName = $windowName
            MemoryMB = $memoryMb
        }) | Out-Null

        return $true
    }

    [void][NativeWindowInterop]::EnumWindows($callback, [IntPtr]::Zero)

    $id = 1
    $sorted = $windows | Sort-Object WindowName, WindowTitle
    foreach ($item in $sorted) {
        $item | Add-Member -MemberType NoteProperty -Name RunningId -Value $id
        $id++
    }

    return $sorted
}

function Read-TabCountFallback {
    while ($true) {
        $inputValue = Read-Host 'Could not detect tab count automatically. Enter number of tabs to export'
        $parsed = 0
        if ([int]::TryParse($inputValue, [ref]$parsed) -and $parsed -gt 0) {
            return $parsed
        }
        Write-Host 'Please enter a positive whole number.' -ForegroundColor Yellow
    }
}

function Get-ChromeTabCount {
    param([IntPtr]$Handle)

    try {
        Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
        $windowElement = [System.Windows.Automation.AutomationElement]::FromHandle($Handle)
        if (-not $windowElement) {
            return 0
        }

        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::TabItem
        )

        $tabs = $windowElement.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
        return [int]$tabs.Count
    }
    catch {
        return 0
    }
}

function Get-ClipboardTextWithRetry {
    param([int]$MaxAttempts = 8)

    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try {
            if ([System.Windows.Forms.Clipboard]::ContainsText()) {
                return [System.Windows.Forms.Clipboard]::GetText().Trim()
            }
        }
        catch {
            Start-Sleep -Milliseconds 120
        }

        Start-Sleep -Milliseconds 80
    }

    return ''
}

function Copy-ActiveTabUrl {
    [System.Windows.Forms.SendKeys]::SendWait('^l')
    Start-Sleep -Milliseconds 120
    [System.Windows.Forms.SendKeys]::SendWait('^c')
    Start-Sleep -Milliseconds 170
    return (Get-ClipboardTextWithRetry)
}

function Focus-ChromeWindow {
    param([pscustomobject]$Window)

    # 9 = SW_RESTORE (restore if minimized)
    [void][NativeWindowInterop]::ShowWindow($Window.Handle, 9)
    [void][NativeWindowInterop]::SetForegroundWindow($Window.Handle)
    [void][Microsoft.VisualBasic.Interaction]::AppActivate($Window.ProcessId)
    Start-Sleep -Milliseconds 350
}

function Export-ChromeWindowTabs {
    param([pscustomobject]$Window)

    Focus-ChromeWindow -Window $Window

    $originalClipboard = ''
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsText()) {
            $originalClipboard = [System.Windows.Forms.Clipboard]::GetText()
        }
    }
    catch {
        $originalClipboard = ''
    }

    $tabCount = Get-ChromeTabCount -Handle $Window.Handle
    if ($tabCount -le 0) {
        $tabCount = Read-TabCountFallback
    }

    $urls = New-Object System.Collections.Generic.List[string]

    Write-Host ''
    Write-Host "Capturing $tabCount tab URL(s) from selected window..." -ForegroundColor Yellow

    for ($i = 1; $i -le $tabCount; $i++) {
        Write-Progress -Activity 'Reading Chrome tabs' -Status "Tab $i of $tabCount" -PercentComplete ([math]::Round(($i / $tabCount) * 100))

        $url = Copy-ActiveTabUrl
        if ([string]::IsNullOrWhiteSpace($url)) {
            $url = '[Could not read URL]'
        }

        $urls.Add($url) | Out-Null

        if ($i -lt $tabCount) {
            [System.Windows.Forms.SendKeys]::SendWait('^{TAB}')
            Start-Sleep -Milliseconds 180
        }
    }

    Write-Progress -Activity 'Reading Chrome tabs' -Completed

    try {
        if (-not [string]::IsNullOrEmpty($originalClipboard)) {
            [System.Windows.Forms.Clipboard]::SetText($originalClipboard)
        }
    }
    catch {
        # Clipboard restore is best effort only.
    }

    $safeName = ($Window.WindowName -replace '[\\/:*?"<>|]', '_').Trim()
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = "Unnamed-$($Window.RunningId)"
    }

    $fileName = "Chrome-Window-$safeName-tabs.txt"
    $outputPath = Join-Path $PSScriptRoot $fileName

    if (Test-Path $outputPath) {
        $overwrite = Read-Host "File already exists: $outputPath. Overwrite? (Y/N)"
        if ($overwrite -notin @('Y', 'y')) {
            Write-Host 'Export cancelled by user.' -ForegroundColor Yellow
            return
        }
    }

    $urls | Set-Content -Path $outputPath -Encoding UTF8

    Write-Host ''
    Write-Host 'Export complete.' -ForegroundColor Green
    Write-Host "Window Name: $($Window.WindowName)" -ForegroundColor White
    Write-Host "Tabs exported: $($urls.Count)" -ForegroundColor White
    Write-Host "Saved file: $outputPath" -ForegroundColor White
}

function Resolve-ChromeExecutable {
    $chrome = Get-Command chrome.exe -ErrorAction SilentlyContinue
    if ($chrome) {
        return $chrome.Source
    }

    $candidates = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Restore-ChromeWindowTabs {
    param([string]$TabsFile)

    if (-not (Test-Path $TabsFile)) {
        Write-Host "Tabs file not found: $TabsFile" -ForegroundColor Red
        exit 1
    }

    $urls = Get-Content -Path $TabsFile | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if (-not $urls -or $urls.Count -eq 0) {
        Write-Host "No URLs found in file: $TabsFile" -ForegroundColor Red
        exit 1
    }

    $chromeExe = Resolve-ChromeExecutable
    if (-not $chromeExe) {
        Write-Host 'Could not locate chrome.exe.' -ForegroundColor Red
        exit 1
    }

    $arguments = @('--new-window') + $urls
    Start-Process -FilePath $chromeExe -ArgumentList $arguments

    Write-Host "Opened new Chrome window with $($urls.Count) tab(s)." -ForegroundColor Green
}

Restart-InSTAIfNeeded
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic
Add-NativeWindowInterop

Write-Host 'Chrome Window Tabs Tool' -ForegroundColor Cyan
Write-Host '=======================' -ForegroundColor Cyan

if ($PSCmdlet.ParameterSetName -eq 'Restore') {
    Write-Host 'Mode: Restore from tabs file' -ForegroundColor Yellow
    Restore-ChromeWindowTabs -TabsFile $RestoreFromFile
    exit 0
}

$chromeWindows = Get-ChromeWindows
if (-not $chromeWindows -or $chromeWindows.Count -eq 0) {
    Write-Host 'No visible Chrome windows were found.' -ForegroundColor Red
    exit 1
}

Write-Host ''
Write-Host 'Detected Chrome windows:' -ForegroundColor Green
Write-Host ''
Write-Host ('{0,-4} {1,-28} {2,12} {3}' -f 'ID', 'Window Name', 'RAM MB', 'Window Title') -ForegroundColor White
Write-Host ('{0,-4} {1,-28} {2,12} {3}' -f '--', '-----------', '------', '------------') -ForegroundColor DarkGray

foreach ($window in $chromeWindows) {
    Write-Host ('{0,-4} {1,-28} {2,12} {3}' -f $window.RunningId, $window.WindowName, $window.MemoryMB, $window.WindowTitle)
}

Write-Host ''
Write-Host 'RAM value is based on the owning chrome.exe process working set for each visible window.' -ForegroundColor DarkGray

$selectedWindow = $null
while (-not $selectedWindow) {
    $selection = Read-Host 'Enter the running ID of the Chrome window to export'
    $selectedId = 0

    if (-not [int]::TryParse($selection, [ref]$selectedId)) {
        Write-Host 'Please enter a valid number.' -ForegroundColor Yellow
        continue
    }

    $selectedWindow = $chromeWindows | Where-Object { $_.RunningId -eq $selectedId } | Select-Object -First 1
    if (-not $selectedWindow) {
        Write-Host 'No window matches that ID. Try again.' -ForegroundColor Yellow
    }
}

Export-ChromeWindowTabs -Window $selectedWindow

Write-Host ''
Write-Host 'Next phase is available now: run with -RestoreFromFile to open a new window from an exported tabs file.' -ForegroundColor Cyan
