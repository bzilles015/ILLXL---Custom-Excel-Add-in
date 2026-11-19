# Opens the add-in in a VISIBLE Excel window, checks the ribbon loaded, switches
# to the ILLXL tab and saves tests\screens\ribbon.png. (A window appears briefly.)
param([string]$AddIn = "$PSScriptRoot\..\build\ILLXL - Nexcess.xlam")
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -Namespace W -Name U -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool PrintWindow(System.IntPtr h, System.IntPtr dc, uint f);
[DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);
public struct RECT { public int L, T, R, B; }
'@
$out = Join-Path $PSScriptRoot 'screens'; New-Item -ItemType Directory -Force $out | Out-Null
$root = Join-Path $env:TEMP ("illxl_rib_" + [guid]::NewGuid().ToString('N').Substring(0, 6))
$xl = New-Object -ComObject Excel.Application
$xlPid = 0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xlPid)
$dog = Start-Job { param($p) Start-Sleep 90; Stop-Process -Id $p -Force -EA SilentlyContinue } -ArgumentList $xlPid
try {
    $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1
    $xl.WindowState = -4143; $xl.Left = 0; $xl.Top = 0; $xl.Width = 1880; $xl.Height = 300; $xl.Visible = $true
    $ai = $xl.Workbooks.Open((Resolve-Path $AddIn).Path); $tag = "'" + $ai.Name + "'!"
    [void]$xl.Run($tag + 'SetSettingsRoot', $root, $true)
    [void]$xl.Workbooks.Add()
    Start-Sleep -Seconds 3
    "ribbon loaded: " + $xl.Run($tag + 'RibbonLoaded')
    # Expand a collapsed ribbon for the shot, and put it back afterwards (user setting)
    $wasMin = $xl.CommandBars.GetPressedMso('MinimizeRibbon')
    if ($wasMin) { $xl.CommandBars.ExecuteMso('MinimizeRibbon') }
    [void]$xl.Run($tag + 'ShowRibbonTab')
    Start-Sleep -Seconds 2
    $h = [IntPtr]$xl.Hwnd
    $r = New-Object W.U+RECT; [void][W.U]::GetWindowRect($h, [ref]$r)
    $bmp = New-Object Drawing.Bitmap ($r.R - $r.L), ($r.B - $r.T)
    $g = [Drawing.Graphics]::FromImage($bmp); $dc = $g.GetHdc()
    [void][W.U]::PrintWindow($h, $dc, 2); $g.ReleaseHdc($dc); $g.Dispose()
    $bmp.Save("$out\ribbon.png"); $bmp.Dispose()
    if ($wasMin) { $xl.CommandBars.ExecuteMso('MinimizeRibbon'); 'ribbon collapsed again (your setting)' }
    "saved $out\ribbon.png"
} finally {
    try { foreach ($wb in @($xl.Workbooks)) { $wb.Close($false) }; $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Stop-Job $dog -EA SilentlyContinue; Remove-Job $dog -Force -EA SilentlyContinue
    Remove-Item $root -Recurse -Force -EA SilentlyContinue
}
