# Opens the Settings window on each page in a hidden Excel and saves PNGs
# to tests\screens\ (for checking the layout without clicking through it).
param([string]$AddIn = "$PSScriptRoot\..\build\ILLXL - Nexcess.xlam")
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -Namespace W -Name U -MemberDefinition @'
[DllImport("user32.dll")] public static extern System.IntPtr FindWindow(string c, string t);
[DllImport("user32.dll")] public static extern bool PrintWindow(System.IntPtr h, System.IntPtr dc, uint f);
[DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);
public struct RECT { public int L, T, R, B; }
'@
$out = Join-Path $PSScriptRoot 'screens'; New-Item -ItemType Directory -Force $out | Out-Null
$root = Join-Path $env:TEMP ("illxl_shots_" + [guid]::NewGuid().ToString('N').Substring(0, 6))

$xl = New-Object -ComObject Excel.Application
$xlPid = 0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xlPid)
$dog = Start-Job { param($p) Start-Sleep 120; Stop-Process -Id $p -Force -EA SilentlyContinue } -ArgumentList $xlPid
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1
    $ai = $xl.Workbooks.Open((Resolve-Path $AddIn).Path); $tag = "'" + $ai.Name + "'!"
    [void]$xl.Run($tag + 'SetSettingsRoot', $root, $true)
    foreach ($p in 'presets', 'formats', 'colors', 'fonts', 'tags', 'keys') {
        [void]$xl.Run($tag + 'DemoSettings', $p)
        Start-Sleep -Milliseconds 700
        $h = [W.U]::FindWindow('ThunderDFrame', 'ILLXL Settings')
        if ($h -eq [IntPtr]::Zero) { "no window for $p"; continue }
        $r = New-Object W.U+RECT; [void][W.U]::GetWindowRect($h, [ref]$r)
        $bmp = New-Object Drawing.Bitmap ($r.R - $r.L), ($r.B - $r.T)
        $g = [Drawing.Graphics]::FromImage($bmp); $dc = $g.GetHdc()
        [void][W.U]::PrintWindow($h, $dc, 2); $g.ReleaseHdc($dc); $g.Dispose()
        $bmp.Save("$out\$p.png"); $bmp.Dispose()
        "saved $out\$p.png"
    }
    [void]$xl.Run($tag + 'DemoSettings', '__close')
} finally {
    try { foreach ($f in @($xl.VBE)) {} ; $ai.Close($false); $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Stop-Job $dog -EA SilentlyContinue; Remove-Job $dog -Force -EA SilentlyContinue
    Remove-Item $root -Recurse -Force -EA SilentlyContinue
}
