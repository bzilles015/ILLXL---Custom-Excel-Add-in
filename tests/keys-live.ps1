# Presses EVERY shortcut in the active preset for real - low-level key presses,
# the same input a physical keyboard sends - in a visible test Excel, and checks
# the right command ran:  key -> OnKey -> RunCmd -> command.
# The test window takes focus for about a minute - don't type while it runs.
param([string]$AddIn = "$PSScriptRoot\..\build\ILLXL - Nexcess.xlam")
$ErrorActionPreference = 'Stop'
Add-Type -Namespace W -Name K -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h);
[DllImport("user32.dll")] public static extern System.IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);
[DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool attach);
[DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
[DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, System.UIntPtr extra);
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h, int cmd);
'@

# friendly key -> (virtual-key code, needs Shift) on a US keyboard
$VK = @{ 'Up' = 0x26; 'Down' = 0x28; 'Left' = 0x25; 'Right' = 0x27; 'Backspace' = 0x08; 'Delete' = 0x2E; 'Insert' = 0x2D
         'Home' = 0x24; 'End' = 0x23; 'PgUp' = 0x21; 'PgDn' = 0x22; 'Enter' = 0x0D; 'Tab' = 0x09 }
$PUNCT = @{ ';' = @(0xBA, 0); '=' = @(0xBB, 0); ',' = @(0xBC, 0); '-' = @(0xBD, 0); '.' = @(0xBE, 0); '/' = @(0xBF, 0)
            '`' = @(0xC0, 0); '[' = @(0xDB, 0); '\' = @(0xDC, 0); ']' = @(0xDD, 0); "'" = @(0xDE, 0)
            '<' = @(0xBC, 1); '>' = @(0xBE, 1); '_' = @(0xBD, 1); '+' = @(0xBB, 1); ':' = @(0xBA, 1) }
function Press([string]$friendly) {
    $parts = $friendly -split '\+(?=.)'; $key = $parts[-1]; $mods = @($parts | Select-Object -SkipLast 1)
    $shift = $mods -contains 'Shift'
    if ($VK.ContainsKey($key)) { $vk = $VK[$key] }
    elseif ($key -match '^F(\d+)$') { $vk = 0x6F + [int]$matches[1] }
    elseif ($PUNCT.ContainsKey($key)) { $vk = $PUNCT[$key][0]; if ($PUNCT[$key][1]) { $shift = $true } }
    else { $vk = [int][char]$key.ToUpper() }
    $down = @(); if ($mods -contains 'Ctrl') { $down += 0x11 }; if ($mods -contains 'Alt') { $down += 0x12 }; if ($shift) { $down += 0x10 }
    foreach ($m in $down) { [W.K]::keybd_event([byte]$m, 0, 0, [UIntPtr]::Zero) }
    # arrows / Home / End / PgUp / PgDn / Insert / Delete are "extended" keys - without the
    # flag Windows sends the numpad versions, which type digits when Num Lock is on
    $ext = [uint32]$(if ($vk -ge 0x21 -and $vk -le 0x2E) { 1 } else { 0 })
    [W.K]::keybd_event([byte]$vk, 0, $ext, [UIntPtr]::Zero); [W.K]::keybd_event([byte]$vk, 0, ($ext -bor 2), [UIntPtr]::Zero)
    [array]::Reverse($down); foreach ($m in $down) { [W.K]::keybd_event([byte]$m, 0, 2, [UIntPtr]::Zero) }
}
function Focus([IntPtr]$h) {
    $fg = [W.K]::GetForegroundWindow(); $p = 0
    $fgThread = [W.K]::GetWindowThreadProcessId($fg, [ref]$p); $me = [W.K]::GetCurrentThreadId()
    [void][W.K]::AttachThreadInput($me, $fgThread, $true)
    [void][W.K]::ShowWindow($h, 9); [void][W.K]::SetForegroundWindow($h)
    [void][W.K]::AttachThreadInput($me, $fgThread, $false)
    Start-Sleep -Milliseconds 150
    return ([W.K]::GetForegroundWindow() -eq $h)
}

$root = Join-Path $env:TEMP ("illxl_keys_" + [guid]::NewGuid().ToString('N').Substring(0, 6))
$xl = New-Object -ComObject Excel.Application
$xlPid = 0; [void][W.K]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xlPid)
$dog = Start-Job { param($p) Start-Sleep 300; Stop-Process -Id $p -Force -EA SilentlyContinue } -ArgumentList $xlPid
$pass = 0; $fail = 0
try {
    $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1; $xl.Visible = $true
    $ai = $xl.Workbooks.Open((Resolve-Path $AddIn).Path); $tag = "'" + $ai.Name + "'!"
    [void]$xl.Run($tag + 'SetSettingsRoot', $root, $true)
    [void]$xl.Run($tag + 'UiTestMode', $true)
    [void]$xl.Run($tag + 'BindAllKeys')                     # what Auto_Open does at startup
    $keys = $xl.Run($tag + 'KeyAssignments')
    $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets(1)
    $h = [IntPtr]$xl.Hwnd
    foreach ($macro in @($keys.Keys())) {
        $friendly = $keys.Item($macro); if (-not $friendly) { continue }
        $ws.Activate(); $ws.Range("B2:C3").Value2 = 5; [void]$ws.Range("B2").Select()
        [void]$xl.Run($tag + 'UiTestMode', $true); [void]$xl.Run($tag + 'UiResetLastCommand')
        $focused = $false
        foreach ($try in 1..5) { if (Focus $h) { $focused = $true; break }; Start-Sleep -Milliseconds 400 }
        if (-not $focused) { $fail++; "SKIP  $friendly - couldn't give the test window focus (something else kept it)"; continue }
        Press $friendly
        Start-Sleep -Milliseconds 400
        $got = $xl.Run($tag + 'UiLastCommand')
        if ($macro -eq 'TogglePerformanceMode' -and $got -eq $macro) { [void]$xl.Run($tag + 'RunCmd', $macro) }   # back off
        if ($got -eq $macro) { $pass++; "ok    $friendly" } else { $fail++; "FAIL  $friendly should run $macro - got '$got'" }
    }
    $wb.Close($false)
} catch { "ERROR: $($_.Exception.Message)"; $fail++ }
finally {
    try { [void]$xl.Run($tag + 'UnbindAllKeys'); $ai.Close($false); $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Stop-Job $dog -EA SilentlyContinue; Remove-Job $dog -Force -EA SilentlyContinue
    Remove-Item $root -Recurse -Force -EA SilentlyContinue
}
"$pass keys reached their command, $fail failed"
