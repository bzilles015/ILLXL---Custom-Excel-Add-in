# ILLXL regression tests - runs the built add-in in a hidden Excel instance.
# Usage:  pwsh -File tests\run-tests.ps1
param([string]$AddIn = "$PSScriptRoot\..\build\ILLXL - Nexcess.xlam")
$ErrorActionPreference = 'Stop'
$AddIn = (Resolve-Path $AddIn).Path
$script:pass = 0; $script:fail = 0

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1                       # allow the add-in's macros (our own file)

# Watchdog: if a VBA error dialog freezes THIS hidden Excel, kill only its PID
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$xlPid = 0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xlPid)
$dog = Start-Job { param($p) Start-Sleep 240; Stop-Process -Id $p -Force -EA SilentlyContinue } -ArgumentList $xlPid

$ai = $xl.Workbooks.Open($AddIn)
$tag = "'" + $ai.Name + "'!"
function Run($m) { [void]$xl.Run($tag + $m) }
function RunV($m, $a) {
    # a VBA "null string" can't cross COM - treat it as ""
    try { if ($PSBoundParameters.ContainsKey('a')) { $xl.Run($tag + $m, $a) } else { $xl.Run($tag + $m) } }
    catch { if ($_.Exception.Message -match "Parameter 'ptr'") { '' } else { throw } }
}

# Settings go to a throwaway folder - never touches the real %APPDATA%\ILLXL
$root = Join-Path $env:TEMP ("illxl_test_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
[void]$xl.Run($tag + "SetSettingsRoot", $root, $true)
[void]$xl.Run($tag + "UiTestMode", $true)                 # prompts answered by the tests, never shown
function Check($name, $cond, $detail = '') {
    if ($cond) { $script:pass++; "PASS  $name" } else { $script:fail++; "FAIL  $name  $detail" }
}
function Sel($ws, $addr) { $ws.Activate(); [void]$ws.Range($addr).Select() }
$NONE = -4142

try {
    # --- Full compile (Debug > Compile). Success greys the menu item out. ------
    $vbe = $xl.VBE; $vbe.ActiveVBProject = $ai.VBProject
    $ctl = $vbe.CommandBars.FindControl(1, 578)
    if ($ctl.Enabled) { $ctl.Execute() }          # a compile error would freeze here -> watchdog
    Check "VBA project compiles" (-not $vbe.CommandBars.FindControl(1, 578).Enabled)

    $wb = $xl.Workbooks.Add(); $wb.Activate()
    while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add() }
    $s1 = $wb.Worksheets(1); $s2 = $wb.Worksheets(2)

    # --- 2) Scale / sign leave blanks and text alone ---------------------------
    $s1.Range("A2").Value2 = 1000; $s1.Range("A3").Value2 = "text"; $s1.Range("A4").Formula = "=A2"
    Sel $s1 "A1:A4"; Run "ScaleUp"
    Check "ScaleUp: blank stays blank"  ($s1.Range("A1").Formula -eq "")
    Check "ScaleUp: number scaled"      ($s1.Range("A2").Value2 -eq 1)
    Check "ScaleUp: text untouched"     ($s1.Range("A3").Value2 -eq "text")
    Check "ScaleUp: formula wrapped"    ($s1.Range("A4").Formula -eq "=(A2)/1000") $s1.Range("A4").Formula
    Run "ToggleSign"
    Check "ToggleSign: blank stays blank" ($s1.Range("A1").Formula -eq "")
    Check "ToggleSign: number flipped"    ($s1.Range("A2").Value2 -eq -1)

    # --- 3) Text case skips formulas; single cell stays single -----------------
    $s1.Range("B1").Value2 = "hello world"; $s1.Range("B2").Formula = '=B1&"x"'; $s1.Range("B3").Value2 = 5
    $s1.Range("Z20").Value2 = "far away"
    Sel $s1 "B1:B3"; Run "CycleTextCase"
    Check "TextCase: text converted"     ($s1.Range("B1").Value2 -eq "Hello World") $s1.Range("B1").Value2
    Check "TextCase: formula kept"       ($s1.Range("B2").HasFormula)
    Sel $s1 "B1"; Run "CycleTextCase"
    Check "TextCase: 1-cell stays local" ($s1.Range("Z20").Value2 -eq "far away") $s1.Range("Z20").Value2

    # --- 4/5) Undo restores formatting, on the right sheet ---------------------
    $s1.Range("C1:C3").Value2 = 7
    Sel $s1 "C1:C3"; Run "CycleFill"; Run "CycleFill"          # yellow
    $filled = $s1.Range("C1").Interior.Color
    Sel $s2 "C1"; Run "UndoLastAction"                          # switched sheets first
    Check "Undo: fill removed on Sheet1" ($s1.Range("C1").Interior.ColorIndex -eq $NONE) $s1.Range("C1").Interior.ColorIndex
    Check "Undo: Sheet2 untouched"       ($s2.Range("C1").Interior.ColorIndex -eq $NONE)
    Check "Undo: values kept"            ($s1.Range("C3").Value2 -eq 7)

    Sel $s1 "C1"; Run "CycleFill"; Run "CycleFill"              # C1 yellow
    Sel $s1 "D1:D6000"; Run "CycleFill"                         # too big to snapshot
    Run "UndoLastAction"
    Check "Undo >5000: older action not replayed" ($s1.Range("C1").Interior.ColorIndex -ne $NONE)

    $s1.Range("E1").Formula2 = "=SEQUENCE(3)"
    Sel $s1 "E1:E3"; Run "CycleFill"; Run "CycleFill"; Run "UndoLastAction"
    Check "Undo: spill formula intact"   ($s1.Range("E1").Formula2 -eq "=SEQUENCE(3)") $s1.Range("E1").Formula2
    Check "Undo: spill still spills"     ($s1.Range("E3").Value2 -eq 3) $s1.Range("E3").Text

    $s1.Range("K1").NumberFormat = "General"; $s1.Range("K1").Value2 = "'00123"
    Sel $s1 "K1"; Run "CycleFill"; Run "UndoLastAction"
    Check "Undo: text '00123' stays text" ($s1.Range("K1").Value2 -is [string]) $s1.Range("K1").Value2

    $s1.Range("M1").Value2 = 5; $s1.Range("M1").Font.Bold = $true
    Sel $s1 "M1"; Run "CycleHeaderStyle"; Run "UndoLastAction"
    Check "Undo: header style reverted"  ($s1.Range("M1").Interior.ColorIndex -eq $NONE)

    Sel $s1 "P1"; Run "InsertVarianceHeaders"; Run "UndoLastAction"
    Check "Undo: variance headers removed (P1:R1)" ($xl.WorksheetFunction.CountA($s1.Range("P1:R1")) -eq 0)

    Sel $s1 "A1"; Run "CycleTabColor"; Run "CycleTabColor"; Run "UndoLastAction"
    Check "Undo: tab color restored"     ($s1.Tab.ColorIndex -eq $NONE) $s1.Tab.ColorIndex

    # --- 9) Font size with mixed sizes -----------------------------------------
    $s1.Range("F1").Value2 = 1; $s1.Range("F2").Value2 = 2
    $s1.Range("F1").Font.Size = 14; $s1.Range("F2").Font.Size = 11
    Sel $s1 "F1:F2"; Run "IncreaseFontSize"
    Check "FontSize mixed: 14->15, 11->12" (($s1.Range("F1").Font.Size -eq 15) -and ($s1.Range("F2").Font.Size -eq 12))

    # --- 1) Auto-color: tables are not external; single cell stays local -------
    $s1.Range("H1").Value2 = "Col"; $s1.Range("H2").Value2 = 1; $s1.Range("H3").Value2 = 2
    $lo = $s1.ListObjects.Add(1, $s1.Range("H1:H3"), $null, 1); $lo.Name = "tblT"
    $s1.Range("G1").Formula = "=SUM(tblT[Col])"
    $s1.Range("G2").Formula = "=" + $s2.Name + "!A1"
    $wb2 = $xl.Workbooks.Add(); $wb.Activate()
    $s1.Range("G3").Formula = "=[" + $wb2.Name + "]Sheet1!A1"
    $s1.Range("Z30").Value2 = 99
    Sel $s1 "G1:G3"; Run "AutoColorSelection"
    $red = 255
    Check "AutoColor: table formula not red" ($s1.Range("G1").Font.Color -ne $red) $s1.Range("G1").Font.Color
    Check "AutoColor: cross-sheet green"     ($s1.Range("G2").Font.Color -eq 32768) $s1.Range("G2").Font.Color
    Check "AutoColor: external red"          ($s1.Range("G3").Font.Color -eq $red) "$($s1.Range('G3').Font.Color) / $($s1.Range('G3').Formula)"
    Sel $s1 "G2"; Run "AutoColorSelection"
    Check "AutoColor: 1-cell stays local"    ($s1.Range("Z30").Font.Color -eq 0) $s1.Range("Z30").Font.Color
    $wb2.Close($false)

    # --- 10) Shape selected: no crash ------------------------------------------
    $shp = $s1.Shapes.AddShape(1, 300, 300, 50, 50); $s1.Activate(); [void]$shp.Select()
    $ok = $true
    foreach ($m in "CycleFill","CycleFont","CycleFontColor","CycleTextCase","CycleDateFormat",
                   "CycleOtherNumbers","IncreaseFontSize","CenterAcrossSelection","InsertStaticNow") {
        try { Run $m } catch { $ok = $false; "      $m -> $($_.Exception.Message)" }
    }
    Check "Shape selected: no errors" $ok

    # --- 6) Performance mode acts on the model ---------------------------------
    $wb.Activate(); Run "TogglePerformanceMode"
    Check "PerfMode ON: model CF calc off" (-not $s1.EnableFormatConditionsCalculation)
    Check "PerfMode ON: manual calc"       ($xl.Calculation -eq -4135)
    Run "TogglePerformanceMode"
    Check "PerfMode OFF: CF calc back on"  ($s1.EnableFormatConditionsCalculation)
    Check "PerfMode OFF: automatic calc"   ($xl.Calculation -eq -4105)

    # --- 7/8) Load / unload run cleanly ----------------------------------------
    $ok = $true; try { Run "Auto_Open"; Run "Auto_Close" } catch { $ok = $false; $_.Exception.Message }
    Check "Auto_Open / Auto_Close run" $ok

    # =========================== PRESETS ======================================
    $pdir = "$root\presets"
    Check "Presets: 3 shipped presets written on first run" ((Get-ChildItem $pdir -Filter *.ini).Count -eq 3) (Get-ChildItem $pdir | Out-String)
    Check "Presets: Nexcess is active by default" ((RunV "PresetNames") -contains "Nexcess")

    # Friendly keys convert to the exact codes the old hard-coded table used
    $expect = @{ 'Ctrl+Alt+Shift+M'='^%+M'; 'Ctrl+Shift+1'='^+1'; 'Ctrl+Shift+.'='^+.'; 'Alt+Shift+<'='%+<';
                 'Ctrl+Alt+Shift+\'='^%+\'; 'Ctrl+Alt+2'='^%2'; 'Ctrl+Alt+L'='^%l'; 'Ctrl+Alt+='='^%{=}';
                 'Ctrl+Alt+Shift+='='^%+{=}'; "Ctrl+'"="^'"; 'Ctrl+Alt+-'='^%{-}'; 'Ctrl+Alt+_'='^%{_}';
                 'Ctrl+Alt+Shift+Up'='^%+{UP}'; 'Ctrl+Alt+Shift+Backspace'='^%+{BACKSPACE}'; 'F1'='{F1}';
                 'ScrollLock'='{SCROLLLOCK}'; 'Shift+Alt+<'='%+<'; 'Ctrl+Banana+X'=''; 'Ctrl+Shift+Enterr'='' }
    $bad = @(); foreach ($k in $expect.Keys) { $got = RunV "KeyToOnKey" $k; if ($got -cne $expect[$k]) { $bad += "$k -> '$got' (want '$($expect[$k])')" } }
    Check "Keys: friendly -> OnKey codes" ($bad.Count -eq 0) ($bad -join '; ')
    Check "Presets: shipped presets load with no warnings" ((RunV "SettingsWarnings") -eq '') (RunV "SettingsWarnings")

    # Nexcess number cycle step 1 / Auto-Color hardcode = soft blue
    $s1.Range("S1").Value2 = 1234
    Sel $s1 "S1"; Run "CycleNumberFormat"
    Check "Nexcess: number format step 1" ($s1.Range("S1").NumberFormat -eq '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)') $s1.Range("S1").NumberFormat
    Run "AutoColorSelection"
    Check "Nexcess: hardcode soft blue" ($s1.Range("S1").Font.Color -eq 12611584) $s1.Range("S1").Font.Color

    # Switch preset: Bruce Home -> pure blue hardcodes, Bruce Input first
    [void]$xl.Run($tag + "ApplyPreset", "Bruce Home")
    Sel $s1 "S1"; Run "AutoColorSelection"
    Check "Bruce Home: hardcode pure blue" ($s1.Range("S1").Font.Color -eq 16711680) $s1.Range("S1").Font.Color
    Sel $s1 "S2"; Run "CycleInputStyle"
    Check "Bruce Home: input style 1 = warm yellow" ($s1.Range("S2").Interior.Color -eq 13431551) $s1.Range("S2").Interior.Color
    Check "Active preset remembered in settings.ini" ((Get-Content "$root\settings.ini" -Raw) -match 'ActivePreset=Bruce Home')

    # Every shipped preset loads clean (catches typos in the .ini files)
    foreach ($p in 'Nexcess', 'Bruce Home', 'Banker') {
        [void]$xl.Run($tag + "ApplyPreset", $p)
        Check "Shipped preset '$p' loads with no warnings" ((RunV "SettingsWarnings") -eq '') (RunV "SettingsWarnings")
    }
    Sel $s1 "S5"; Run "CycleNumberFormat"
    Check "Banker: number step 1 keeps en dash" ($s1.Range("S5").NumberFormat -eq ('#,##0_);(#,##0);"' + [char]0x2013 + '"_);@_)')) $s1.Range("S5").NumberFormat
    [void]$xl.Run($tag + "ApplyPreset", "Bruce Home")
    Sel $s1 "S6"; Run "CycleNumberFormat"
    Check "Bruce Home: original public-build number format" ($s1.Range("S6").NumberFormat -eq '#,##0_);(#,##0);"--";@') $s1.Range("S6").NumberFormat

    # A partial custom preset: its list replaces Nexcess's, everything else falls back
    Set-Content "$pdir\Custom.ini" -Encoding utf8 -Value "[Info]`nName=Custom`n[NumberFormats]`n1=0.00`n2=0.000`n"
    [void]$xl.Run($tag + "ApplyPreset", "Custom")
    Sel $s1 "S3"; $seq = @(); 1..3 | ForEach-Object { Run "CycleNumberFormat"; $seq += $s1.Range("S3").NumberFormat }
    Check "Custom: 2-step cycle wraps" (($seq -join '|') -eq '0.00|0.000|0.00') ($seq -join '|')
    Sel $s1 "S1"; Run "AutoColorSelection"
    Check "Custom: missing sections fall back to Nexcess" ($s1.Range("S1").Font.Color -eq 12611584) $s1.Range("S1").Font.Color

    # A broken preset: bad lines are skipped + reported, nothing crashes
    Set-Content "$pdir\Broken.ini" -Encoding utf8 -Value ("[Info]`nName=Broken`n[Fills]`n1=Bad | #GGGGGG`n2=Ok | #112233`n" +
        "[NumberFormats]`n1=0.0`"`n2=#,##0`n[Keys]`nCycleFill=Ctrl+Shift+Banana`nNotAMacro=Ctrl+Shift+J`nCycleFont=Ctrl+Shift+C`n")
    [void]$xl.Run($tag + "ApplyPreset", "Broken")
    $w = RunV "SettingsWarnings"
    Check "Broken: bad color reported"        ($w -match 'GGGGGG') $w
    Check "Broken: bad key reported"          ($w -match 'Banana') $w
    Check "Broken: unknown command reported"  ($w -match 'NotAMacro') $w
    Check "Broken: duplicate key reported"    ($w -match 'used by both') $w
    Sel $s1 "S4"; Run "CycleFill"
    Check "Broken: good fill still used"      ($s1.Range("S4").Interior.Color -eq 3351057) $s1.Range("S4").Interior.Color
    Sel $s1 "S4"; Run "CycleNumberFormat"
    Check "Broken: invalid number format skipped" ($s1.Range("S4").NumberFormat -eq '#,##0') $s1.Range("S4").NumberFormat

    [void]$xl.Run($tag + "ApplyPreset", "Nexcess")

    # --- Modules not covered above: formulas, unit tags, borders ---------------
    $s1.Range("U1").Value2 = 100; $s1.Range("V1").Value2 = 110
    Sel $s1 "W1"; Run "InsertPercentChange"
    Check "PctChange: formula + preset % format" (($s1.Range("W1").Value2 -eq 0.1) -and ($s1.Range("W1").NumberFormat -eq '0.0%;(0.0%);"—";@')) "$($s1.Range('W1').Value2) / $($s1.Range('W1').NumberFormat)"
    $s1.Range("U3").Value2 = "Revenue"
    Sel $s1 "U3"; Run "CycleUnitTag_Value_Uniform"; Run "CycleUnitTag_Value_Uniform"
    Check "UnitTag: cycles [#] -> [%]" ($s1.Range("U3").Value2 -eq "Revenue [%]") $s1.Range("U3").Value2
    Run "RemoveUnitTag"
    Check "UnitTag: removed" ($s1.Range("U3").Value2 -eq "Revenue") $s1.Range("U3").Value2
    Sel $s1 "U5"; Run "BorderBottom"
    $b1 = $s1.Range("U5").Borders(9).Weight; Run "BorderBottom"; $b2 = $s1.Range("U5").Borders(9).LineStyle; Run "BorderBottom"
    $b3 = $s1.Range("U5").Borders(9).Weight
    Check "Border cycle: thin > none > medium" (($b1 -eq 2) -and ($b2 -eq -4142) -and ($b3 -eq -4138)) "$b1 / $b2 / $b3"
    Sel $s1 "U7"; Run "ApplyInputNexcess"
    Check "ApplyInputNexcess by name" ($s1.Range("U7").Interior.Color -eq 13434879) $s1.Range("U7").Interior.Color

    # --- Ribbon galleries (the logic behind each dropdown) ----------------------
    Check "Ribbon: number gallery lists 5 steps" ((RunV "GalleryCount" "galNum") -eq 5)
    $xl.Run($tag + "GalleryLabel", "galNum", 0) | ForEach-Object { $lbl = $_ }
    Check "Ribbon: number gallery shows a formatted sample" ($lbl -match '1,234,568') $lbl
    $s1.Range("Y1").Value2 = 1234567
    Sel $s1 "Y1"; [void]$xl.Run($tag + "GalleryPick", "galNum", 2)
    Check "Ribbon: pick applies step 3" ($s1.Range("Y1").NumberFormat -eq '_(* #,##0,"K"_);_(* (#,##0,"K");_(* "-"_);_(@_)') $s1.Range("Y1").NumberFormat
    Run "CycleNumberFormat"
    Check "Ribbon: next key press continues at step 4" ($s1.Range("Y1").NumberFormat -eq '_(* #,##0.0,,_);_(* (#,##0.0,,);_(* "-"_);_(@_)') $s1.Range("Y1").NumberFormat
    Sel $s1 "Y3"; [void]$xl.Run($tag + "GalleryPick", "galInput", 1)
    Check "Ribbon: input style pick (Bruce Input fill)" ($s1.Range("Y3").Interior.Color -eq 13431551) $s1.Range("Y3").Interior.Color
    Run "UndoLastAction"
    Check "Ribbon: gallery pick is undoable" ($s1.Range("Y3").Interior.ColorIndex -eq $NONE)
    Sel $s1 "Y5"; [void]$xl.Run($tag + "GalleryPick", "galTagValue", 2)
    Check "Ribbon: tag pick" ($s1.Range("Y5").Value2 -eq '[$]') $s1.Range("Y5").Value2
    $sw = $xl.Run($tag + "GallerySwatchFile", "galFill", 1)
    Check "Ribbon: color swatch drawn" ((Test-Path $sw) -and (Get-Item $sw).Length -gt 2000) $sw
    $sw2 = $xl.Run($tag + "GallerySwatchFile", "galFill", 0)
    Check "Ribbon: 'none' swatch drawn" ($sw2 -match '_none_') $sw2
    $plabels = 0..((RunV "GalleryCount" "galPreset") - 1) | ForEach-Object { $xl.Run($tag + "GalleryLabel", "galPreset", $_) }
    Check "Ribbon: preset gallery marks only the active one" ((@($plabels | Where-Object { $_ -match 'active' }).Count -eq 1) -and
        (@($plabels | Where-Object { $_ -match '^Nexcess.*active' }).Count -eq 1)) ($plabels -join ' | ')
    Check "Ribbon: tooltip shows the preset's key" ((RunV "ScreentipFor" "CycleNumberFormat") -match 'Ctrl\+Shift\+1') (RunV "ScreentipFor" "CycleNumberFormat")

    # --- Settings window: drive every page without showing it ------------------
    $uiRoot = "$root\ui"
    [void]$xl.Run($tag + "SetSettingsRoot", $uiRoot, $true); [void]$xl.Run($tag + "LoadSettings")
    Remove-Item "$uiRoot\presets\Banker.ini"                 # room for Duplicate (max 3)
    $ui = RunV "SettingsUiSelfTest"
    Check "Settings window self-test" ($ui -eq '') $ui
    [void]$xl.Run($tag + "SetSettingsRoot", $root, $true); [void]$xl.Run($tag + "ApplyPreset", "Nexcess")

    Check "Back to Nexcess: no warnings" ((RunV "SettingsWarnings") -eq '') (RunV "SettingsWarnings")

    $wb.Close($false)
} catch {
    "ERROR: $($_.Exception.Message)"; $script:fail++
} finally {
    try { $ai.Close($false); $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()   # let Excel exit
    Stop-Job $dog -EA SilentlyContinue; Remove-Job $dog -Force -EA SilentlyContinue
    Remove-Item $root -Recurse -Force -EA SilentlyContinue
}
"`n$($script:pass) passed, $($script:fail) failed"
