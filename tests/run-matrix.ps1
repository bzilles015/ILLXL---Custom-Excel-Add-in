# ILLXL full matrix - every command x every kind of selection, plus a behaviour
# check for each macro the regression suite doesn't already cover.
#   pwsh -File tests\run-matrix.ps1 [-AddIn path]
param([string]$AddIn = "$PSScriptRoot\..\build\ILLXL - Nexcess.xlam")
$ErrorActionPreference = 'Stop'
$AddIn = (Resolve-Path $AddIn).Path
$script:pass = 0; $script:fail = 0; $script:fails = @()

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$xlPid = 0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xlPid)
$dog = Start-Job { param($p) Start-Sleep 900; Stop-Process -Id $p -Force -EA SilentlyContinue } -ArgumentList $xlPid

$ai = $xl.Workbooks.Open($AddIn); $tag = "'" + $ai.Name + "'!"
function Run($m) { [void]$xl.Run($tag + $m) }
function RunV($m, $a) {
    try { if ($PSBoundParameters.ContainsKey('a')) { $xl.Run($tag + $m, $a) } else { $xl.Run($tag + $m) } }
    catch { if ($_.Exception.Message -match "Parameter 'ptr'") { '' } else { throw } }
}
function Cmd($c) { [void]$xl.Run($tag + 'RunCmd', $c) }
function Answer($a) { [void]$xl.Run($tag + 'UiAnswer', $a) }
function Check($name, $cond, $detail = '') {
    if ($cond) { $script:pass++ } else { $script:fail++; $script:fails += "FAIL  $name  $detail"; "FAIL  $name  $detail" }
}
function Sel($ws, $addr) { $ws.Activate(); [void]$ws.Range($addr).Select() }
$root = Join-Path $env:TEMP ("illxl_mx_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
[void]$xl.Run($tag + 'SetSettingsRoot', $root, $true)
[void]$xl.Run($tag + 'UiTestMode', $true)
$NONE = -4142

# ---- fixture: one sheet of realistic mixed content ---------------------------------
function Fill-Data($ws) {
    $ws.Range("A1:H1").Value2 = [object[]]@('Item', 'Formula', 'Text', 'Date', 'Error', 'Merged', 'Blank', 'Spill')
    for ($r = 2; $r -le 20; $r++) {
        $ws.Cells($r, 1).Value2 = $r * 100; $ws.Cells($r, 2).Formula = "=A$r*2"; $ws.Cells($r, 3).Value2 = "row $r"
    }
    $ws.Range("D2").Value2 = 46000; $ws.Range("D2").NumberFormat = 'm/d/yyyy'
    $ws.Range("E2").Formula = '=1/0'; $ws.Range("F2:F3").Merge(); $ws.Range("F2").Value2 = 5
    $ws.Range("H2").Formula2 = '=SEQUENCE(3)'
}

try {
    $wb = $xl.Workbooks.Add(); $wb.Activate()
    while ($wb.Worksheets.Count -lt 6) { [void]$wb.Worksheets.Add() }
    $data = $wb.Worksheets(1); $data.Name = 'Data'; Fill-Data $data
    $prot = $wb.Worksheets(2); $prot.Name = 'Protected'; Fill-Data $prot; $prot.Protect()
    $filt = $wb.Worksheets(3); $filt.Name = 'Filtered'; Fill-Data $filt
    [void]$filt.Range("A1:H20").AutoFilter(1, ">1000")
    $tbl = $wb.Worksheets(4); $tbl.Name = 'Table'; Fill-Data $tbl
    $lo = $tbl.ListObjects.Add(1, $tbl.Range("A1:D20"), $null, 1); $lo.Name = 'tblData'
    $empty = $wb.Worksheets(5); $empty.Name = 'Empty'
    $shp = $data.Shapes.AddShape(1, 600, 300, 40, 40)
    $chart = $wb.Charts.Add(); $chart.Name = 'ChartSheet'

    $contexts = [ordered]@{
        'single cell'    = { Sel $data 'A5' }
        'mixed block'    = { Sel $data 'A2:H10' }
        'multi-area'     = { $data.Activate(); [void]$data.Range('A2:A5,C2:C5').Select() }
        'whole column'   = { Sel $data 'A:A' }
        'shape'          = { $data.Activate(); [void]$shp.Select() }
        'protected'      = { Sel $prot 'A5' }
        'filtered'       = { Sel $filt 'A1:H20' }
        'table'          = { Sel $tbl 'A2:D20' }
        'empty sheet'    = { Sel $empty 'B3' }
        'chart sheet'    = { $chart.Activate() }
    }
    $commands = @(RunV 'CommandNames')
    "crash matrix: $($commands.Count) commands x $($contexts.Count) contexts"

    # =============================== PART 1: CRASH MATRIX =============================
    foreach ($ctx in $contexts.Keys) {
        foreach ($c in $commands) {
            & $contexts[$ctx]
            $t0 = [Diagnostics.Stopwatch]::StartNew()
            try { Cmd $c } catch { Check "[$ctx] $c returns" $false $_.Exception.Message; continue }
            $ms = $t0.ElapsedMilliseconds
            if ($c -eq 'TogglePerformanceMode') { Cmd $c }                       # switch it back off
            $err = RunV 'UiLastError'
            if ($ctx -eq 'protected') {
                Check "[$ctx] ${c}: protected sheet gets a friendly message" (($err -eq '') -or ($err -match 'protected')) $err
            } else {
                Check "[$ctx] ${c}: no error" ($err -eq '') $err
            }
            Check "[$ctx] ${c}: screen updating restored" ($xl.ScreenUpdating) ''
            Check "[$ctx] ${c}: events restored" ($xl.EnableEvents) ''
            Check "[$ctx] ${c}: calc automatic" ($xl.Calculation -eq -4105) $xl.Calculation
            if ($ctx -eq 'whole column') { Check "[$ctx] ${c}: fast on a whole column" ($ms -lt 4000) "$ms ms" }
        }
    }
    "crash matrix done: $($script:pass) checks passed, $($script:fail) failed"

    # =============================== PART 2: BEHAVIOUR ================================
    $ws = $empty; $ws.Activate(); $ws.Cells.Clear()
    $s2 = $wb.Worksheets.Add(); $s2.Name = 'Other'; $ws.Activate()

    # Make refs absolute / relative (+ a formula too long to convert is skipped)
    $ws.Range("A1").Value2 = 1; $ws.Range("B1").Formula = '=A1+1'
    $long = '=A1' + ('+A1' * 90)
    $ws.Range("C1").Formula = $long
    Sel $ws 'B1:C1'; Cmd 'MakeRefsAbsolute'
    Check "MakeRefsAbsolute" ($ws.Range("B1").Formula -eq '=$A$1+1') $ws.Range("B1").Formula
    Check "MakeRefsAbsolute: long formula reported, not crashed" ((RunV 'UiShown') -match '255') ''
    Cmd 'MakeRefsRelative'
    Check "MakeRefsRelative" ($ws.Range("B1").Formula -eq '=A1+1') $ws.Range("B1").Formula

    # Next blank / next error
    $ws.Range("E1:E5").Value2 = 1; $ws.Range("E3").ClearContents(); $ws.Range("E5").Formula = '=1/0'
    Sel $ws 'E1:E5'; $ws.Range("E1").Activate(); Cmd 'GoToNextBlank'
    Check "GoToNextBlank" ($xl.ActiveCell.Address(0, 0) -eq 'E3') $xl.ActiveCell.Address(0, 0)
    Sel $ws 'E1:E5'; Cmd 'GoToNextError'
    Check "GoToNextError" ($xl.ActiveCell.Address(0, 0) -eq 'E5') $xl.ActiveCell.Address(0, 0)

    # Error checking toggles and toggles back
    $before = $xl.ErrorCheckingOptions.BackgroundChecking
    Cmd 'ToggleErrorChecking'; $mid = $xl.ErrorCheckingOptions.BackgroundChecking; Cmd 'ToggleErrorChecking'
    Check "ToggleErrorChecking" (($mid -ne $before) -and ($xl.ErrorCheckingOptions.BackgroundChecking -eq $before))

    # Decimals on a custom accounting format (both sections), and back
    $ws.Range("G1").Value2 = 1234.5; $ws.Range("G1").NumberFormat = '#,##0_);(#,##0)'
    Sel $ws 'G1'; Cmd 'IncreaseDecimal'
    Check "IncreaseDecimal" ($ws.Range("G1").NumberFormat -eq '#,##0.0_);(#,##0.0)') $ws.Range("G1").NumberFormat
    Cmd 'DecreaseDecimal'
    Check "DecreaseDecimal" ($ws.Range("G1").NumberFormat -eq '#,##0_);(#,##0)') $ws.Range("G1").NumberFormat

    # Scale down, /100, x100; sign wrap + unwrap; dynamic array kept intact
    $ws.Range("H1").Value2 = 2; Sel $ws 'H1'; Cmd 'ScaleDown'
    Check "ScaleDown" ($ws.Range("H1").Value2 -eq 2000)
    Cmd 'DivideByHundred'; Check "DivideByHundred" ($ws.Range("H1").Value2 -eq 20)
    Cmd 'MultiplyByHundred'; Check "MultiplyByHundred" ($ws.Range("H1").Value2 -eq 2000)
    $ws.Range("H2").Formula = '=H1+1'; Sel $ws 'H2'; Cmd 'ToggleSign'
    Check "ToggleSign wraps a formula" ($ws.Range("H2").Formula -eq '=-(H1+1)') $ws.Range("H2").Formula
    Cmd 'ToggleSign'; Check "ToggleSign unwraps it again" ($ws.Range("H2").Formula -eq '=H1+1') $ws.Range("H2").Formula
    $ws.Range("J1").Formula2 = '=SEQUENCE(3)'; Sel $ws 'J1'; Cmd 'ScaleDown'
    Check "ScaleDown keeps a dynamic array spilling" (($ws.Range("J1").Formula2 -eq '=(SEQUENCE(3))*1000') -and ($ws.Range("J3").Value2 -eq 3000)) $ws.Range("J1").Formula2

    # CAGR (scripted picks), incl. a cell on another sheet
    $ws.Range("A3").Value2 = 100; $s2.Range("B2").Value2 = 200
    Sel $ws 'A5'; Answer 'A3'; Answer 'Other!B2'; Answer 3; Cmd 'InsertCAGR'
    Check "CAGR formula, cross-sheet end cell" (($ws.Range("A5").Formula -replace "'","") -eq "=IFERROR((Other!B2/A3)^(1/3)-1,0)") $ws.Range("A5").Formula
    Check "CAGR value" ([math]::Abs($ws.Range("A5").Value2 - 0.259921) -lt 0.0001) $ws.Range("A5").Value2
    Sel $ws 'A6'; Answer 'A3'; Answer 'cancel'; Cmd 'InsertCAGR'
    Check "CAGR cancel leaves the cell alone" ($ws.Range("A6").Formula -eq '') $ws.Range("A6").Formula
    Sel $ws 'A6'; Answer 'A3'; Answer 'A3'; Answer 2.5; Cmd 'InsertCAGR'
    Check "CAGR periods written with '.' decimals" ($ws.Range("A6").Formula -match '\^\(1/2\.5\)') $ws.Range("A6").Formula

    # Growth rate: typed number, then a cell on another sheet (locked)
    $ws.Range("K1").Value2 = 100; Sel $ws 'L1:M1'; Answer 0.1; Cmd 'ApplyGrowthRate'
    Check "Growth rate, typed" (($ws.Range("L1").Formula -eq '=K1*(1+0.1)') -and ([math]::Round($ws.Range("M1").Value2, 6) -eq 121)) $ws.Range("L1").Formula
    $s2.Range("C3").Value2 = 0.05; Sel $ws 'L1'; Answer 'Other!C3'; Cmd 'ApplyGrowthRate'
    Check "Growth rate, cell on another sheet" (($ws.Range("L1").Formula -replace "'","") -eq "=K1*(1+Other!`$C`$3)") $ws.Range("L1").Formula

    # Quick SUM / AVERAGE, cross-sheet
    $s2.Range("D1:D3").Value2 = 2
    Sel $ws 'N1'; Answer 'Other!D1:D3'; Cmd 'InsertQuickSum'
    Check "Quick SUM cross-sheet" ((($ws.Range("N1").Formula -replace "'","") -eq "=SUM(Other!D1:D3)") -and ($ws.Range("N1").Value2 -eq 6)) $ws.Range("N1").Formula
    Sel $ws 'N2'; Answer 'Other!D1:D3'; Cmd 'InsertQuickAverage'
    Check "Quick AVERAGE" ($ws.Range("N2").Value2 -eq 2) $ws.Range("N2").Formula
    Sel $ws 'N1:N2'; Cmd 'InsertQuickSum'
    Check "Quick SUM needs one cell (message, no change)" ((RunV 'UiShown') -match 'single cell')

    # Equals left
    $ws.Range("O1").Value2 = 7; Sel $ws 'P1:Q1'; Cmd 'EqualsLeft'
    Check "EqualsLeft" (($ws.Range("P1").Formula -eq '=O1') -and ($ws.Range("Q1").Formula -eq '=P1'))

    # Font cycle, text case cycle
    $ws.Range("R1").Value2 = 'hello world'; Sel $ws 'R1'
    Cmd 'CycleFont'; $f1 = $ws.Range("R1").Font.Name; Cmd 'CycleFont'; $f2 = $ws.Range("R1").Font.Name
    Check "CycleFont steps through the preset" (($f1 -eq 'Arial') -and ($f2 -eq 'Aptos Narrow')) "$f1 / $f2"
    Cmd 'CycleTextCase'; $c1 = $ws.Range("R1").Value2; Cmd 'CycleTextCase'; $c2 = $ws.Range("R1").Value2; Cmd 'CycleTextCase'
    Check "CycleTextCase Title -> lower -> UPPER" (($c1 -eq 'Hello World') -and ($c2 -eq 'hello world') -and ($ws.Range("R1").Value2 -eq 'HELLO WORLD'))

    # Indent in/out, center across
    $ws.Range("S1").HorizontalAlignment = -4108; Sel $ws 'S1'; Cmd 'IndentIn'
    Check "IndentIn left-aligns and indents" (($ws.Range("S1").IndentLevel -eq 1) -and ($ws.Range("S1").HorizontalAlignment -eq -4131))
    Cmd 'IndentOut'; Check "IndentOut" ($ws.Range("S1").IndentLevel -eq 0)
    Sel $ws 'S2:U2'; Cmd 'CenterAcrossSelection'
    Check "Center across selection (no merge)" (($ws.Range("S2").HorizontalAlignment -eq 7) -and -not $ws.Range("S2").MergeCells)

    # Timestamp (value, not formula; format cycles)
    Sel $ws 'V1'; Cmd 'InsertStaticNow'
    Check "Timestamp is a value" ((-not $ws.Range("V1").HasFormula) -and ($ws.Range("V1").Value2 -gt 46000))
    Cmd 'InsertStaticNow'; Check "Timestamp 2nd press = EST format" ($ws.Range("V1").NumberFormat -match 'EST') $ws.Range("V1").NumberFormat

    # Insert headers (scripted answer)
    Sel $ws 'A10'; Answer 'FY24, FY25 ,FY26'; Cmd 'InsertHeadersFromPrompt'
    Check "Insert headers fill across, trimmed, bold" (($ws.Range("C10").Value2 -eq 'FY26') -and ($ws.Range("B10").Value2 -eq 'FY25') -and $ws.Range("A10").Font.Bold)
    Sel $ws 'A11'; Answer 'cancel'; Cmd 'InsertHeadersFromPrompt'
    Check "Insert headers cancel = no change" ($ws.Range("A11").Value2 -eq $null)

    # Zero check apply / clear
    $ws.Range("A12").Value2 = 0; Sel $ws 'A12'; Cmd 'ApplyZeroCheckCF'
    Check "Zero-check adds 2 rules" ($ws.Range("A12").FormatConditions.Count -eq 2) $ws.Range("A12").FormatConditions.Count
    Cmd 'ApplyZeroCheckCF'; Check "Zero-check again doesn't stack rules" ($ws.Range("A12").FormatConditions.Count -eq 2)
    Cmd 'ClearZeroCheckCF'; Check "Zero-check cleared" ($ws.Range("A12").FormatConditions.Count -eq 0)

    # Borders
    Sel $ws 'B14:D16'; Cmd 'BordersOutlineInside'
    Check "Outline medium + inside thin" (($ws.Range("B14").Borders(8).Weight -eq -4138) -and ($ws.Range("C15").Borders(12).Weight -eq 2))
    Sel $ws 'B18:D18'; Cmd 'ApplySumBar'
    Check "Sum bar = thin top" (($ws.Range("C18").Borders(8).LineStyle -eq 1) -and ($ws.Range("C18").Borders(8).Weight -eq 2))
    Sel $ws 'B19:D19'; Cmd 'ApplyDoubleSumBar'
    Check "Double sum bar = double bottom" ($ws.Range("C19").Borders(9).LineStyle -eq -4119) $ws.Range("C19").Borders(9).LineStyle
    foreach ($e in @(@('BorderTop', 8), @('BorderLeft', 7), @('BorderRight', 10))) {
        Sel $ws 'F21'; Cmd $e[0]
        Check "$($e[0]) step 1 = thin" ($ws.Range("F21").Borders($e[1]).Weight -eq 2)
    }

    # Unit tags: duration + rate cycles
    $ws.Range("A23").Value2 = 'Term'; Sel $ws 'A23'; Cmd 'CycleUnitTag_Duration_Uniform'; Cmd 'CycleUnitTag_Duration_Uniform'
    Check "Duration tag cycles [d] -> [m]" ($ws.Range("A23").Value2 -eq 'Term [m]') $ws.Range("A23").Value2
    Sel $ws 'A24'; Cmd 'CycleUnitTag_Rate_Uniform'
    Check "Rate tag into an empty cell" ($ws.Range("A24").Value2 -eq '[%/y]') $ws.Range("A24").Value2

    # Tab color cycle through RunCmd
    Cmd 'CycleTabColor'; Cmd 'CycleTabColor'
    Check "Tab color step 2 = blue" ($ws.Tab.Color -eq 12874308) $ws.Tab.Color

    # Direct input styles (all six, by name, Nexcess preset)
    $named = @{ ApplyInputNexcess = 13434879; ApplyInputBruce = 13431551; ApplyInputBanker = 13434879; ApplyInputGray = 14277081; ApplyInputPeach = 10079487; ApplyInputOverride = 16247773 }
    $row = 26
    foreach ($k in $named.Keys) { Sel $ws "A$row"; Cmd $k; Check "$k" ($ws.Range("A$row").Interior.Color -eq $named[$k]) $ws.Range("A$row").Interior.Color; $row++ }

    # Clean up styles: unused custom style goes, used one stays
    [void]$wb.Styles.Add('ILLXL Junk'); [void]$wb.Styles.Add('ILLXL Used'); $ws.Range("A40").Style = 'ILLXL Used'
    [void]$xl.Run($tag + 'UiTestMode', $true)                 # fresh answer queue
    Cmd 'RemoveUnusedStyles'
    Check "Clean Up Styles explains a protected sheet first" ((RunV 'UiShown') -match 'protected') (RunV 'UiShown')
    $prot.Unprotect(); [void]$xl.Run($tag + 'UiTestMode', $true)
    Cmd 'RemoveUnusedStyles'; $prot.Protect()
    $why = (RunV 'UiShown') + ' / ' + (RunV 'UiLastError')
    $names = @($wb.Styles | ForEach-Object Name)
    Check "Clean Up Styles removes the unused one" (-not ($names -contains 'ILLXL Junk')) $why
    Check "Clean Up Styles keeps the used one" ($names -contains 'ILLXL Used')

    # Huge selection: a 1,048,576-cell write asks first, and Cancel means nothing is written
    Sel $ws 'AA:AA'; Cmd 'CycleUnitTag_Value_Uniform'
    Check "Whole-column tag fills only the used rows" (($xl.WorksheetFunction.CountA($ws.Range('AA:AA')) -le $ws.UsedRange.Rows.Count) -and ($xl.WorksheetFunction.CountA($ws.Range('AA:AA')) -gt 0))
    $ws.Range("AC1:AC60000").Value2 = 1; Sel $ws 'AD1:AD60000'; Answer 'cancel'; Cmd 'InsertStaticNow'
    Check "60,000-cell write asks first; Cancel writes nothing" (($xl.WorksheetFunction.CountA($ws.Range('AD:AD')) -eq 0) -and ((RunV 'UiShown') -match '60,000'))

    # Friendly errors: protected sheet, unknown command
    Sel $prot 'A2'; Cmd 'CycleFill'
    Check "Protected sheet: friendly message" ((RunV 'UiLastError') -match 'protected') (RunV 'UiLastError')
    Cmd 'NoSuchCommand'
    Check "Unknown command: reported, not crashed" ((RunV 'UiLastError') -match 'not an ILLXL command')

    # Every catalog command has a description (ribbon tooltip + PDF)
    $missing = @($commands | Where-Object { ($xl.Run($tag + 'ScreentipFor', $_)) -eq '' })
    Check "Every command has a label" ($missing.Count -eq 0) ($missing -join ', ')

    $wb.Close($false)
} catch {
    "ERROR: $($_.Exception.Message)"; $script:fail++
} finally {
    try { $ai.Close($false); $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Stop-Job $dog -EA SilentlyContinue; Remove-Job $dog -Force -EA SilentlyContinue
    Remove-Item $root -Recurse -Force -EA SilentlyContinue
}
"`n$($script:pass) passed, $($script:fail) failed"
