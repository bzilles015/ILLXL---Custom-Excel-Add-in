# ILLXL build - assembles the add-in from src\*.bas and presets\*.ini
#
#   pwsh -File build.ps1                      -> build\ILLXL - Nexcess.xlam  (Nexcess preset first)
#   pwsh -File build.ps1 -Edition Public      -> build\ILLXL.xlam            (Bruce Home preset first)
#
# Starts from a blank workbook (no old log rows or file paths carried over), then:
#   - imports every module in src\ (UTF-8 -> Windows-1252 + CRLF for the VBA editor)
#   - adds the hidden _Undo sheet (undo engine scratch space)
#   - adds the hidden _Presets sheet (row 1 = preset name, row 2 = file text;
#     column 1 is the edition's default and the fallback for missing settings)
# Requires Excel option: Trust access to the VBA project object model.
param(
    [ValidateSet('Nexcess', 'Public')] [string]$Edition = 'Nexcess',
    [string]$Out
)
$ErrorActionPreference = 'Stop'
$cfg = @{
    Nexcess = @{ File = 'ILLXL - Nexcess.xlam'; Title = 'ILLXL - Nexcess'; Presets = 'Nexcess', 'Bruce Home', 'Banker' }
    Public  = @{ File = 'ILLXL.xlam';           Title = 'ILLXL';           Presets = 'Bruce Home', 'Banker', 'Nexcess' }
}[$Edition]
if (-not $Out) { $Out = Join-Path $PSScriptRoot "build\$($cfg.File)" }
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null

$tmp = Join-Path $env:TEMP "illxl_build"
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory $tmp | Out-Null

# UTF-8 source -> Windows-1252 + CRLF (what the VBA editor expects)
#   *.bas = standard modules   *.cls = class modules   *.vba = UserForm code
python -I -c @"
import pathlib, sys
src, dst = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
for pat in ('*.bas', '*.cls'):
    for f in src.glob(pat):
        t = f.read_text(encoding='utf-8').replace('\r\n', '\n').replace('\n', '\r\n')
        (dst / f.name).write_bytes(t.encode('cp1252'))
"@ "$PSScriptRoot\src" $tmp

$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets($wb.Worksheets.Count).Delete() }

    # --- modules ---------------------------------------------------------------
    $proj = $wb.VBProject
    foreach ($f in Get-ChildItem $tmp -Filter *.bas) { [void]$proj.VBComponents.Import($f.FullName) }
    # UserForms: created empty (they build their controls in code), then given their code.
    # Adding a form also adds the Microsoft Forms 2.0 reference the class below needs.
    foreach ($f in Get-ChildItem "$PSScriptRoot\src" -Filter *.vba) {
        $form = $proj.VBComponents.Add(3)                  # 3 = vbext_ct_MSForm
        $form.Name = $f.BaseName
        $cm = $form.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString([IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8))
    }
    foreach ($f in Get-ChildItem $tmp -Filter *.cls) { [void]$proj.VBComponents.Import($f.FullName) }

    # --- hidden sheets -----------------------------------------------------------
    $undo = $wb.Worksheets(1); $undo.Name = '_Undo'
    $ps = $wb.Worksheets.Add(); $ps.Name = '_Presets'
    for ($i = 0; $i -lt $cfg.Presets.Count; $i++) {
        $name = $cfg.Presets[$i]
        $text = [IO.File]::ReadAllText("$PSScriptRoot\presets\$name.ini", [Text.Encoding]::UTF8).Replace("`r`n", "`n")
        if ($text.Length -gt 32000) { throw "$name.ini is too long for one cell" }
        $ps.Cells(1, $i + 1).NumberFormat = '@'; $ps.Cells(1, $i + 1).Value2 = $name
        $ps.Cells(2, $i + 1).NumberFormat = '@'; $ps.Cells(2, $i + 1).Value2 = $text
    }

    # --- name shown in File > Options > Add-ins -----------------------------------
    $props = $wb.BuiltinDocumentProperties          # Office properties need late-bound calls from PowerShell
    $set = { param($n, $v) $p = [__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($n))
             [void][__ComObject].InvokeMember('Value', 'SetProperty', $null, $p, @($v)) }
    & $set 'Title'    $cfg.Title
    & $set 'Comments' '50+ keyboard shortcuts for finance pros. github.com/bzilles015'
    & $set 'Author'   'Bruce Zachary Illes'

    $wb.IsAddin = $true
    if (Test-Path $Out) { Remove-Item $Out -Force }
    $wb.SaveAs($Out, 55)                # 55 = xlOpenXMLAddIn
    $wb.Close($false)
    "Built ($Edition): $Out"
} finally {
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()   # let Excel exit
}
