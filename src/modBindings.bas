Attribute VB_Name = "modBindings"

Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Module: modBindings
' Purpose:
'   Auto-load and bind all keyboard shortcuts for the add-in.
'   The keys come from the active preset ([Keys] in the .ini - see
'   modSettings). This module holds the command catalog - every macro a key
'   can run, its group, and a plain-English label for the Settings window.
'
'   Every key ILLXL takes is remembered, so it can be handed back to Excel
'   when the preset changes or the add-in unloads.
'
' OnKey codes:  ^ = Ctrl   % = Alt   + = Shift   {KEY} = special key
'==============================================================================

Private BoundKeys As Collection                 ' OnKey codes ILLXL currently owns
Private Catalog   As Object                     ' macro -> "Group|Label"

'------------------------------------------------------------------------------
' Auto_Open - Runs when the add-in loads
'------------------------------------------------------------------------------
Sub Auto_Open()
    LoadSettings
    BindAllKeys
    ' quiet once-a-day update check, after Excel has finished starting
    Application.OnTime Now + TimeSerial(0, 0, 15), "'" & ThisWorkbook.Name & "'!StartupUpdateCheck"
End Sub

'------------------------------------------------------------------------------
' BindAllKeys - releases old keys, then binds the active preset's keys
'------------------------------------------------------------------------------
Sub BindAllKeys()
    Dim item As Variant
    UnbindAllKeys
    Set BoundKeys = PresetKeyCodes()
    On Error Resume Next
    For Each item In BoundKeys
        If Len(item(1)) = 0 Then
            Application.OnKey item(0), ""                                   ' nuisance key: off
        Else
            Application.OnKey item(0), KeyProcedure(item(1))
        End If
    Next item
    On Error GoTo 0
End Sub

' Every key runs its command through RunCmd (modUi), so an error shows a
' friendly message instead of the VBA Debug dialog. Format (verified):
'   'ILLXL.xlam'!'RunCmd "CycleFill"'
Public Function KeyProcedure(ByVal macro As String) As String
    KeyProcedure = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!'RunCmd """ & macro & """'"
End Function

'------------------------------------------------------------------------------
' UnbindAllKeys - every key ILLXL took goes back to Excel's default
'     (runs from Auto_Close and before a preset switch). If VBA was reset
'     and the list is gone, the active preset's keys are released instead.
'------------------------------------------------------------------------------
Sub UnbindAllKeys()
    Dim item As Variant
    If BoundKeys Is Nothing Then Set BoundKeys = PresetKeyCodes()
    On Error Resume Next
    For Each item In BoundKeys
        Application.OnKey item(0)
    Next item
    On Error GoTo 0
    Set BoundKeys = Nothing
End Sub

'------------------------------------------------------------------------------
' PresetKeyCodes - the active preset's keys as Array(OnKey code, macro)
'     macro "" = nuisance key to switch off. Unknown macros and keys that
'     can't be read are skipped.
'------------------------------------------------------------------------------
Private Function PresetKeyCodes() As Collection
    Dim out As New Collection, keys As Object, macro As Variant, code As String, k As Variant
    EnsureSettings
    For Each k In DisabledKeyList()
        out.Add Array(k, "")
    Next k
    Set keys = KeyAssignments()
    For Each macro In keys.keys
        code = KeyToOnKey(keys(macro))
        If Len(code) > 0 And IsCommand(macro) Then out.Add Array(code, macro)
    Next macro
    Set PresetKeyCodes = out
End Function

'------------------------------------------------------------------------------
' Command catalog - IsCommand / CommandInfo
'------------------------------------------------------------------------------
Public Function IsCommand(ByVal macro As String) As Boolean
    LoadCatalog
    IsCommand = Catalog.Exists(macro)
End Function

' Returns Array(group, label, description) - for the Settings window,
' ribbon tooltips and the shortcut PDF (tools/make-pdf.py reads this list)
Public Function CommandInfo(ByVal macro As String) As Variant
    LoadCatalog
    If Catalog.Exists(macro) Then CommandInfo = Split(Catalog(macro), "|") Else CommandInfo = Array("", macro, "")
End Function

Public Function CommandNames() As Variant
    LoadCatalog
    CommandNames = Catalog.keys
End Function

Private Sub LoadCatalog()
    If Not Catalog Is Nothing Then Exit Sub
    Set Catalog = CreateObject("Scripting.Dictionary")
    Catalog.CompareMode = vbTextCompare

    '==== modSettings / modCore (Cxx) =========================================
    Cmd "ShowSettings", "Core", "ILLXL Settings / Presets", _
        "Open the Settings window: presets, cycles, colors, fonts, tags and shortcut keys."
    Cmd "TogglePerformanceMode", "Core", "Toggle Performance Mode", _
        "Manual calc, events and conditional formatting off for heavy formatting. Press again to switch back. Undo is off while it is on."
    Cmd "MakeRefsAbsolute", "Core", "Make References Absolute", _
        "Turn A1 into $A$1 in every formula in the selection."
    Cmd "MakeRefsRelative", "Core", "Make References Relative", _
        "Turn $A$1 into A1 in every formula in the selection."
    Cmd "GoToNextBlank", "Core", "Go To Next Blank Cell", _
        "Jump to the next empty cell in the selection."
    Cmd "GoToNextError", "Core", "Go To Next Error Cell", _
        "Jump to the next #REF!, #N/A, #DIV/0! ... in the selection."
    Cmd "ToggleErrorChecking", "Core", "Toggle Error Checking", _
        "Hide or show Excel's green error triangles - for presentation-ready sheets."
    Cmd "RemoveUnusedStyles", "Core", "Remove Unused Cell Styles", _
        "Delete custom cell styles that no cell uses (asks first)."

    '==== modFormatCycles (Fxx) ===============================================
    Cmd "CycleNumberFormat", "Formats", "Cycle Number Format", _
        "Step through the number formats in your preset (e.g. 1,234 -> 1K -> 1.2M)."
    Cmd "CycleDateFormat", "Formats", "Cycle Date Format", _
        "Step through the date formats in your preset."
    Cmd "CycleCurrencyFormat", "Formats", "Cycle Currency Format", _
        "Step through the currency formats in your preset."
    Cmd "CyclePercentFormat", "Formats", "Cycle Percent Format", _
        "Step through the percent formats in your preset."
    Cmd "CycleOtherNumbers", "Formats", "Cycle Other Formats (A/B/F/Q/x)", _
        "Year tags (2026A / B / F / E), quarters and multiples (12.5x)."
    Cmd "IncreaseDecimal", "Formats", "Increase Decimals", _
        "One more decimal place - works on custom formats too."
    Cmd "DecreaseDecimal", "Formats", "Decrease Decimals", _
        "One less decimal place - works on custom formats too."
    Cmd "ScaleUp", "Formats", "Divide by 1,000", _
        "Divide numbers by 1,000; formulas become =(...)/1000."
    Cmd "ScaleDown", "Formats", "Multiply by 1,000", _
        "Multiply numbers by 1,000; formulas become =(...)*1000."
    Cmd "ToggleSign", "Formats", "Toggle Sign", _
        "Flip positive <-> negative; formulas are wrapped as =-(...) and unwrapped again."
    Cmd "DivideByHundred", "Formats", "Divide by 100", _
        "Divide by 100 (5 -> 0.05) for entering percentages."
    Cmd "MultiplyByHundred", "Formats", "Multiply by 100", _
        "Multiply by 100 (0.05 -> 5)."

    '==== modFormulas (Mxx) ===================================================
    Cmd "InsertCAGR", "Formulas", "Insert CAGR", _
        "=(End/Begin)^(1/n)-1. Pick the two cells; n = end year minus start year."
    Cmd "InsertPercentChange", "Formulas", "Insert % Change", _
        "=(Current-Prior)/ABS(Prior), using the two cells to the left."
    Cmd "EqualsLeft", "Formulas", "Equals Left", _
        "Link each selected cell to the cell on its left."
    Cmd "ApplyGrowthRate", "Formulas", "Apply Growth Rate", _
        "=Left*(1+rate) across the selection. Type a rate or click the cell that holds it."
    Cmd "InsertQuickSum", "Formulas", "Quick SUM", _
        "Pick a range (any sheet) and SUM it into the selected cell."
    Cmd "InsertQuickAverage", "Formulas", "Quick AVERAGE", _
        "Pick a range (any sheet) and AVERAGE it into the selected cell."

    '==== modStyles (Sxx) =====================================================
    Cmd "AutoColorSelection", "Styles", "Auto-Color (inputs / formulas / links)", _
        "Color cells by what they hold: hardcodes, formulas, same-sheet, other-sheet and other-file links (colors from your preset)."
    Cmd "CycleFont", "Styles", "Cycle Font", _
        "Step through the fonts in your preset."
    Cmd "CycleFill", "Styles", "Cycle Fill", _
        "Step through the fill colors in your preset."
    Cmd "CycleTextCase", "Styles", "Cycle Text Case", _
        "Title Case -> lower case -> UPPER CASE (typed text only, never formulas)."
    Cmd "CycleFontColor", "Styles", "Cycle Font Color", _
        "Step through the font colors in your preset."
    Cmd "IncreaseFontSize", "Styles", "Increase Font Size", _
        "Font one point bigger (mixed sizes each step from their own size)."
    Cmd "DecreaseFontSize", "Styles", "Decrease Font Size", _
        "Font one point smaller."
    Cmd "IndentIn", "Styles", "Indent In", _
        "Indent one level (left-aligns so the indent shows)."
    Cmd "IndentOut", "Styles", "Indent Out", _
        "Remove one indent level."
    Cmd "CenterAcrossSelection", "Styles", "Center Across Selection", _
        "Center across the selected cells without merging them."
    Cmd "InsertStaticNow", "Styles", "Insert Timestamp", _
        "Stamp the current date and time as a value; press again for the next format."
    Cmd "CycleInputStyle", "Styles", "Cycle Input Style", _
        "Step through the input styles in your preset (fill, font color, dotted border)."
    Cmd "CycleHeaderStyle", "Styles", "Cycle Header Style", _
        "Step through the header styles in your preset (fill, font color, bold)."
    Cmd "InsertHeadersFromPrompt", "Styles", "Insert Headers", _
        "Type headers like 2024A,2025B,2026E and they fill across, styled."
    Cmd "InsertVarianceHeaders", "Styles", "Insert Variance Headers", _
        "Insert the variance header labels from your preset, styled."
    Cmd "ApplyZeroCheckCF", "Styles", "Apply Zero-Check", _
        "Green when the cell is 0, red when it isn't - for balance checks."
    Cmd "ClearZeroCheckCF", "Styles", "Clear Zero-Check", _
        "Remove ILLXL's zero-check formatting from the selection."
    Cmd "CycleTabColor", "Styles", "Cycle Tab Color", _
        "Step the sheet tab through the tab colors in your preset."

    '==== modBorders (Bxx) ====================================================
    Cmd "BorderTop", "Borders", "Cycle Top Border", _
        "Step the top edge through your border cycle (e.g. thin -> none -> medium -> hairline)."
    Cmd "BorderBottom", "Borders", "Cycle Bottom Border", _
        "Step the bottom edge through your border cycle."
    Cmd "BorderLeft", "Borders", "Cycle Left Border", _
        "Step the left edge through your border cycle."
    Cmd "BorderRight", "Borders", "Cycle Right Border", _
        "Step the right edge through your border cycle."
    Cmd "BordersOutlineInside", "Borders", "Outline + Inside Borders", _
        "Medium outline with thin inside lines."
    Cmd "ApplySumBar", "Borders", "Sum Bar", _
        "Thin line above the totals row."
    Cmd "ApplyDoubleSumBar", "Borders", "Double Sum Bar", _
        "Thin line above and double line below - for grand totals."

    '==== modUnitTags (Uxx) ===================================================
    Cmd "CycleUnitTag_Value_Uniform", "Unit Tags", "Cycle Value Tag", _
        "Add or cycle a value tag: [#] [%] [$] [mln $] ... (FAST standard)."
    Cmd "CycleUnitTag_Duration_Uniform", "Unit Tags", "Cycle Duration Tag", _
        "Add or cycle a duration tag: [d] [m] [q] [y]."
    Cmd "CycleUnitTag_Rate_Uniform", "Unit Tags", "Cycle Rate Tag", _
        "Add or cycle a rate tag: [%/y] [$/unit] [$/FTE] [$/yr]."
    Cmd "RemoveUnitTag", "Unit Tags", "Remove Unit Tag", _
        "Remove the last [...] tag from each selected cell."

    '==== modTools (Txx) =====================================================
    Cmd "AuditHardcodes", "Model Tools", "Find Hardcodes in Formulas", _
        "Select every formula with a typed-in number (=A1*1.05). Works on the selection, or the whole sheet when one cell is selected."
    Cmd "AuditInconsistent", "Model Tools", "Find Inconsistent Formulas", _
        "Select formulas that break the pattern of the formulas beside them in their row."
    Cmd "AuditExternalLinks", "Model Tools", "External Link Report", _
        "List every cell and name that links to another file, each with a link to jump to it."
    Cmd "SendReady", "Model Tools", "Ready to Send", _
        "Every sheet back to its first cell below the freeze panes (or A1), first sheet shown - before you send a model."
    Cmd "ToggleIfError", "Model Tools", "Wrap / Unwrap IFERROR", _
        "Wrap the selected formulas in IFERROR(...,0); press again to take it off. The fallback value is in your preset."
    Cmd "ToggleRound", "Model Tools", "Wrap / Unwrap ROUND", _
        "Wrap the selected formulas in ROUND(...,0); press again to take it off. Digits are in your preset."
    Cmd "BuildSheetIndex", "Model Tools", "Sheet Index", _
        "Add (or refresh) a hyperlinked table of contents at the front of the workbook."
    Cmd "CheckForUpdates", "Core", "Check for Updates", _
        "See whether a newer ILLXL is out on GitHub."
End Sub

Private Sub Cmd(ByVal macro As String, ByVal group As String, ByVal label As String, ByVal desc As String)
    Catalog(macro) = group & "|" & label & "|" & desc
End Sub

'------------------------------------------------------------------------------
' Performance Mode Quick Reference
'
' Turn ON before heavy actions:
'   - Mass formatting, filling/copying big ranges
'   - Applying/removing lots of CF, duplicating sheets
'
' It sets Manual calc, turns off screen updates & events, skips undo/logging.
'
' Turn OFF for review and normal use:
'   - Automatic calc, Ctrl+Z history, logging
'------------------------------------------------------------------------------

