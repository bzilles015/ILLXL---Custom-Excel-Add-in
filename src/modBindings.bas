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
            Application.OnKey item(0), "'" & ThisWorkbook.Name & "'!" & item(1)
        End If
    Next item
    On Error GoTo 0
End Sub

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

' Returns Array(group, label) - for the Settings window
Public Function CommandInfo(ByVal macro As String) As Variant
    LoadCatalog
    If Catalog.Exists(macro) Then CommandInfo = Split(Catalog(macro), "|") Else CommandInfo = Array("", macro)
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
    Cmd "ShowSettings", "Core", "ILLXL Settings / Presets"
    Cmd "TogglePerformanceMode", "Core", "Toggle Performance Mode"
    Cmd "MakeRefsAbsolute", "Core", "Make References Absolute"
    Cmd "MakeRefsRelative", "Core", "Make References Relative"
    Cmd "GoToNextBlank", "Core", "Go To Next Blank Cell"
    Cmd "GoToNextError", "Core", "Go To Next Error Cell"
    Cmd "ToggleErrorChecking", "Core", "Toggle Error Checking"
    Cmd "RemoveUnusedStyles", "Core", "Remove Unused Cell Styles"

    '==== modFormatCycles (Fxx) ===============================================
    Cmd "CycleNumberFormat", "Formats", "Cycle Number Format"
    Cmd "CycleDateFormat", "Formats", "Cycle Date Format"
    Cmd "CycleCurrencyFormat", "Formats", "Cycle Currency Format"
    Cmd "CyclePercentFormat", "Formats", "Cycle Percent Format"
    Cmd "CycleOtherNumbers", "Formats", "Cycle Other Formats (A/B/F/Q/x)"
    Cmd "IncreaseDecimal", "Formats", "Increase Decimals"
    Cmd "DecreaseDecimal", "Formats", "Decrease Decimals"
    Cmd "ScaleUp", "Formats", "Divide by 1,000"
    Cmd "ScaleDown", "Formats", "Multiply by 1,000"
    Cmd "ToggleSign", "Formats", "Toggle Sign"
    Cmd "DivideByHundred", "Formats", "Divide by 100"
    Cmd "MultiplyByHundred", "Formats", "Multiply by 100"

    '==== modFormulas (Mxx) ===================================================
    Cmd "InsertCAGR", "Formulas", "Insert CAGR"
    Cmd "InsertPercentChange", "Formulas", "Insert % Change"
    Cmd "EqualsLeft", "Formulas", "Equals Left"
    Cmd "ApplyGrowthRate", "Formulas", "Apply Growth Rate"
    Cmd "InsertQuickSum", "Formulas", "Quick SUM"
    Cmd "InsertQuickAverage", "Formulas", "Quick AVERAGE"

    '==== modStyles (Sxx) =====================================================
    Cmd "AutoColorSelection", "Styles", "Auto-Color (inputs / formulas / links)"
    Cmd "CycleFont", "Styles", "Cycle Font"
    Cmd "CycleFill", "Styles", "Cycle Fill"
    Cmd "CycleTextCase", "Styles", "Cycle Text Case"
    Cmd "CycleFontColor", "Styles", "Cycle Font Color"
    Cmd "IncreaseFontSize", "Styles", "Increase Font Size"
    Cmd "DecreaseFontSize", "Styles", "Decrease Font Size"
    Cmd "IndentIn", "Styles", "Indent In"
    Cmd "IndentOut", "Styles", "Indent Out"
    Cmd "CenterAcrossSelection", "Styles", "Center Across Selection"
    Cmd "InsertStaticNow", "Styles", "Insert Timestamp"
    Cmd "CycleInputStyle", "Styles", "Cycle Input Style"
    Cmd "CycleHeaderStyle", "Styles", "Cycle Header Style"
    Cmd "InsertHeadersFromPrompt", "Styles", "Insert Headers"
    Cmd "InsertVarianceHeaders", "Styles", "Insert Variance Headers"
    Cmd "ApplyZeroCheckCF", "Styles", "Apply Zero-Check"
    Cmd "ClearZeroCheckCF", "Styles", "Clear Zero-Check"
    Cmd "CycleTabColor", "Styles", "Cycle Tab Color"

    '==== modBorders (Bxx) ====================================================
    Cmd "BorderTop", "Borders", "Cycle Top Border"
    Cmd "BorderBottom", "Borders", "Cycle Bottom Border"
    Cmd "BorderLeft", "Borders", "Cycle Left Border"
    Cmd "BorderRight", "Borders", "Cycle Right Border"
    Cmd "BordersOutlineInside", "Borders", "Outline + Inside Borders"
    Cmd "ApplySumBar", "Borders", "Sum Bar"
    Cmd "ApplyDoubleSumBar", "Borders", "Double Sum Bar"

    '==== modUnitTags (Uxx) ===================================================
    Cmd "CycleUnitTag_Value_Uniform", "Unit Tags", "Cycle Value Tag"
    Cmd "CycleUnitTag_Duration_Uniform", "Unit Tags", "Cycle Duration Tag"
    Cmd "CycleUnitTag_Rate_Uniform", "Unit Tags", "Cycle Rate Tag"
    Cmd "RemoveUnitTag", "Unit Tags", "Remove Unit Tag"
End Sub

Private Sub Cmd(ByVal macro As String, ByVal group As String, ByVal label As String)
    Catalog(macro) = group & "|" & label
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

