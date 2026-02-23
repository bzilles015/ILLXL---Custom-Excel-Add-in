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
'   Keep bindings grouped by module so maintenance is painless.
'==============================================================================

'------------------------------------------------------------------------------
' Auto_Open - Runs when the add-in loads
'------------------------------------------------------------------------------
Sub Auto_Open()
    ResetCycleState
    BindAllKeys
End Sub

'------------------------------------------------------------------------------
' Workbook_Activate - Reassert mappings when the add-in regains focus
'------------------------------------------------------------------------------
Private Sub Workbook_Activate()
    BindAllKeys
End Sub

'------------------------------------------------------------------------------
' BindAllKeys - Binds every shortcut (and disables nuisance keys)
'------------------------------------------------------------------------------
Sub BindAllKeys()
    On Error Resume Next

    '==== Disable nuisance keys ===============================================
    Application.OnKey "{F1}", ""
    Application.OnKey "{SCROLLLOCK}", ""
    Application.OnKey "{NUMLOCK}", ""
    Application.OnKey "{INSERT}", ""

    '==== modCore (Cxx) - performance + helpers ===============================
    Application.OnKey "^%+M", "'" & ThisWorkbook.Name & "'!TogglePerformanceMode"
    Application.OnKey "^%+A", "'" & ThisWorkbook.Name & "'!MakeRefsAbsolute"
    Application.OnKey "^%+R", "'" & ThisWorkbook.Name & "'!MakeRefsRelative"
    Application.OnKey "^%+N", "'" & ThisWorkbook.Name & "'!GoToNextBlank"
    Application.OnKey "^%+E", "'" & ThisWorkbook.Name & "'!GoToNextError"
    Application.OnKey "^%+L", "'" & ThisWorkbook.Name & "'!BreakExternalLinksInSelection"

    '==== modFormatCycles (Fxx) ===============================================
    Application.OnKey "^+1", "'" & ThisWorkbook.Name & "'!CycleNumberFormat"
    Application.OnKey "^+3", "'" & ThisWorkbook.Name & "'!CycleDateFormat"
    Application.OnKey "^+4", "'" & ThisWorkbook.Name & "'!CycleCurrencyFormat"
    Application.OnKey "^+5", "'" & ThisWorkbook.Name & "'!CyclePercentFormat"
    Application.OnKey "^+8", "'" & ThisWorkbook.Name & "'!CycleOtherNumbers"
    Application.OnKey "^+.", "'" & ThisWorkbook.Name & "'!IncreaseDecimal"
    Application.OnKey "^+,", "'" & ThisWorkbook.Name & "'!DecreaseDecimal"
    Application.OnKey "+%<", "'" & ThisWorkbook.Name & "'!ScaleUp"
    Application.OnKey "+%>", "'" & ThisWorkbook.Name & "'!ScaleDown"
    Application.OnKey "^%+\", "'" & ThisWorkbook.Name & "'!ToggleSign"
    Application.OnKey "^%2", "'" & ThisWorkbook.Name & "'!DivideByHundred"
    Application.OnKey "^%+2", "'" & ThisWorkbook.Name & "'!MultiplyByHundred"

    '==== modFormulas (Mxx) - formula insertion tools =========================
    Application.OnKey "^%+C", "'" & ThisWorkbook.Name & "'!InsertCAGR"
    Application.OnKey "^%+W", "'" & ThisWorkbook.Name & "'!InsertPercentChange"
    Application.OnKey "^%l", "'" & ThisWorkbook.Name & "'!EqualsLeft"
    Application.OnKey "^%+G", "'" & ThisWorkbook.Name & "'!ApplyGrowthRate"
    Application.OnKey "^%{=}", "'" & ThisWorkbook.Name & "'!InsertQuickSum"
    Application.OnKey "^%+{=}", "'" & ThisWorkbook.Name & "'!InsertQuickAverage"

    '==== modStyles (Sxx) - colors/styles/layout/CF ===========================
    Application.OnKey "^%a", "'" & ThisWorkbook.Name & "'!AutoColorSelection"
    Application.OnKey "^'", "'" & ThisWorkbook.Name & "'!CycleFont"
    Application.OnKey "^+K", "'" & ThisWorkbook.Name & "'!CycleFill"
    Application.OnKey "^%+I", "'" & ThisWorkbook.Name & "'!CycleTextCase"
    Application.OnKey "^+C", "'" & ThisWorkbook.Name & "'!CycleFontColor"
    Application.OnKey "^+F", "'" & ThisWorkbook.Name & "'!IncreaseFontSize"
    Application.OnKey "^+G", "'" & ThisWorkbook.Name & "'!DecreaseFontSize"
    Application.OnKey "^%.", "'" & ThisWorkbook.Name & "'!IndentIn"
    Application.OnKey "^%,", "'" & ThisWorkbook.Name & "'!IndentOut"
    Application.OnKey "^%e", "'" & ThisWorkbook.Name & "'!CenterAcrossSelection"
    Application.OnKey "^+N", "'" & ThisWorkbook.Name & "'!InsertStaticNow"
    Application.OnKey "^%+U", "'" & ThisWorkbook.Name & "'!CycleInputStyle"
    Application.OnKey "^%+H", "'" & ThisWorkbook.Name & "'!CycleHeaderStyle"
    Application.OnKey "^%+Y", "'" & ThisWorkbook.Name & "'!InsertHeadersFromPrompt"
    Application.OnKey "^%+D", "'" & ThisWorkbook.Name & "'!InsertVarianceHeaders"
    Application.OnKey "^%+Z", "'" & ThisWorkbook.Name & "'!ApplyZeroCheckCF"
    Application.OnKey "^%+X", "'" & ThisWorkbook.Name & "'!ClearZeroCheckCF"

    '==== modBorders (Bxx) ====================================================
    Application.OnKey "^%+{UP}", "'" & ThisWorkbook.Name & "'!BorderTop"
    Application.OnKey "^%+{DOWN}", "'" & ThisWorkbook.Name & "'!BorderBottom"
    Application.OnKey "^%+{LEFT}", "'" & ThisWorkbook.Name & "'!BorderLeft"
    Application.OnKey "^%+{RIGHT}", "'" & ThisWorkbook.Name & "'!BorderRight"
    Application.OnKey "^%+B", "'" & ThisWorkbook.Name & "'!BordersOutlineInside"
    Application.OnKey "^%{-}", "'" & ThisWorkbook.Name & "'!ApplySumBar"
    Application.OnKey "^%{_}", "'" & ThisWorkbook.Name & "'!ApplyDoubleSumBar"

    '==== modUnitTags (Uxx) ===================================================
    Application.OnKey "^%+T", "'" & ThisWorkbook.Name & "'!CycleUnitTag_Value_Uniform"
    Application.OnKey "^%+O", "'" & ThisWorkbook.Name & "'!CycleUnitTag_Duration_Uniform"
    Application.OnKey "^%+P", "'" & ThisWorkbook.Name & "'!CycleUnitTag_Rate_Uniform"
    Application.OnKey "^%+{BACKSPACE}", "'" & ThisWorkbook.Name & "'!RemoveUnitTag"

    On Error GoTo 0
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

