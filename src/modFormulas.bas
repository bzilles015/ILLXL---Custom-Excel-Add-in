Attribute VB_Name = "modFormulas"
Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Module: modFormulas (Mxx)
' Purpose:
'   Formula insertion tools for financial modeling - CAGR, % change,
'   equals left, growth rates, and aggregation shortcuts.
'
'   Picked cells can be on another sheet or workbook - RefFrom adds the
'   sheet name so the formula points at the right place. Numbers are written
'   with "." decimals (NumText) so formulas work on any regional setting.
'
' Default shortcuts (see the active preset):
'   M01 - InsertCAGR              Ctrl+Alt+Shift+C
'   M02 - InsertPercentChange     Ctrl+Alt+Shift+V
'   M03 - EqualsLeft              Ctrl+Alt+L
'   M04 - ApplyGrowthRate         Ctrl+Alt+Shift+G
'   M05 - InsertQuickSum          Ctrl+Alt+=
'   M06 - InsertQuickAverage      Ctrl+Alt+Shift+=
'==============================================================================

'------------------------------------------------------------------------------
' M01 - InsertCAGR  (Ctrl+Alt+Shift+C)
'     Inserts a CAGR formula: (End/Begin)^(1/n) - 1
'     Prompts for beginning value, ending value, and number of periods.
'     n = end year MINUS start year (2023 -> 2026 = 3 periods, not 4).
'------------------------------------------------------------------------------
Sub InsertCAGR()
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.CountLarge <> 1 Then
        UiMsg "Select a single cell where you want the CAGR formula."
        Exit Sub
    End If

    Dim host As Range, beginCell As Variant, endCell As Variant, periods As Variant
    Set host = Selection

    StoreIt beginCell, UiPick("Select the BEGINNING value cell:", "CAGR - Beginning Value", 8)
    If Not IsObject(beginCell) Then Exit Sub                          ' cancelled
    StoreIt endCell, UiPick("Select the ENDING value cell:", "CAGR - Ending Value", 8)
    If Not IsObject(endCell) Then Exit Sub

    periods = UiPick("Number of periods (= end year minus start year)" & vbCrLf & _
                     "e.g. 2023 to 2026 = 3 periods, NOT 4", "CAGR Periods", 1, 5)
    If VarType(periods) = vbBoolean Then Exit Sub                     ' cancelled
    If periods <= 0 Then
        UiMsg "Periods must be greater than 0.", vbExclamation
        Exit Sub
    End If

    BeginMacroWithUndo host
    ' Formula: (End/Begin)^(1/n) - 1, shown as 0 if a value is missing or zero
    host.Formula = "=IFERROR((" & RefFrom(host, endCell.Cells(1, 1)) & "/" & _
                   RefFrom(host, beginCell.Cells(1, 1)) & ")^(1/" & NumText(periods) & ")-1,0)"
    host.NumberFormat = "0.0%"

    LogAction "InsertCAGR", host.Address(False, False)
    RegisterUndo "Insert CAGR"
End Sub

' Store a UiPick result as-is: a Range stays a Range, False stays False.
' (Passing it in as an argument is what keeps the Range from being read as a value.)
Private Sub StoreIt(ByRef dst As Variant, ByVal v As Variant)
    If IsObject(v) Then Set dst = v Else dst = v
End Sub

'------------------------------------------------------------------------------
' M02 - InsertPercentChange  (Ctrl+Alt+Shift+V)
'     Inserts a % change formula referencing cells to the left.
'     Formula: (Current - Prior) / ABS(Prior)
'     Current = 1 cell left, Prior = 2 cells left
'------------------------------------------------------------------------------
Sub InsertPercentChange()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim tgt As Range, c As Range
    Set tgt = WriteCells(Selection)                             ' whole columns: used rows only
    If tgt Is Nothing Then Exit Sub
    If Not ConfirmBig(tgt, "% Change") Then Exit Sub
    BeginMacroWithUndo tgt

    Application.ScreenUpdating = False
    For Each c In tgt.Cells
        If c.Column > 2 Then
            ' Using ABS in the denominator handles negative base values correctly
            c.Formula = "=IFERROR((" & c.Offset(0, -1).Address(False, False) & "-" & _
                        c.Offset(0, -2).Address(False, False) & ")/ABS(" & _
                        c.Offset(0, -2).Address(False, False) & "),0)"
            c.NumberFormat = cfgPctFormats(0)          ' first preset [PercentFormats]
        End If
    Next c
    Application.ScreenUpdating = True

    LogAction "InsertPctChg", tgt.Address(False, False)
    RegisterUndo "Insert % Change"
End Sub

'------------------------------------------------------------------------------
' M03 - EqualsLeft  (Ctrl+Alt+L)
'     Links each selected cell to the cell directly on its left.
'     e.g. select D5:G5 -> D5=C5, E5=D5, F5=E5, G5=F5
'------------------------------------------------------------------------------
Sub EqualsLeft()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim tgt As Range, c As Range
    Set tgt = WriteCells(Selection)                             ' whole columns: used rows only
    If tgt Is Nothing Then Exit Sub
    If Not ConfirmBig(tgt, "Equals Left") Then Exit Sub
    BeginMacroWithUndo tgt

    Application.ScreenUpdating = False
    For Each c In tgt.Cells
        If c.Column > 1 Then c.Formula = "=" & c.Offset(0, -1).Address(False, False)
    Next c
    Application.ScreenUpdating = True

    LogAction "EqualsLeft", tgt.Address(False, False)
    RegisterUndo "Equals Left"
End Sub

'------------------------------------------------------------------------------
' M04 - ApplyGrowthRate  (Ctrl+Alt+Shift+G)
'     Inserts formula: =LEFT * (1 + rate)
'     Type a rate (0.05 or 5%) or click the cell that holds it (locked with $,
'     sheet name added if it's on another sheet).
'------------------------------------------------------------------------------
Sub ApplyGrowthRate()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim tgt As Range, c As Range, rate As Variant, rateStr As String
    Set tgt = WriteCells(Selection)                             ' whole columns: used rows only
    If tgt Is Nothing Then Exit Sub
    If Not ConfirmBig(tgt, "Growth Rate") Then Exit Sub

    StoreIt rate, UiPick("Enter growth rate (as decimal, e.g., 0.05 for 5%):" & vbCrLf & _
                         "Or click the cell that holds the rate", "Growth Rate", 1 + 8, "0.05")
    If IsObject(rate) Then
        rateStr = RefFrom(tgt, rate.Cells(1, 1), True)          ' locked, sheet-qualified
    ElseIf VarType(rate) = vbBoolean Then
        Exit Sub                                                ' cancelled
    Else
        rateStr = NumText(CDbl(rate))
    End If

    BeginMacroWithUndo tgt
    Application.ScreenUpdating = False
    For Each c In tgt.Cells
        If c.Column > 1 Then
            c.Formula = "=" & c.Offset(0, -1).Address(False, False) & "*(1+" & rateStr & ")"
        End If
    Next c
    Application.ScreenUpdating = True

    LogAction "GrowthRate", tgt.Address(False, False)
    RegisterUndo "Apply Growth Rate"
End Sub

'------------------------------------------------------------------------------
' M05 - InsertQuickSum  (Ctrl+Alt+=)
' M06 - InsertQuickAverage  (Ctrl+Alt+Shift+=)
'     Select the destination cell, then pick the range (any sheet).
'------------------------------------------------------------------------------
Sub InsertQuickSum()
    InsertQuickFunction "SUM", "Quick SUM"
End Sub

Sub InsertQuickAverage()
    InsertQuickFunction "AVERAGE", "Quick Average"
End Sub

Private Sub InsertQuickFunction(ByVal fn As String, ByVal title As String)
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.CountLarge <> 1 Then
        UiMsg "Select a single cell where you want the " & fn & " formula."
        Exit Sub
    End If

    Dim host As Range, src As Variant
    Set host = Selection
    StoreIt src, UiPick("Select range to " & fn & ":", title, 8)
    If Not IsObject(src) Then Exit Sub                                ' cancelled

    BeginMacroWithUndo host
    host.Formula = "=" & fn & "(" & RefFrom(host, src) & ")"
    LogAction "Quick" & fn, host.Address(False, False)
    RegisterUndo title
End Sub
