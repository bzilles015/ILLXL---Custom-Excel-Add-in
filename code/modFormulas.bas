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
' Bound shortcuts (see modBindings):
'   M01 – InsertCAGR              Ctrl+Alt+Shift+C
'   M02 – InsertPercentChange     Ctrl+Alt+Shift+G
'   M03 – EqualsLeft              Ctrl+Alt+D
'   M04 – ApplyGrowthRate         Ctrl+Alt+Shift+W
'   M05 – InsertQuickSum          Ctrl+Alt+=
'   M06 – InsertQuickAverage      Ctrl+Alt+Shift+=
'==============================================================================

'------------------------------------------------------------------------------
' M01 - InsertCAGR  (Ctrl+Alt+Shift+C)
'     Inserts a CAGR formula: (End/Begin)^(1/n) - 1
'     Prompts for beginning value, ending value, and number of periods.
'------------------------------------------------------------------------------
Sub InsertCAGR()
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.count <> 1 Then
        MsgBox "Select a single cell where you want the CAGR formula.", vbInformation
        Exit Sub
    End If
    
    BeginMacroWithUndo
    
    Dim beginCell As Range, endCell As Range
    Dim periods As Variant
    
    On Error Resume Next
    Set beginCell = Application.InputBox("Select the BEGINNING value cell:", "CAGR - Beginning Value", Type:=8)
    If beginCell Is Nothing Then Exit Sub
    
    Set endCell = Application.InputBox("Select the ENDING value cell:", "CAGR - Ending Value", Type:=8)
    If endCell Is Nothing Then Exit Sub
    On Error GoTo 0
    
    periods = Application.InputBox("Number of periods (= end year minus start year)" & vbCrLf & _
                   "e.g. 2023 to 2026 = 3 periods, NOT 4", "CAGR Periods", 5, Type:=1)
    If VarType(periods) = vbBoolean Then Exit Sub  ' User cancelled
    If periods <= 0 Then
        MsgBox "Periods must be greater than 0.", vbExclamation
        Exit Sub
    End If
    
    ' Formula: (End/Begin)^(1/n) - 1
    ' With error handling for divide by zero
    Selection.Formula = "=IFERROR((" & endCell.Address(False, False) & "/" & _
                        beginCell.Address(False, False) & ")^(1/" & periods & ")-1,0)"
    Selection.NumberFormat = "0.0%"
    
    LogAction "InsertCAGR", Selection.Address(False, False)
    RegisterUndo "Insert CAGR"
End Sub

'------------------------------------------------------------------------------
' M02 - InsertPercentChange  (Ctrl+Alt+Shift+G)
'     Inserts a % change formula referencing cells to the left.
'     Formula: (Current - Prior) / ABS(Prior)
'     Current = 1 cell left, Prior = 2 cells left
'------------------------------------------------------------------------------
Sub InsertPercentChange()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    
    Dim c As Range
    Application.ScreenUpdating = False
    
    For Each c In Selection.Cells
        If c.Column > 2 Then
            ' Current is 1 left, Prior is 2 left
            ' Using ABS in denominator handles negative base values correctly
            c.Formula = "=IFERROR((" & c.Offset(0, -1).Address(False, False) & "-" & _
                        c.Offset(0, -2).Address(False, False) & ")/ABS(" & _
                        c.Offset(0, -2).Address(False, False) & "),0)"
            c.NumberFormat = "0.0%;(0.0%);""—"";@"
        End If
    Next c
    
    Application.ScreenUpdating = True
    LogAction "InsertPctChg", Selection.Address(False, False)
    RegisterUndo "Insert % Change"
End Sub

'------------------------------------------------------------------------------
' M03 - EqualsLeft  (Ctrl+Alt+D)
'     Links each selected cell to the cell directly on its left.
'     e.g. select D5:G5 -> D5=C5, E5=D5, F5=E5, G5=F5
'------------------------------------------------------------------------------
Sub EqualsLeft()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection.Cells
        If c.Column > 1 Then
            c.Formula = "=" & c.Offset(0, -1).Address(False, False)
        End If
    Next c
    LogAction "EqualsLeft", Selection.Address(False, False)
    RegisterUndo "Equals Left"
End Sub



'------------------------------------------------------------------------------
' M04 - ApplyGrowthRate  (Ctrl+Alt+Shift+W)
'     Inserts formula: =LEFT * (1 + rate)
'     Prompts for growth rate, then applies to all selected cells.
'------------------------------------------------------------------------------
Sub ApplyGrowthRate()
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Dim rate As Variant
    rate = Application.InputBox("Enter growth rate (as decimal, e.g., 0.05 for 5%):" & vbCrLf & _
                                "Or enter a cell reference like A1", _
                                "Growth Rate", "0.05", Type:=1 + 8)  ' Allow number or range
    
    If VarType(rate) = vbBoolean Then Exit Sub  ' User cancelled
    
    BeginMacroWithUndo
    
    Dim c As Range
    Dim rateStr As String
    
    ' Check if rate is a range reference or a number
    If TypeName(rate) = "Range" Then
        rateStr = rate.Address(True, True)  ' Use absolute reference
    Else
        rateStr = CStr(rate)
    End If
    
    Application.ScreenUpdating = False
    
    For Each c In Selection.Cells
        If c.Column > 1 Then
            c.Formula = "=" & c.Offset(0, -1).Address(False, False) & "*(1+" & rateStr & ")"
        End If
    Next c
    
    Application.ScreenUpdating = True
    LogAction "GrowthRate", Selection.Address(False, False)
    RegisterUndo "Apply Growth Rate"
End Sub

'------------------------------------------------------------------------------
' M05 - InsertQuickSum  (Ctrl+Alt+=)
'     Quick SUM formula - select destination cell, then pick range to sum.
'------------------------------------------------------------------------------
Sub InsertQuickSum()
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.count <> 1 Then
        MsgBox "Select a single cell where you want the SUM formula.", vbInformation
        Exit Sub
    End If
    
    Dim sumRange As Range
    On Error Resume Next
    Set sumRange = Application.InputBox("Select range to SUM:", "Quick SUM", Type:=8)
    On Error GoTo 0
    
    If sumRange Is Nothing Then Exit Sub
    
    BeginMacroWithUndo
    
    Selection.Formula = "=SUM(" & sumRange.Address(False, False) & ")"
    
    LogAction "QuickSum", Selection.Address(False, False)
    RegisterUndo "Quick SUM"
End Sub

'------------------------------------------------------------------------------
' M06 - InsertQuickAverage  (Ctrl+Alt+Shift+=)
'     Quick AVERAGE formula - select destination cell, then pick range.
'------------------------------------------------------------------------------
Sub InsertQuickAverage()
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.count <> 1 Then
        MsgBox "Select a single cell where you want the AVERAGE formula.", vbInformation
        Exit Sub
    End If
    
    Dim avgRange As Range
    On Error Resume Next
    Set avgRange = Application.InputBox("Select range to AVERAGE:", "Quick Average", Type:=8)
    On Error GoTo 0
    
    If avgRange Is Nothing Then Exit Sub
    
    BeginMacroWithUndo
    
    Selection.Formula = "=AVERAGE(" & avgRange.Address(False, False) & ")"
    
    LogAction "QuickAvg", Selection.Address(False, False)
    RegisterUndo "Quick Average"
End Sub



