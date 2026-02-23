Attribute VB_Name = "modCore"

Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+


'==============================================================================
' Module: modCore (Cxx)
' Purpose:
'   - Shared state (cycle indices, prev addresses)
'   - Undo buffer + custom undo engine
'   - Logging + Performance Mode
'   - Global selection helpers
'
' Public entry points:
'   C01 - BeginMacroWithUndo       C02 - RegisterUndo
'   C03 - UndoLastAction           C04 - LogAction
'   C05 - MakeRefsAbsolute         C06 - MakeRefsRelative
'   C07 - GoToNextBlank            C08 - GoToNextError
'   C09 - BreakExternalLinksInSelection
'   C10 - PerformanceModeOn        C11 - PerformanceModeOff
'   C12 - RemoveUnusedStyles
'==============================================================================

'==== Constants ====
Private Const MAX_UNDO_CELLS As Long = 5000
Private Const LOG_ROTATION_THRESHOLD As Long = 10000

'==== Performance Mode state ====
Private PerfMode As Boolean
Private PrevCalc As XlCalculation
Private PrevScreenUpdating As Boolean
Private PrevEnableEvents As Boolean
Private PrevDisplayPageBreaks As Boolean
Public LoggingEnabled As Boolean
Private CFCalcStates() As Boolean

'------------------------------------------------------------------------------
' Cycle-State Variables (reset by ResetCycleState)
'------------------------------------------------------------------------------

'--------------------Format Cycles--------------------
Public FormatNumIndex      As Integer
Public FormatDateIndex     As Integer
Public FormatPctIndex      As Integer
Public FormatCurIndex      As Integer
Public FormatOtherIndex    As Integer

'--------------------Style Cycles--------------------
Public FontCycleIndex      As Integer
Public FillCycleIndex      As Integer
Public TextCaseIndex       As Integer
Public FontColorCycleIndex As Integer

'------------------------------------------------------------------------------
' Prev-Address Trackers for per-cell cycles
'------------------------------------------------------------------------------

'--------------------Format Cycles--------------------
Public PrevNumAddress       As String
Public PrevDateAddress      As String
Public PrevPctAddress       As String
Public PrevCurAddress       As String
Public PrevOtherAddress     As String

'--------------------Style Cycles--------------------
Public PrevFontAddress      As String
Public PrevFillAddress      As String
Public PrevTextAddress      As String
Public PrevFontColorAddress As String

'------------------------------------------------------------------------------
' Border-Cycle State & Trackers
'------------------------------------------------------------------------------
Public BorderTopIndex      As Integer
Public BorderBottomIndex   As Integer
Public BorderLeftIndex     As Integer
Public BorderRightIndex    As Integer

Public PrevTopAddress      As String
Public PrevBottomAddress   As String
Public PrevLeftAddress     As String
Public PrevRightAddress    As String

'------------------------------------------------------------------------------
' Undo Buffer Type & Variables
'------------------------------------------------------------------------------
Private Type CellState
    Address      As String
    Value        As Variant
    Formula      As String
    NumberFormat As String
End Type
Private UndoBuffer()    As CellState
Private UndoBufferCount As Long

'------------------------------------------------------------------------------
' Auto_Close - Cleanup on workbook close
'------------------------------------------------------------------------------
Sub Auto_Close()
    If PerfMode Then PerformanceModeOff  ' Force cleanup if Excel crashes
End Sub

'------------------------------------------------------------------------------
' ResetCycleState
'------------------------------------------------------------------------------
Sub ResetCycleState()
    FormatNumIndex = 0:    PrevNumAddress = ""
    FormatDateIndex = 0:   PrevDateAddress = ""
    FormatPctIndex = 0:    PrevPctAddress = ""
    FormatCurIndex = 0:    PrevCurAddress = ""
    FormatOtherIndex = 0:  PrevOtherAddress = ""
    FontCycleIndex = 0:    PrevFontAddress = ""
    FillCycleIndex = 0:    PrevFillAddress = ""
    TextCaseIndex = 0:     PrevTextAddress = ""
    FontColorCycleIndex = 0: PrevFontColorAddress = ""
    LoggingEnabled = True
End Sub

'------------------------------------------------------------------------------
' C01 - BeginMacroWithUndo
'------------------------------------------------------------------------------
Sub BeginMacroWithUndo()
    If PerfMode Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    Dim sel As Range, c As Range, n As Double
    Set sel = Selection
    n = sel.Cells.CountLarge

    If n > MAX_UNDO_CELLS Then Exit Sub

    ' Clear previous buffer before allocating new one
    Erase UndoBuffer
    
    ReDim UndoBuffer(1 To CLng(n))
    UndoBufferCount = 0

    For Each c In sel.Cells
        UndoBufferCount = UndoBufferCount + 1
        With UndoBuffer(UndoBufferCount)
            .Address = c.Address(True, True)
            .Value = c.Value2
            .Formula = IIf(c.HasFormula, c.Formula, "")
            .NumberFormat = c.NumberFormat
        End With
    Next c
End Sub

'------------------------------------------------------------------------------
' C02 - RegisterUndo
'------------------------------------------------------------------------------
Sub RegisterUndo(ByVal MacroName As String)
    If PerfMode Then Exit Sub
    Application.OnUndo "Undo " & MacroName, _
        "'" & ThisWorkbook.Name & "'!UndoLastAction"
End Sub

'------------------------------------------------------------------------------
' C03 - UndoLastAction
'------------------------------------------------------------------------------
Sub UndoLastAction()
    Dim ws As Worksheet, i As Long, c As Range
    Set ws = ActiveSheet
    On Error Resume Next
    For i = 1 To UndoBufferCount
        With UndoBuffer(i)
            Set c = ws.Range(.Address)
            If .Formula <> "" Then c.Formula = .Formula Else c.Value = .Value
            c.NumberFormat = .NumberFormat
        End With
    Next i
    On Error GoTo 0
    ReDim UndoBuffer(0)
    UndoBufferCount = 0
End Sub

'------------------------------------------------------------------------------
' C04 - LogAction
'------------------------------------------------------------------------------
Sub LogAction(ByVal actionName As String, ByVal Target As String)
    If Not LoggingEnabled Then Exit Sub
    Const LogSheetName As String = "AddInLog"
    Dim ws  As Worksheet, nxt As Long
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(LogSheetName): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = LogSheetName
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1:C1").Value = Array("Timestamp", "Action", "Target")
    End If
    nxt = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    
    ' Rotate log if it gets too large
    If nxt > LOG_ROTATION_THRESHOLD Then
        ws.Rows("2:5001").Delete
        nxt = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    End If
    
    ws.Cells(nxt, 1).Value = Now
    ws.Cells(nxt, 2).Value = actionName
    ws.Cells(nxt, 3).Value = Target
End Sub

'------------------------------------------------------------------------------
' C05 - MakeRefsAbsolute  (Ctrl+Alt+Shift+A)
'------------------------------------------------------------------------------
Sub MakeRefsAbsolute()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection.Cells
        If c.HasFormula Then
            c.Formula = Application.ConvertFormula(c.Formula, xlA1, xlA1, xlAbsolute)
        End If
    Next c
    LogAction "MakeAbs", Selection.Address(False, False)
    RegisterUndo "Make Refs Absolute"
End Sub

'------------------------------------------------------------------------------
' C06 - MakeRefsRelative  (Ctrl+Alt+Shift+R)
'------------------------------------------------------------------------------
Sub MakeRefsRelative()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection.Cells
        If c.HasFormula Then
            c.Formula = Application.ConvertFormula(c.Formula, xlA1, xlA1, xlRelative)
        End If
    Next c
    LogAction "MakeRel", Selection.Address(False, False)
    RegisterUndo "Make Refs Relative"
End Sub

'------------------------------------------------------------------------------
' C07 - GoToNextBlank  (Ctrl+Alt+Shift+N)
'------------------------------------------------------------------------------
Sub GoToNextBlank()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim blanks As Range, nxt As Range
    On Error Resume Next
    Set blanks = Selection.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
    If blanks Is Nothing Then Beep: Exit Sub
    Set nxt = blanks.Find(What:="", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, _
                          SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not nxt Is Nothing Then nxt.Select Else Beep
End Sub

'------------------------------------------------------------------------------
' C08 - GoToNextError  (Ctrl+Alt+Shift+E)
'------------------------------------------------------------------------------
Sub GoToNextError()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim errs As Range, nxt As Range
    On Error Resume Next
    Set errs = Selection.SpecialCells(xlCellTypeFormulas, xlErrors)
    If errs Is Nothing Then Set errs = Selection.SpecialCells(xlCellTypeConstants, xlErrors)
    On Error GoTo 0
    If errs Is Nothing Then Beep: Exit Sub
    Set nxt = errs.Find(What:="*", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not nxt Is Nothing Then nxt.Select Else Beep
End Sub

'------------------------------------------------------------------------------
' C09 - BreakExternalLinksInSelection  (Ctrl+Alt+Shift+L)
'------------------------------------------------------------------------------
Sub BreakExternalLinksInSelection()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range, f As String
    For Each c In Selection.Cells
        If c.HasFormula Then
            f = c.Formula
            If InStr(1, f, "[", vbTextCompare) > 0 Or InStr(1, f, "http", vbTextCompare) > 0 Then
                c.Value = c.Value
            End If
        End If
    Next c
    LogAction "BreakExtLinks", Selection.Address(False, False)
    RegisterUndo "Break External Links"
End Sub

'------------------------------------------------------------------------------
' C10 - PerformanceModeOn  (Ctrl+Alt+Shift+M)
'------------------------------------------------------------------------------
Sub PerformanceModeOn()
    On Error Resume Next
    PrevCalc = Application.Calculation
    PrevScreenUpdating = Application.ScreenUpdating
    PrevEnableEvents = Application.EnableEvents
    PrevDisplayPageBreaks = ActiveWindow.DisplayPageBreaks
    On Error GoTo 0

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next: ActiveWindow.DisplayPageBreaks = False: On Error GoTo 0

    ' Disable CF calc on all sheets (store previous states)
    Dim ws As Worksheet, i As Long
    ReDim CFCalcStates(1 To ThisWorkbook.Worksheets.count)
    i = 0
    For Each ws In ThisWorkbook.Worksheets
        i = i + 1
        On Error Resume Next
        CFCalcStates(i) = ws.EnableFormatConditionsCalculation
        ws.EnableFormatConditionsCalculation = False
        On Error GoTo 0
    Next ws

    PerfMode = True
    LoggingEnabled = False
    
    ' Add visual indicator in Excel window title
    Application.Caption = "Excel [PERFORMANCE MODE - UNDO DISABLED]"
    Application.StatusBar = "Performance Mode: ON  (Calc=Manual, ScreenUpdating OFF, Events OFF, CF OFF)"
End Sub

'------------------------------------------------------------------------------
' C11 - PerformanceModeOff  (Ctrl+Alt+Shift+M)
'------------------------------------------------------------------------------
Sub PerformanceModeOff()
    ' Restore CF calc
    Dim ws As Worksheet, i As Long
    i = 0
    For Each ws In ThisWorkbook.Worksheets
        i = i + 1
        On Error Resume Next
        If Not IsEmpty(CFCalcStates) Then ws.EnableFormatConditionsCalculation = CFCalcStates(i)
        On Error GoTo 0
    Next ws

    On Error Resume Next
    Application.Calculation = PrevCalc
    Application.ScreenUpdating = PrevScreenUpdating
    Application.EnableEvents = PrevEnableEvents
    ActiveWindow.DisplayPageBreaks = PrevDisplayPageBreaks
    Application.Caption = "Excel"
    Application.StatusBar = False
    On Error GoTo 0

    PerfMode = False
    LoggingEnabled = True
End Sub

'------------------------------------------------------------------------------
' TogglePerformanceMode - Helper toggle
'------------------------------------------------------------------------------
Sub TogglePerformanceMode()
    If PerfMode Then
        PerformanceModeOff
    Else
        PerformanceModeOn
    End If
End Sub

'------------------------------------------------------------------------------
' C12 - RemoveUnusedStyles
'     Workbook cleanup - removes custom styles not in use
'------------------------------------------------------------------------------
Sub RemoveUnusedStyles()
    On Error Resume Next
    
    Dim sty As Style
    Dim removeCount As Long
    removeCount = 0
    
    For Each sty In ActiveWorkbook.Styles
        If Not sty.BuiltIn Then
            sty.Delete
            removeCount = removeCount + 1
        End If
    Next sty
    
    On Error GoTo 0
    MsgBox "Removed " & removeCount & " unused custom styles.", vbInformation
    LogAction "RemoveStyles", CStr(removeCount)
End Sub



