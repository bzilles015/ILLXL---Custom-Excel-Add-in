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
'   C09 - BreakExternalLinksInSelection  (retired - commented out)
'   C10 - PerformanceModeOn        C11 - PerformanceModeOff
'   C12 - RemoveUnusedStyles
'   C13 - ToggleErrorChecking
'   C14 - IsNumberCell (helper)    C15 - SpecialCellsIn (helper)
'==============================================================================

'==== Constants ====
Public Const ILLXL_VERSION As String = "dev"           ' stamped by build.ps1 (-Version)
Private Const MAX_UNDO_CELLS As Long = 5000
Private Const LOG_ROTATION_THRESHOLD As Long = 10000
Private Const UNDO_SHEET As String = "_Undo"           ' scratch sheet inside the add-in
Private Const PERF_CAPTION As String = "Excel [PERFORMANCE MODE - UNDO DISABLED]"

'==== Performance Mode state ====
Private PerfMode As Boolean
Private PerfSaved As Boolean                  ' True once the "before" settings were captured
Private PrevCalc As XlCalculation
Private PrevEnableEvents As Boolean
Private PrevDisplayPageBreaks As Boolean
Private PerfWorkbook As Workbook              ' the model Performance Mode was turned on in
Private CFCalcStates As Object                ' sheet name -> EnableFormatConditionsCalculation
Public LoggingEnabled As Boolean

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
' Undo Buffer
'   Contents (value or formula) are kept per cell in UndoBuffer.
'   Formatting (fill, font, borders, number format, alignment, CF) is kept
'   as a copy of the range on the add-in's hidden _Undo sheet.
'   The buffer remembers its own sheet, so Ctrl+Z still lands on the right
'   sheet after you switch tabs.
'------------------------------------------------------------------------------
Private Type CellState
    Address      As String
    Value        As Variant
    Formula      As String
    SpillChild   As Boolean      ' cell filled by a dynamic-array spill - leave empty on undo
End Type
Private UndoBuffer()    As CellState
Private UndoBufferCount As Long
Private UndoReady       As Boolean          ' True only when the buffer matches the last action
Private UndoSheet       As Object           ' Worksheet (or Chart sheet) the buffer belongs to
Private UndoRange       As Range            ' captured cells (all areas); Nothing = tab only
Private UndoHasFormats  As Boolean          ' False on filtered sheets (Excel copies visible rows only)
Private UndoTabNone     As Boolean          ' tab had no color before the action
Private UndoTabColor    As Long

'------------------------------------------------------------------------------
' Auto_Close - Cleanup when the add-in unloads or Excel closes
'------------------------------------------------------------------------------
Sub Auto_Close()
    If PerfMode Then PerformanceModeOff     ' never leave Excel in manual calc
    UnbindAllKeys                           ' give Excel its own shortcuts back
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
'     Snapshot the cells a macro is about to change. Defaults to the
'     selection; pass Target when a macro writes outside it (e.g. headers).
'     Also remembers the sheet tab color. Over MAX_UNDO_CELLS: no snapshot,
'     and RegisterUndo then skips, so Ctrl+Z can never replay an old action.
'------------------------------------------------------------------------------
Sub BeginMacroWithUndo(Optional ByVal Target As Range)
    EnsureSettings                                  ' presets loaded (e.g. after a VBA reset)
    UndoReady = False                               ' disarm any older snapshot
    UndoBufferCount = 0
    Set UndoRange = Nothing
    If PerfMode Then Exit Sub

    On Error GoTo Fail
    Set UndoSheet = ActiveSheet
    UndoTabNone = (UndoSheet.Tab.ColorIndex = xlColorIndexNone)
    If Not UndoTabNone Then UndoTabColor = UndoSheet.Tab.Color

    If Target Is Nothing Then
        If TypeName(Selection) = "Range" Then Set Target = Selection
    End If
    If Not Target Is Nothing Then
        If Target.Cells.CountLarge > MAX_UNDO_CELLS Then
            FlashStatus "ILLXL: over " & Format$(MAX_UNDO_CELLS, "#,##0") & " cells - this action can't be undone with Ctrl+Z."
            Exit Sub
        End If
        Set UndoSheet = Target.Worksheet

        '-- 1) Formatting: copy each area to the same address on _Undo
        '      (skipped on filtered sheets - contents are still undone)
        Dim scratch As Worksheet, area As Range, c As Range
        Set scratch = ThisWorkbook.Worksheets(UNDO_SHEET)
        scratch.Cells.Clear
        UndoHasFormats = Not UndoSheet.FilterMode
        If UndoHasFormats Then
            For Each area In Target.Areas
                area.Copy Destination:=scratch.Range(area.Address)
            Next area
            scratch.Cells.ClearContents             ' keep the formats, drop copied formulas
        End If

        '-- 2) Contents: formula or value for every cell
        ReDim UndoBuffer(1 To CLng(Target.Cells.CountLarge))
        For Each c In Target.Cells
            UndoBufferCount = UndoBufferCount + 1
            With UndoBuffer(UndoBufferCount)
                .Address = c.Address
                .SpillChild = IsSpillChild(c)
                If c.HasFormula Then
                    .Formula = FormulaOf(c)
                Else
                    .Formula = ""
                    .Value = c.Value2
                End If
            End With
        Next c
        Set UndoRange = Target
    End If

    UndoReady = True
    Exit Sub
Fail:
    UndoReady = False
    UndoBufferCount = 0
End Sub

'------------------------------------------------------------------------------
' C02 - RegisterUndo
'     Wires Ctrl+Z to UndoLastAction - only when a fresh snapshot exists.
'------------------------------------------------------------------------------
Sub RegisterUndo(ByVal MacroName As String)
    If PerfMode Or Not UndoReady Then Exit Sub
    Application.OnUndo "Undo " & MacroName, _
        "'" & ThisWorkbook.Name & "'!UndoLastAction"
End Sub

'------------------------------------------------------------------------------
' C03 - UndoLastAction
'     Puts back formatting, then contents, then the tab color.
'------------------------------------------------------------------------------
Sub UndoLastAction()
    If Not UndoReady Then Exit Sub
    UndoReady = False                               ' one undo per snapshot

    Dim scratch As Worksheet, area As Range, c As Range, i As Long
    On Error GoTo Gone
    Dim probe As String: probe = UndoSheet.Name     ' errors if the workbook was closed
    On Error GoTo 0
    Application.ScreenUpdating = False

    If Not UndoRange Is Nothing Then
        '-- 1) Formatting back (this also empties the cells)
        Set scratch = ThisWorkbook.Worksheets(UNDO_SHEET)
        On Error Resume Next                        ' e.g. part of a legacy array formula
        If UndoHasFormats Then
            For Each area In UndoRange.Areas
                scratch.Range(area.Address).Copy Destination:=area
            Next area
        End If

        '-- 2) Contents back
        For i = 1 To UndoBufferCount
            With UndoBuffer(i)
                Set c = UndoSheet.Range(.Address)
                If .SpillChild Then
                    ' leave empty so the parent formula can spill again
                ElseIf Len(.Formula) > 0 Then
                    SetFormula c, .Formula
                ElseIf IsEmpty(.Value) Then
                    If Not UndoHasFormats Then c.ClearContents  ' (the format copy already emptied it)
                Else
                    c.Value2 = .Value
                    ' text that looks like a number ("00123") must stay text
                    If VarType(.Value) = vbString And VarType(c.Value2) <> vbString Then
                        c.Value2 = "'" & .Value
                    End If
                End If
            End With
        Next i
        On Error GoTo 0
    End If

    '-- 3) Tab color back
    On Error Resume Next
    If UndoTabNone Then
        UndoSheet.Tab.ColorIndex = xlColorIndexNone
    Else
        UndoSheet.Tab.Color = UndoTabColor
    End If
    On Error GoTo 0

    ThisWorkbook.Worksheets(UNDO_SHEET).Cells.Clear
    UndoBufferCount = 0
    Exit Sub
Gone:
    UiMsg "Nothing to undo - the sheet from the last ILLXL action is no longer open."
End Sub

'--- Undo helpers -------------------------------------------------------------
Private Function IsSpillChild(ByVal c As Range) As Boolean
    On Error Resume Next                            ' HasSpill needs Excel 365
    If c.HasSpill Then IsSpillChild = (c.SpillParent.Address <> c.Address)
End Function

Public Function FormulaOf(ByVal c As Range) As String
    On Error Resume Next
    FormulaOf = c.Formula2                          ' keeps dynamic arrays intact (365)
    If Err.Number <> 0 Then FormulaOf = c.Formula   ' older Excel
End Function

Public Sub SetFormula(ByVal c As Range, ByVal f As String)
    On Error Resume Next
    c.Formula2 = f
    If Err.Number <> 0 Then Err.Clear: c.Formula = f
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
' C05 - MakeRefsAbsolute  (Ctrl+Alt+Shift+A)   A1 -> $A$1
' C06 - MakeRefsRelative  (Ctrl+Alt+Shift+R)   $A$1 -> A1
'     Formula cells only. Excel can't convert formulas over 255 characters -
'     those are skipped and counted, the rest still convert.
'------------------------------------------------------------------------------
Sub MakeRefsAbsolute()
    ConvertRefs xlAbsolute, "Make Refs Absolute"
End Sub

Sub MakeRefsRelative()
    ConvertRefs xlRelative, "Make Refs Relative"
End Sub

Private Sub ConvertRefs(ByVal toType As XlReferenceType, ByVal undoName As String)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim formulas As Range, c As Range, f As String, skipped As Long
    Set formulas = SpecialCellsIn(Selection, xlCellTypeFormulas)
    If formulas Is Nothing Then Exit Sub
    BeginMacroWithUndo
    For Each c In formulas.Cells
        f = ""
        On Error Resume Next
        f = Application.ConvertFormula(FormulaOf(c), xlA1, xlA1, toType)
        On Error GoTo 0
        If Len(f) = 0 Or VarType(f) <> vbString Then
            skipped = skipped + 1
        ElseIf f <> FormulaOf(c) Then
            SetFormula c, f
        End If
    Next c
    LogAction undoName, Selection.Address(False, False)
    RegisterUndo undoName
    If skipped > 0 Then UiMsg skipped & " formula(s) are longer than Excel can convert (255 characters) and were left as they were."
End Sub

'------------------------------------------------------------------------------
' C07 - GoToNextBlank  (Ctrl+Alt+Shift+N)
' C08 - GoToNextError  (Ctrl+Alt+Shift+E)
'     Jump to the next blank / error cell after the active cell, row by row,
'     wrapping to the top. With a range selected it searches that range and
'     keeps it selected (so you can keep pressing); with one cell selected it
'     searches the whole used area of the sheet.
'------------------------------------------------------------------------------
Sub GoToNextBlank()
    If TypeName(Selection) <> "Range" Then Exit Sub
    JumpToNext SpecialCellsIn(SearchArea(), xlCellTypeBlanks)
End Sub

Sub GoToNextError()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim f As Range, k As Range
    Set f = SpecialCellsIn(SearchArea(), xlCellTypeFormulas, xlErrors)   ' #REF! from formulas
    Set k = SpecialCellsIn(SearchArea(), xlCellTypeConstants, xlErrors)  ' typed #N/A
    If f Is Nothing Then
        Set f = k
    ElseIf Not k Is Nothing Then
        Set f = Union(f, k)
    End If
    JumpToNext f
End Sub

Private Function SearchArea() As Range
    If Selection.Cells.CountLarge = 1 Then Set SearchArea = ActiveSheet.UsedRange Else Set SearchArea = Selection
End Function

' First target after the active cell (row by row), else the first one overall
Private Sub JumpToNext(ByVal targets As Range)
    Dim a As Range, c As Range, best As Range, first As Range, r0 As Long, c0 As Long
    If targets Is Nothing Then Beep: Exit Sub
    r0 = ActiveCell.Row: c0 = ActiveCell.Column
    For Each a In targets.Areas
        For Each c In a.Cells
            If IsBefore(c, first) Then Set first = c
            If c.Row > r0 Or (c.Row = r0 And c.Column > c0) Then
                If IsBefore(c, best) Then Set best = c
            End If
        Next c
    Next a
    If best Is Nothing Then Set best = first                 ' wrap to the top
    If Selection.Cells.CountLarge = 1 Then best.Select Else best.Activate
End Sub

Private Function IsBefore(ByVal c As Range, ByVal other As Range) As Boolean
    If other Is Nothing Then IsBefore = True: Exit Function
    IsBefore = (c.Row < other.Row) Or (c.Row = other.Row And c.Column < other.Column)
End Function

'------------------------------------------------------------------------------
' C09 - BreakExternalLinksInSelection  (retired)
'     Treated any "[" as an external link, so table formulas such as
'     =SUM(tblSales[Amount]) were turned into values. Not used - kept for
'     reference. Data > Edit Links > Break Link does this safely.
'------------------------------------------------------------------------------
'Sub BreakExternalLinksInSelection()
'    If TypeName(Selection) <> "Range" Then Exit Sub
'    BeginMacroWithUndo
'    Dim c As Range, f As String
'    For Each c In Selection.Cells
'        If c.HasFormula Then
'            f = c.Formula
'            If InStr(1, f, "[", vbTextCompare) > 0 Or InStr(1, f, "http", vbTextCompare) > 0 Then
'                c.Value = c.Value
'            End If
'        End If
'    Next c
'    LogAction "BreakExtLinks", Selection.Address(False, False)
'    RegisterUndo "Break External Links"
'End Sub

'------------------------------------------------------------------------------
' C10 - PerformanceModeOn  (Ctrl+Alt+Shift+M)
'     Acts on the ACTIVE workbook (your model), not on the add-in.
'------------------------------------------------------------------------------
Sub PerformanceModeOn()
    If ActiveWorkbook Is Nothing Then Exit Sub
    On Error Resume Next
    PrevCalc = Application.Calculation
    PrevEnableEvents = Application.EnableEvents
    PrevDisplayPageBreaks = ActiveWindow.DisplayPageBreaks
    PerfSaved = True

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveWindow.DisplayPageBreaks = False

    ' Disable CF calc on every sheet of the model (remember each sheet's setting)
    Dim ws As Worksheet
    Set PerfWorkbook = ActiveWorkbook
    Set CFCalcStates = CreateObject("Scripting.Dictionary")
    For Each ws In PerfWorkbook.Worksheets
        CFCalcStates(ws.Name) = ws.EnableFormatConditionsCalculation
        ws.EnableFormatConditionsCalculation = False
    Next ws
    On Error GoTo 0

    PerfMode = True
    LoggingEnabled = False

    ' Visual indicator in the Excel title bar (also how Toggle finds us after a VBA reset)
    Application.Caption = PERF_CAPTION
    Application.StatusBar = "Performance Mode: ON  (Calc=Manual, Events OFF, CF OFF)"
End Sub

'------------------------------------------------------------------------------
' C11 - PerformanceModeOff  (Ctrl+Alt+Shift+M)
'     Safe even after a VBA reset wiped the saved state: falls back to
'     Automatic calc and events ON.
'------------------------------------------------------------------------------
Sub PerformanceModeOff()
    Dim ws As Worksheet
    On Error Resume Next

    ' Restore CF calc on the model it was turned off in
    If Not PerfWorkbook Is Nothing And Not CFCalcStates Is Nothing Then
        For Each ws In PerfWorkbook.Worksheets
            If CFCalcStates.Exists(ws.Name) Then ws.EnableFormatConditionsCalculation = CFCalcStates(ws.Name)
        Next ws
    End If

    If PerfSaved Then
        Application.Calculation = PrevCalc
        Application.EnableEvents = PrevEnableEvents
        ActiveWindow.DisplayPageBreaks = PrevDisplayPageBreaks
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If
    Application.ScreenUpdating = True
    Application.Caption = Empty                     ' back to Excel's normal title
    Application.StatusBar = False
    On Error GoTo 0

    Set PerfWorkbook = Nothing
    Set CFCalcStates = Nothing
    PerfSaved = False
    PerfMode = False
    LoggingEnabled = True
End Sub

'------------------------------------------------------------------------------
' TogglePerformanceMode - Helper toggle
'     The title-bar check catches the case where VBA was reset while ON.
'------------------------------------------------------------------------------
Sub TogglePerformanceMode()
    If PerfMode Or Application.Caption = PERF_CAPTION Then
        PerformanceModeOff
    Else
        PerformanceModeOn
    End If
End Sub

'------------------------------------------------------------------------------
' C12 - RemoveUnusedStyles
'     Workbook cleanup - deletes custom cell styles that no cell uses.
'     Built-in styles and any style in use are kept. Asks first.
'------------------------------------------------------------------------------
Sub RemoveUnusedStyles()
    If ActiveWorkbook Is Nothing Then Exit Sub
    Dim used As Object, ws As Worksheet, i As Long, removeCount As Long, sty As Style, locked As String

    ' Excel refuses to delete ANY style while a sheet in the workbook is protected
    For Each ws In ActiveWorkbook.Worksheets
        If ws.ProtectContents Then locked = locked & vbCrLf & "   " & ws.Name
    Next ws
    If Len(locked) > 0 Then
        UiMsg "Excel can't delete styles while these sheets are protected:" & locked & vbCrLf & vbCrLf & _
              "Unprotect them (Review > Unprotect Sheet), run Clean Up Styles, then protect them again.", vbExclamation
        Exit Sub
    End If

    If UiMsg("Delete custom cell styles that no cell in '" & ActiveWorkbook.Name & _
             "' uses?" & vbCrLf & "Built-in styles and styles in use are kept.", _
             vbOKCancel + vbQuestion) <> vbOK Then Exit Sub

    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare

    ' 1) Collect every style actually applied to a cell
    For Each ws In ActiveWorkbook.Worksheets
        Application.StatusBar = "ILLXL: scanning styles on " & ws.Name & "..."
        CollectStyles ws.UsedRange, used
    Next ws

    ' 2) Delete the custom ones nobody uses (backwards - the collection shrinks)
    On Error Resume Next
    For i = ActiveWorkbook.Styles.count To 1 Step -1
        Set sty = ActiveWorkbook.Styles(i)
        If Not sty.BuiltIn Then
            If Not used.Exists(sty.Name) Then
                sty.Delete
                If Err.Number = 0 Then removeCount = removeCount + 1
                Err.Clear
            End If
        End If
    Next i
    On Error GoTo 0

    Application.StatusBar = False
    UiMsg "Removed " & removeCount & " unused custom styles."
    LogAction "RemoveStyles", CStr(removeCount)
End Sub

' Splits the range in half until each piece has one style - fast on big
' sheets because most of a model is a single style (Normal).
Private Sub CollectStyles(ByVal rng As Range, ByVal used As Object)
    Dim half As Long
    If TypeName(rng.Style) = "Style" Then                ' mixed styles = Nothing/Null in VBA
        used(rng.Style.Name) = True
    ElseIf rng.Rows.count > 1 Then
        half = rng.Rows.count \ 2
        CollectStyles rng.Resize(half), used
        CollectStyles rng.Offset(half).Resize(rng.Rows.count - half), used
    Else
        half = rng.Columns.count \ 2
        CollectStyles rng.Resize(, half), used
        CollectStyles rng.Offset(, half).Resize(, rng.Columns.count - half), used
    End If
End Sub

'------------------------------------------------------------------------------
' C13 - ToggleErrorChecking     Ctrl+Alt+Shift+K
'     Remove green triangles when you need a presentation ready
'------------------------------------------------------------------------------
Sub ToggleErrorChecking()
    With Application.ErrorCheckingOptions
        .BackgroundChecking = Not .BackgroundChecking
        Application.StatusBar = IIf(.BackgroundChecking, "Error checking ON", "Error checking OFF")
    End With
End Sub

'------------------------------------------------------------------------------
' C14 - IsNumberCell
'     True only for real numbers - blanks, text and dates return False.
'     (VBA's IsNumeric says True for a blank cell, which turned blanks into 0.)
'------------------------------------------------------------------------------
Public Function IsNumberCell(ByVal c As Range) As Boolean
    Select Case VarType(c.Value)
        Case vbDouble, vbCurrency, vbInteger, vbLong, vbSingle, vbDecimal
            IsNumberCell = True
    End Select
End Function

'------------------------------------------------------------------------------
' C15 - SpecialCellsIn
'     SpecialCells kept inside rng. Excel's SpecialCells on a SINGLE cell
'     searches the whole sheet - this trims the result back to rng.
'     Returns Nothing when no cells match.
'------------------------------------------------------------------------------
Public Function SpecialCellsIn(ByVal rng As Range, ByVal cellType As XlCellType, _
                               Optional ByVal valueType As Variant) As Range
    Dim hit As Range
    On Error Resume Next
    If IsMissing(valueType) Then
        Set hit = rng.SpecialCells(cellType)
    Else
        Set hit = rng.SpecialCells(cellType, valueType)
    End If
    On Error GoTo 0
    If Not hit Is Nothing Then Set SpecialCellsIn = Intersect(rng, hit)
End Function

