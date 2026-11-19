Attribute VB_Name = "modUi"

Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Module: modUi (Qxx)
' Purpose:
'   1) RunCmd - the ONE way every shortcut, ribbon button and gallery runs a
'      command. A single error handler means a failure never drops you into
'      the VBA Debug/End dialog: Excel is put back to normal and you get a
'      plain-English message instead.
'   2) Every message box and prompt in ILLXL goes through Ui* here, so the
'      automated tests can answer them (UiTestMode) instead of a person.
'   3) Small shared guards: WorkCells (skip the empty part of a whole-column
'      selection) and ConfirmBig (ask before writing to a huge selection).
'
' Public entry points:
'   Q01 - RunCmd                    Q02 - UiMsg / UiInput / UiPick
'   Q03 - UiOpenFile / UiSaveFile   Q04 - WorkCells / ConfirmBig / RefFrom
'   Q05 - UiTestMode (tests)
'==============================================================================

Public Const BIG_SELECTION As Long = 50000       ' ask before writing more cells than this

'==== Test mode: scripted answers instead of dialogs ====
Private TestMode  As Boolean
Private Answers   As Collection                  ' next answers, in order
Private Shown     As String                      ' every message "shown" in test mode
Private LastErr   As String                      ' last error RunCmd caught
Private Picked    As Variant                     ' UiPick result holder (see Keep)
Private LastCmd   As String                      ' last command RunCmd started

'------------------------------------------------------------------------------
' Q01 - RunCmd - run one ILLXL command safely
'     Errors anywhere inside the command land here (VBA hands an unhandled
'     error up to the nearest handler). Unknown names are reported too.
'------------------------------------------------------------------------------
Public Sub RunCmd(ByVal cmd As String)
    On Error GoTo Failed
    LastErr = ""
    LastCmd = cmd                                    ' tests check a key press reached the right command
    Select Case cmd
        '---- Core / Settings
        Case "ShowSettings":                  ShowSettings
        Case "TogglePerformanceMode":         TogglePerformanceMode
        Case "MakeRefsAbsolute":              MakeRefsAbsolute
        Case "MakeRefsRelative":              MakeRefsRelative
        Case "GoToNextBlank":                 GoToNextBlank
        Case "GoToNextError":                 GoToNextError
        Case "ToggleErrorChecking":           ToggleErrorChecking
        Case "RemoveUnusedStyles":            RemoveUnusedStyles
        '---- Formats
        Case "CycleNumberFormat":             CycleNumberFormat
        Case "CycleDateFormat":               CycleDateFormat
        Case "CycleCurrencyFormat":           CycleCurrencyFormat
        Case "CyclePercentFormat":            CyclePercentFormat
        Case "CycleOtherNumbers":             CycleOtherNumbers
        Case "IncreaseDecimal":               IncreaseDecimal
        Case "DecreaseDecimal":               DecreaseDecimal
        Case "ScaleUp":                       ScaleUp
        Case "ScaleDown":                     ScaleDown
        Case "ToggleSign":                    ToggleSign
        Case "DivideByHundred":               DivideByHundred
        Case "MultiplyByHundred":             MultiplyByHundred
        '---- Formulas
        Case "InsertCAGR":                    InsertCAGR
        Case "InsertPercentChange":           InsertPercentChange
        Case "EqualsLeft":                    EqualsLeft
        Case "ApplyGrowthRate":               ApplyGrowthRate
        Case "InsertQuickSum":                InsertQuickSum
        Case "InsertQuickAverage":            InsertQuickAverage
        '---- Styles
        Case "AutoColorSelection":            AutoColorSelection
        Case "CycleFont":                     CycleFont
        Case "CycleFill":                     CycleFill
        Case "CycleTextCase":                 CycleTextCase
        Case "CycleFontColor":                CycleFontColor
        Case "IncreaseFontSize":              IncreaseFontSize
        Case "DecreaseFontSize":              DecreaseFontSize
        Case "IndentIn":                      IndentIn
        Case "IndentOut":                     IndentOut
        Case "CenterAcrossSelection":         CenterAcrossSelection
        Case "InsertStaticNow":               InsertStaticNow
        Case "CycleInputStyle":               CycleInputStyle
        Case "CycleHeaderStyle":              CycleHeaderStyle
        Case "InsertHeadersFromPrompt":       InsertHeadersFromPrompt
        Case "InsertVarianceHeaders":         InsertVarianceHeaders
        Case "ApplyZeroCheckCF":              ApplyZeroCheckCF
        Case "ClearZeroCheckCF":              ClearZeroCheckCF
        Case "CycleTabColor":                 CycleTabColor
        Case "ApplyInputNexcess":             ApplyInputNexcess
        Case "ApplyInputBruce":               ApplyInputBruce
        Case "ApplyInputBanker":              ApplyInputBanker
        Case "ApplyInputGray":                ApplyInputGray
        Case "ApplyInputPeach":               ApplyInputPeach
        Case "ApplyInputOverride":            ApplyInputOverride
        '---- Borders
        Case "BorderTop":                     BorderTop
        Case "BorderBottom":                  BorderBottom
        Case "BorderLeft":                    BorderLeft
        Case "BorderRight":                   BorderRight
        Case "BordersOutlineInside":          BordersOutlineInside
        Case "ApplySumBar":                   ApplySumBar
        Case "ApplyDoubleSumBar":             ApplyDoubleSumBar
        '---- Unit tags
        Case "CycleUnitTag_Value_Uniform":    CycleUnitTag_Value_Uniform
        Case "CycleUnitTag_Duration_Uniform": CycleUnitTag_Duration_Uniform
        Case "CycleUnitTag_Rate_Uniform":     CycleUnitTag_Rate_Uniform
        Case "RemoveUnitTag":                 RemoveUnitTag
        '---- Undo (Ctrl+Z after an ILLXL command)
        Case "UndoLastAction":                UndoLastAction
        Case Else
            If cmd Like "pick:*:*" Then                 ' ribbon gallery: "pick:galNum:2"
                GalleryPick Split(cmd, ":")(1), CLng(Split(cmd, ":")(2))
            Else
                Err.Raise vbObjectError + 513, , "'" & cmd & "' is not an ILLXL command."
            End If
    End Select
    Exit Sub

Failed:
    Dim msg As String
    msg = FriendlyError(Err.Number, Err.Description)
    On Error Resume Next
    Application.ScreenUpdating = True                ' never leave the screen frozen
    Application.Cursor = xlDefault
    On Error GoTo 0
    LastErr = cmd & ": " & msg
    UiMsg IIf(cmd Like "pick:*", "That ribbon choice", "'" & CommandInfo(cmd)(1) & "'") & _
          " couldn't finish." & vbCrLf & vbCrLf & msg, vbExclamation
End Sub

' Plain-English versions of the errors people actually hit
Private Function FriendlyError(ByVal num As Long, ByVal desc As String) As String
    On Error Resume Next
    If num = 1004 And TypeName(ActiveSheet) = "Worksheet" Then     ' Excel refused a change...
        If ActiveSheet.ProtectContents Then                          ' ...because of protection
            FriendlyError = "This sheet is protected. Unprotect it first (Review > Unprotect Sheet)."
            Exit Function
        End If
    End If
    Select Case num
        Case 1004: FriendlyError = "Excel refused the change: " & desc
        Case 7, 14: FriendlyError = "Out of memory - try a smaller selection, or turn on Performance Mode."
        Case Else:  FriendlyError = desc & "  (error " & num & ")"
    End Select
End Function

'------------------------------------------------------------------------------
' Q02 - Messages and prompts
'------------------------------------------------------------------------------
Public Function UiMsg(ByVal text As String, Optional ByVal buttons As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    If TestMode Then                                 ' default = the "go ahead" button
        Shown = Shown & text & vbLf
        UiMsg = NextAnswer(IIf((buttons And 7) <= vbOKCancel, vbOK, vbYes))
    Else
        UiMsg = MsgBox(text, buttons, "ILLXL")
    End If
End Function

' Plain text box. Returns "" when cancelled.
Public Function UiInput(ByVal prompt As String, ByVal title As String, Optional ByVal default As String = "") As String
    If TestMode Then
        Shown = Shown & prompt & vbLf
        Dim a As Variant
        a = NextAnswer(default)
        If VarType(a) = vbBoolean Then UiInput = "" Else UiInput = CStr(a)   ' "cancel" -> ""
    Else
        UiInput = InputBox(prompt, "ILLXL - " & title, default)
    End If
End Function

' Excel's InputBox: kind 1 = number, 8 = cell/range, 9 = number OR range.
' Returns a Double, a Range, or False when cancelled. In test mode a range
' answer is an address on the active sheet ("B2", "Sheet2!A1:A5").
Public Function UiPick(ByVal prompt As String, ByVal title As String, ByVal kind As Long, _
                       Optional ByVal default As Variant) As Variant
    Dim a As Variant
    If TestMode Then
        Shown = Shown & prompt & vbLf
        a = NextAnswer(False)
        If VarType(a) = vbBoolean Then UiPick = False: Exit Function
        If IsNumeric(a) And kind <> 8 Then UiPick = CDbl(a): Exit Function
        If InStr(a, "!") > 0 Then
            Set UiPick = ActiveWorkbook.Worksheets(Replace(Split(a, "!")(0), "'", "")).Range(Split(a, "!")(1))
        Else
            Set UiPick = ActiveSheet.Range(a)
        End If
        Exit Function
    End If
    ' Passing the result straight into a Variant argument keeps it as-is - a
    ' Range stays a Range (a plain assignment would read its value instead).
    Keep Application.InputBox(prompt, "ILLXL - " & title, default, Type:=kind)
    If IsObject(Picked) Then
        Set UiPick = Picked
    ElseIf IsEmpty(Picked) Then
        UiPick = False
    Else
        UiPick = Picked                              ' a number, or False = cancelled
    End If
End Function

Private Sub Keep(ByVal v As Variant)
    If IsObject(v) Then Set Picked = v Else Picked = v
End Sub

'------------------------------------------------------------------------------
' Q03 - File dialogs (False when cancelled)
'------------------------------------------------------------------------------
Public Function UiOpenFile(ByVal filter As String, ByVal title As String) As Variant
    If TestMode Then UiOpenFile = NextAnswer(False) Else UiOpenFile = Application.GetOpenFilename(filter, , title)
End Function

Public Function UiSaveFile(ByVal initial As String, ByVal filter As String, ByVal title As String) As Variant
    If TestMode Then UiSaveFile = NextAnswer(False) Else UiSaveFile = Application.GetSaveAsFilename(initial, filter, , title)
End Function

'------------------------------------------------------------------------------
' Q04 - Shared guards
'------------------------------------------------------------------------------

' The part of rng that can hold anything - a whole-column selection becomes
' just its used rows, so loops stay fast. Nothing when it's all empty.
Public Function WorkCells(ByVal rng As Range) As Range
    Set WorkCells = Intersect(rng, rng.Worksheet.UsedRange)
End Function

' Cells for a formatting loop: the whole selection when it's a normal size;
' for a huge one (whole columns) just the used part, so it stays instant.
Public Function LoopCells(ByVal rng As Range) As Range
    If rng.Cells.CountLarge <= BIG_SELECTION Then Set LoopCells = rng: Exit Function
    Set LoopCells = WorkCells(rng)
    If LoopCells Is Nothing Then Set LoopCells = rng.Cells(1, 1)
End Function

' Cells a "write into every cell" command should fill. A normal selection is
' used as-is; a huge one (whole columns / rows) becomes the used part of the
' sheet - selecting column A means "the rows I have", not a million empty cells.
Public Function WriteCells(ByVal rng As Range) As Range
    If rng.Cells.CountLarge <= BIG_SELECTION Then
        Set WriteCells = rng
    Else                                             ' the sheet's used ROWS, in the selected columns
        Set WriteCells = Intersect(rng, rng.Worksheet.UsedRange.EntireRow)
    End If
End Function

' True if it's OK to write into every cell of rng (asks above BIG_SELECTION)
Public Function ConfirmBig(ByVal rng As Range, ByVal what As String) As Boolean
    If rng.Cells.CountLarge <= BIG_SELECTION Then ConfirmBig = True: Exit Function
    ConfirmBig = (UiMsg(what & " will write to " & Format$(rng.Cells.CountLarge, "#,##0") & _
                        " cells. Continue?", vbOKCancel + vbQuestion) = vbOK)
End Function

' Address of target as seen from a formula in host: adds the sheet (and
' workbook) name when target lives somewhere else.
Public Function RefFrom(ByVal host As Range, ByVal target As Range, Optional ByVal absolute As Boolean = False) As String
    If target.Worksheet.Parent Is host.Worksheet.Parent Then
        RefFrom = target.Address(absolute, absolute)
        If Not target.Worksheet Is host.Worksheet Then RefFrom = "'" & Replace(target.Worksheet.Name, "'", "''") & "'!" & RefFrom
    Else
        RefFrom = target.Address(absolute, absolute, External:=True)
    End If
End Function

' A number as formula text - always "." for decimals, whatever the PC's locale
Public Function NumText(ByVal v As Double) As String
    NumText = Trim$(Str$(v))
    If Left$(NumText, 1) = "." Then NumText = "0" & NumText
    If Left$(NumText, 2) = "-." Then NumText = "-0" & Mid$(NumText, 2)
End Function

'------------------------------------------------------------------------------
' Q05 - Test mode
'------------------------------------------------------------------------------
Public Sub UiTestMode(ByVal onOff As Boolean)
    TestMode = onOff
    Set Answers = New Collection
    Shown = ""
End Sub

' Queue the next answer(s). "cancel" = the user cancels.
Public Sub UiAnswer(ByVal a As Variant)
    If Answers Is Nothing Then Set Answers = New Collection
    If VarType(a) = vbString Then
        If LCase$(a) = "cancel" Then a = False
    End If
    Answers.Add a
End Sub

Public Function UiShown() As String
    UiShown = "" & Shown
End Function

Public Function UiLastError() As String
    UiLastError = "" & LastErr
End Function

Public Function UiLastCommand() As String
    UiLastCommand = "" & LastCmd
End Function

Public Sub UiResetLastCommand()
    LastCmd = ""
End Sub

Public Function UiIsTest() As Boolean
    UiIsTest = TestMode
End Function

Private Function NextAnswer(ByVal fallback As Variant) As Variant
    If Answers Is Nothing Then NextAnswer = fallback: Exit Function
    If Answers.Count = 0 Then NextAnswer = fallback: Exit Function
    NextAnswer = Answers(1)
    Answers.Remove 1
End Function
