Attribute VB_Name = "modStyles"
Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+
 
'==============================================================================
' Module: modStyles (Sxx)
' Purpose:
'   - Font/fill/text/font-color cycles
'   - Input styles & header styles
'   - Layout helpers (indent, center across)
'   - Zero-check conditional formatting
'
' Bound shortcuts (see modBindings):
'   S01 - AutoColorSelection      Ctrl+Alt+A
'   S02 - CycleFont               Ctrl+'
'   S03 - CycleFill               Ctrl+Shift+K
'   S04 - CycleTextCase           Ctrl+Alt+Shift+I
'   S05 - CycleFontColor          Ctrl+Shift+C
'   S06 - CycleInputStyle         Ctrl+Alt+Shift+U
'   S07 - CycleHeaderStyle        Ctrl+Alt+Shift+H
'   S08 - InsertHeadersFromPrompt Ctrl+Alt+Shift+Y
'   S09 - InsertVarianceHeaders   Ctrl+Alt+Shift+D
'   S10 - CenterAcrossSelection   Ctrl+Alt+E
'   S11 - IncreaseFontSize        Ctrl+Shift+F
'   S12 - DecreaseFontSize        Ctrl+Shift+G
'   S13 - IndentIn                Ctrl+Alt+.
'   S14 - IndentOut               Ctrl+Alt+,
'   S15 - InsertStaticNow         Ctrl+Shift+N
'   S16 - ApplyZeroCheckCF        Ctrl+Alt+Shift+Z
'   S17 - ClearZeroCheckCF        Ctrl+Alt+Shift+X
'   (ZoomIn/ZoomOut commented out - use native Excel zoom)
'==============================================================================

'------------------------------------------------------------------------------
' S01 - AutoColorSelection  (Ctrl+Alt+A)
'     Auto-colors cells: blue=inputs, black=formulas, red=external links,
'     green=cross-sheet refs, purple=same-sheet refs
'------------------------------------------------------------------------------
Sub AutoColorSelection()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Application.ScreenUpdating = False
    Dim rF As Range, rN As Range
    ' SpecialCellsIn keeps a single-cell selection from recoloring the whole sheet
    Set rF = SpecialCellsIn(Selection, xlCellTypeFormulas)
    Set rN = SpecialCellsIn(Selection, xlCellTypeConstants, xlNumbers)

    If Not rF Is Nothing Then Set rF = SpecialCellsIn(rF, xlCellTypeVisible)
    If Not rN Is Nothing Then Set rN = SpecialCellsIn(rN, xlCellTypeVisible)
    Dim rLinks As Range, rSheet As Range, rOps As Range, rPlain As Range, c As Range, f As String
    Dim links As Variant
    links = ExternalLinkFiles(ActiveWorkbook)
    If Not rF Is Nothing Then
        For Each c In rF.Cells
            f = c.Formula
            If IsExternalFormula(f, links) Then
                Set rLinks = AddToUnion(rLinks, c)
            ElseIf InStr(f, "!") > 0 Then
                Set rSheet = AddToUnion(rSheet, c)
            ElseIf InStr(f, "+") > 0 Or InStr(f, "-") > 0 Or InStr(f, "*") > 0 Or InStr(f, "/") > 0 Or InStr(f, "^") > 0 Then
                Set rOps = AddToUnion(rOps, c)
            Else
                Set rPlain = AddToUnion(rPlain, c)
            End If
        Next c
    End If
    ' Colors come from the preset's [AutoColor] section (Nexcess defaults shown)
    If Not rLinks Is Nothing Then rLinks.Font.Color = cfgClrExternal       ' red = external link
    If Not rSheet Is Nothing Then rSheet.Font.Color = cfgClrCrossSheet     ' green = cross-sheet ref
    If Not rOps Is Nothing Then rOps.Font.Color = cfgClrFormula            ' black = formula
    If Not rPlain Is Nothing Then rPlain.Font.Color = cfgClrSameSheet      ' purple = same-sheet ref
    If Not rN Is Nothing Then rN.Font.Color = cfgClrHardcode               ' soft blue = hardcode
    Application.ScreenUpdating = True
    LogAction "AutoColor", Selection.Address(False, False)
    RegisterUndo "Auto Color"
End Sub



Private Function AddToUnion(ByVal acc As Range, ByVal c As Range) As Range
    If acc Is Nothing Then
        Set AddToUnion = c
    Else
        Set AddToUnion = Application.Union(acc, c)
    End If
End Function

'------------------------------------------------------------------------------
' ExternalLinkFiles / IsExternalFormula
'     Uses the workbook's own link list (Data > Edit Links) instead of
'     looking for "[" - table refs like tblSales[Amount] also contain "[".
'     A formula is external if it names a linked file or a web address.
'------------------------------------------------------------------------------
Private Function ExternalLinkFiles(ByVal wb As Workbook) As Variant
    Dim src As Variant, i As Long, p As String, cut As Long
    src = wb.LinkSources(xlExcelLinks)              ' full paths/URLs, or Empty
    If IsEmpty(src) Then Exit Function
    For i = LBound(src) To UBound(src)
        p = src(i)
        cut = InStrRev(p, "\"): If InStrRev(p, "/") > cut Then cut = InStrRev(p, "/")
        src(i) = Mid$(p, cut + 1)                   ' file name only, e.g. Budget.xlsx
    Next i
    ExternalLinkFiles = src
End Function

Private Function IsExternalFormula(ByVal f As String, ByVal links As Variant) As Boolean
    Dim i As Long
    If InStr(1, f, "http", vbTextCompare) > 0 Then IsExternalFormula = True: Exit Function
    If IsEmpty(links) Then Exit Function
    For i = LBound(links) To UBound(links)
        If InStr(1, f, links(i), vbTextCompare) > 0 Then IsExternalFormula = True: Exit Function
    Next i
End Function

'------------------------------------------------------------------------------
' S02 - CycleFont  (Ctrl+')     list: preset [Fonts]
'------------------------------------------------------------------------------
Sub CycleFont()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevFontAddress Then FontCycleIndex = 0
    PrevFontAddress = Selection.Address(False, False)

    Dim ni As Long
    ni = FontCycleIndex Mod (UBound(cfgFonts) + 1)

    On Error Resume Next
    Selection.Font.Name = cfgFonts(ni)
    If Err.Number <> 0 Then
        MsgBox "Font '" & cfgFonts(ni) & "' not available. Using Calibri.", vbInformation
        Selection.Font.Name = "Calibri"
        Err.Clear
    End If
    On Error GoTo 0

    FontCycleIndex = FontCycleIndex + 1
    LogAction "FontCyc" & (ni + 1), PrevFontAddress
    RegisterUndo "Font Cycle"
End Sub

'------------------------------------------------------------------------------
' S03 - CycleFill  (Ctrl+Shift+K)     list: preset [Fills]
'     "none" clears the fill, "stripe" is a light-gray diagonal stripe.
'------------------------------------------------------------------------------
Sub CycleFill()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevFillAddress Then FillCycleIndex = 0
    PrevFillAddress = Selection.Address(False, False)
    Dim ni As Long
    ni = FillCycleIndex Mod (UBound(cfgFills) + 1)
    Select Case cfgFills(ni).Color
        Case CLR_NONE
            Selection.Interior.Pattern = xlNone
        Case CLR_STRIPE
            With Selection.Interior
                .Pattern = xlUp
                .PatternColor = RGB(191, 191, 191)
                .Color = RGB(255, 255, 255)
            End With
        Case Else
            With Selection.Interior
                .Pattern = xlSolid
                .Color = cfgFills(ni).Color
            End With
    End Select
    FillCycleIndex = FillCycleIndex + 1
    LogAction "FillCyc" & (ni + 1), PrevFillAddress
    RegisterUndo "Fill Cycle"
End Sub

'------------------------------------------------------------------------------
' S04 - CycleTextCase  (Ctrl+Alt+Shift+I)
'------------------------------------------------------------------------------
Sub CycleTextCase()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevTextAddress Then TextCaseIndex = 0
    PrevTextAddress = Selection.Address(False, False)
    Dim modes As Variant, ni As Long, c As Range
    modes = Array(vbProperCase, vbLowerCase, vbUpperCase)
    ni = TextCaseIndex Mod (UBound(modes) + 1)
    ' Typed text only - formulas, numbers and dates are left untouched
    Dim txt As Range
    Set txt = SpecialCellsIn(Selection, xlCellTypeConstants, xlTextValues)
    If Not txt Is Nothing Then
        For Each c In txt.Cells: c.Value = VBA.StrConv(c.Value, modes(ni)): Next c
    End If
    TextCaseIndex = TextCaseIndex + 1
    LogAction "TextCase" & (ni + 1), PrevTextAddress
    RegisterUndo "Text Case"
End Sub

'------------------------------------------------------------------------------
' S05 - CycleFontColor  (Ctrl+Shift+C)     list: preset [FontColors]
'------------------------------------------------------------------------------
Sub CycleFontColor()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevFontColorAddress Then FontColorCycleIndex = 0
    PrevFontColorAddress = Selection.Address(False, False)
    Dim ni As Long
    ni = FontColorCycleIndex Mod (UBound(cfgFontColors) + 1)
    Selection.Font.Color = cfgFontColors(ni).Color
    FontColorCycleIndex = FontColorCycleIndex + 1
    LogAction "FontColorCyc" & (ni + 1), PrevFontColorAddress
    RegisterUndo "Font Color Cycle"
End Sub

'------------------------------------------------------------------------------
' DISABLED - ZoomIn/ZoomOut - Use native Excel zoom (Ctrl+scroll or View > Zoom)
'------------------------------------------------------------------------------
'Sub ZoomIn()
'    BeginMacroWithUndo
'    With ActiveWindow: .Zoom = Application.Min(.Zoom + 10, 400): End With
'    LogAction "ZoomIn", ""
'    RegisterUndo "Zoom In"
'End Sub

'Sub ZoomOut()
'    BeginMacroWithUndo
'    With ActiveWindow: .Zoom = Application.Max(.Zoom - 10, 10): End With
'    LogAction "ZoomOut", ""
'    RegisterUndo "Zoom Out"
'End Sub


'------------------------------------------------------------------------------
' S06 - CycleInputStyle  (Ctrl+Alt+Shift+U)     list: preset [InputStyles]
'     Cycles through color-coded input styles. Resets on selection change.
'     Each style = fill + font color + a hairline dotted border, so input
'     cells read as a group without heavy formatting noise.
'
'     Nexcess preset:
'     STEP  STYLE          FILL     FONT        TYPICAL USE
'     ----  -------------  -------  ----------  -------------------------------
'     1     Nexcess        #FFFFCC  Soft Blue   Nexcess primary assumption
'     2     Bruce Input    #FFF2CC  Blue        Signature warm yellow hardcode
'     3     Banker Input   #FFFFCC  Blue        Standard banker hardcode style
'     4     Gray           #D9D9D9  Blue        Linked / locked / non-editable
'     5     FMWC Peach     #FFCC99  Teal        Flagged / exception inputs
'     6     Override       #DDEBF7  Dark Navy   Override cells / model exceptions
'------------------------------------------------------------------------------
Sub CycleInputStyle()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Static idx As Integer, prev As String
    If Selection.Address(False, False) <> prev Then idx = 0
    prev = Selection.Address(False, False)
    ApplyInputStyle Selection, cfgInputStyles(idx Mod (UBound(cfgInputStyles) + 1))
    idx = idx + 1
    LogAction "InputStyle", prev
    RegisterUndo "Input Style"
End Sub

'==============================================================================
' INDIVIDUAL INPUT STYLE HELPERS
'==============================================================================
' Apply one input style directly, without cycling - for ribbon buttons, the
' Quick Access Toolbar, or other macros. Each looks its style up BY NAME in
' the active preset, so it keeps working if you reorder the cycle. If the
' preset has no style with that name, Excel just beeps. Undoable via Ctrl+Z.
'
'   ApplyInputNexcess   "Nexcess"       primary hardcoded drivers (Nexcess models)
'   ApplyInputBruce     "Bruce Input"   hardcodes in personal / non-Nexcess models
'   ApplyInputBanker    "Banker Input"  classic IB / WSP / Macabacus hardcode
'   ApplyInputGray      "Gray"          linked / locked - "do not edit"
'   ApplyInputPeach     "FMWC Peach"    temporary overrides, outliers, follow-ups
'   ApplyInputOverride  "Override"      formula deliberately replaced by a value
'==============================================================================
Sub ApplyInputNexcess()
    ApplyInputNamed "Nexcess"
End Sub
Sub ApplyInputBruce()
    ApplyInputNamed "Bruce Input"
End Sub
Sub ApplyInputBanker()
    ApplyInputNamed "Banker Input"
End Sub
Sub ApplyInputGray()
    ApplyInputNamed "Gray"
End Sub
Sub ApplyInputPeach()
    ApplyInputNamed "FMWC Peach"
End Sub
Sub ApplyInputOverride()
    ApplyInputNamed "Override"
End Sub

Private Sub ApplyInputNamed(ByVal styleName As String)
    Dim i As Long
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    For i = 0 To UBound(cfgInputStyles)
        If StrComp(cfgInputStyles(i).Name, styleName, vbTextCompare) = 0 Then
            ApplyInputStyle Selection, cfgInputStyles(i)
            LogAction "Input " & styleName, Selection.Address(False, False)
            RegisterUndo "Input " & styleName
            Exit Sub
        End If
    Next i
    Beep                                            ' not in this preset
End Sub

'==============================================================================
' CORE STYLING ENGINES  (Private)
'==============================================================================
' ApplyInputStyle  - fill + font color + hairline dotted border on all six
'                    border positions (edges and inside lines), so single
'                    cells and blocks both look right.
' ApplyHeaderStyle - fill + font color + bold. Number formats are left alone.
'==============================================================================
Private Sub ApplyInputStyle(ByVal rng As Range, ByRef sty As StyleItem)
    Dim parts As Variant, i As Long
    If rng Is Nothing Then Exit Sub
    rng.Font.Color = sty.Font
    With rng.Interior
        .Pattern = xlSolid
        .Color = sty.Fill
        .TintAndShade = 0
    End With
    rng.Borders.LineStyle = xlNone
    parts = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                  xlInsideVertical, xlInsideHorizontal)
    For i = LBound(parts) To UBound(parts)
        With rng.Borders(parts(i))
            .LineStyle = xlDot
            .Weight = xlHairline
            .Color = sty.Border
            .TintAndShade = 0
        End With
    Next i
End Sub

Private Sub ApplyHeaderStyle(ByVal rng As Range, ByVal fillColor As Long, ByVal fontColor As Long)
    With rng
        .Interior.Pattern = xlSolid
        .Interior.Color = fillColor
        .Font.Color = fontColor
        .Font.Bold = True
    End With
End Sub

'------------------------------------------------------------------------------
' S07 - CycleHeaderStyle  (Ctrl+Alt+Shift+H)     list: preset [HeaderStyles]
'     Nexcess preset: Nexcess Navy > Mid Navy > Black > Blue > Cream
'------------------------------------------------------------------------------
Sub CycleHeaderStyle()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Static idx As Integer, prev As String
    If Selection.Address(False, False) <> prev Then idx = 0
    prev = Selection.Address(False, False)
    With cfgHeaderStyles(idx Mod (UBound(cfgHeaderStyles) + 1))
        ApplyHeaderStyle Selection, .Fill, .Font
    End With
    idx = idx + 1
    LogAction "HeaderStyle", prev
    RegisterUndo "Header Style"
End Sub

'------------------------------------------------------------------------------
' S08 - InsertHeadersFromPrompt  (Ctrl+Alt+Shift+Y)
'     Prompt default and header colors: preset [Headers]
'------------------------------------------------------------------------------
Sub InsertHeadersFromPrompt()
    If TypeName(Selection) <> "Range" Then Exit Sub
    EnsureSettings
    Dim s As String, parts() As String, i As Long, hdr As Range
    s = InputBox("Enter comma-separated headers (e.g., 2024A,2025B,2026E):", _
                 "Insert Headers", cfgHeaderPrompt)
    If Len(Trim$(s)) = 0 Then Exit Sub
    parts = Split(s, ",")
    ' Headers can run past the selection - snapshot the cells actually written
    Set hdr = Selection.Cells(1, 1).Resize(1, UBound(parts) - LBound(parts) + 1)
    BeginMacroWithUndo hdr
    For i = LBound(parts) To UBound(parts)
        hdr.Cells(1, i + 1).Value = Trim$(parts(i))
    Next i
    ApplyHeaderStyle hdr, cfgHdrInsertFill, cfgHdrInsertFont
    LogAction "InsertHeaders", Selection.Address(False, False)
    RegisterUndo "Insert Headers"
End Sub

'------------------------------------------------------------------------------
' S09 - InsertVarianceHeaders  (Ctrl+Alt+Shift+D)
'     Labels: preset [Headers] VarianceLabels (Nexcess: AvF % / AvB% / Var AvB)
'------------------------------------------------------------------------------
Sub InsertVarianceHeaders()
    If TypeName(Selection) <> "Range" Then Exit Sub
    EnsureSettings
    Dim hdr As Range
    ' one cell per label, starting at the selection - even from a 1-cell selection
    Set hdr = Selection.Cells(1, 1).Resize(1, UBound(cfgVarianceLabels) + 1)
    BeginMacroWithUndo hdr
    hdr.Value = cfgVarianceLabels
    ApplyHeaderStyle hdr, cfgHdrInsertFill, cfgHdrInsertFont
    LogAction "InsertVarHdrs", Selection.Address(False, False)
    RegisterUndo "Insert Variance Headers"
End Sub

'------------------------------------------------------------------------------
' S10 - CenterAcrossSelection  (Ctrl+Alt+E)
'------------------------------------------------------------------------------
Sub CenterAcrossSelection()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Selection.HorizontalAlignment = xlCenterAcrossSelection
    LogAction "CenterAcross", Selection.Address(False, False)
    RegisterUndo "Center Across"
End Sub

'------------------------------------------------------------------------------
' S11 - IncreaseFontSize  (Ctrl+Shift+F)
'------------------------------------------------------------------------------
Sub IncreaseFontSize()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    BumpFontSize 1
    LogAction "FontInc", Selection.Address(False, False)
    RegisterUndo "Increase Font Size"
End Sub

'------------------------------------------------------------------------------
' S12 - DecreaseFontSize  (Ctrl+Shift+G)
'------------------------------------------------------------------------------
Sub DecreaseFontSize()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    BumpFontSize -1
    LogAction "FontDec", Selection.Address(False, False)
    RegisterUndo "Decrease Font Size"
End Sub

'------------------------------------------------------------------------------
' BumpFontSize - shared by S11/S12
'     One size across the selection: change it in one shot.
'     Mixed sizes (Font.Size returns Null): step each used cell from its own size.
'------------------------------------------------------------------------------
Private Sub BumpFontSize(ByVal delta As Long)
    Dim rng As Range, c As Range, sz As Variant
    sz = Selection.Font.Size
    If Not IsNull(sz) Then
        Selection.Font.Size = Application.Max(1, sz + delta)
        Exit Sub
    End If
    Set rng = Intersect(Selection, Selection.Worksheet.UsedRange)
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        sz = c.Font.Size
        If IsNull(sz) Then sz = c.Characters(1, 1).Font.Size    ' rich text inside one cell
        c.Font.Size = Application.Max(1, sz + delta)
    Next c
End Sub

'------------------------------------------------------------------------------
' S13 - IndentIn  (Ctrl+Alt+.)
'------------------------------------------------------------------------------
Sub IndentIn()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection.Cells
        If c.HorizontalAlignment = xlCenter Or _
           c.HorizontalAlignment = xlRight Or _
           c.HorizontalAlignment = xlCenterAcrossSelection Then
            c.HorizontalAlignment = xlLeft
        End If
        c.IndentLevel = Application.Min(c.IndentLevel + 1, 15)
    Next c
    LogAction "IndentIn", Selection.Address(False, False)
    RegisterUndo "Indent In"
End Sub

'------------------------------------------------------------------------------
' S14 - IndentOut  (Ctrl+Alt+,)
'------------------------------------------------------------------------------
Sub IndentOut()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection.Cells
        c.IndentLevel = Application.Max(c.IndentLevel - 1, 0)
        If c.IndentLevel = 0 Then
            c.HorizontalAlignment = xlGeneral
        End If
    Next c
    LogAction "IndentOut", Selection.Address(False, False)
    RegisterUndo "Indent Out"
End Sub

'------------------------------------------------------------------------------
' S15 - InsertStaticNow  (Ctrl+Shift+N)     formats: preset [TimestampFormats]
'     Press again on the same cells to cycle the format (plain -> EST suffix).
'------------------------------------------------------------------------------
Sub InsertStaticNow()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Static idx As Integer, prev As String
    If Selection.Address(False, False) <> prev Then idx = 0
    prev = Selection.Address(False, False)

    Dim c As Range
    Dim fmt As String
    fmt = cfgStampFormats(idx Mod (UBound(cfgStampFormats) + 1))

    For Each c In Selection
        c.Value = Now
        c.NumberFormat = fmt
    Next c

    idx = idx + 1
    LogAction "InsertNow:" & fmt, Selection.Address(False, False)
    RegisterUndo "Insert Static Now"
End Sub


'------------------------------------------------------------------------------
' S16 - ApplyZeroCheckCF  (Ctrl+Alt+Shift+Z)
'     Applies conditional formatting: green if =0, red if <>0
'------------------------------------------------------------------------------
Sub ApplyZeroCheckCF()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim area As Range
    Application.ScreenUpdating = False
    For Each area In Selection.Areas
        ApplyZeroCheckToArea area
    Next area
    Application.ScreenUpdating = True
    LogAction "ZeroCheckCF_Apply", Selection.Address(False, False)
    RegisterUndo "Zero-Check CF"
End Sub

Private Sub ApplyZeroCheckToArea(ByVal rng As Range)
    Dim tl As Range, fEq As String, fNe As String, i As Long
    Set tl = rng.Cells(1, 1)
    fEq = "=" & tl.Address(False, False) & "=0"
    fNe = "=" & tl.Address(False, False) & "<>0"
    For i = rng.FormatConditions.count To 1 Step -1
        With rng.FormatConditions(i)
            If .Type = xlExpression Then
                If .Formula1 = fEq Or .Formula1 = fNe Then .Delete
            End If
        End With
    Next i
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fEq)
        .Interior.Color = cfgClrZeroOK              ' preset [ZeroCheck] Zero
        .StopIfTrue = True
    End With
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fNe)
        .Interior.Color = cfgClrZeroBad             ' preset [ZeroCheck] NonZero
    End With
End Sub

'------------------------------------------------------------------------------
' S17 - ClearZeroCheckCF  (Ctrl+Alt+Shift+X)
'------------------------------------------------------------------------------
Sub ClearZeroCheckCF()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim area As Range, tl As Range, fEq As String, fNe As String, i As Long
    Application.ScreenUpdating = False
    For Each area In Selection.Areas
        Set tl = area.Cells(1, 1)
        fEq = "=" & tl.Address(False, False) & "=0"
        fNe = "=" & tl.Address(False, False) & "<>0"
        For i = area.FormatConditions.count To 1 Step -1
            With area.FormatConditions(i)
                If .Type = xlExpression Then
                    If .Formula1 = fEq Or .Formula1 = fNe Then .Delete
                End If
            End With
        Next i
    Next area
    Application.ScreenUpdating = True
    LogAction "ZeroCheckCF_Clear", Selection.Address(False, False)
    RegisterUndo "Clear Zero-Check CF"
End Sub


'--------------------------------------------------------------------
' S18 - CycleTabColor  (Ctrl+Alt+Shift+W)     list: preset [TabColors]
'     Nexcess preset: none > Blue (output) > Yellow (inputs) > Green (calcs)
'                     > Purple (pivots/PQ) > Red (review) > Gray (archive)
'--------------------------------------------------------------------
Sub CycleTabColor()
    BeginMacroWithUndo ActiveCell                  ' snapshot = tab color (+ one cell)
    Static idx As Integer, prev As String
    If ActiveSheet.Name <> prev Then idx = 0
    prev = ActiveSheet.Name

    With cfgTabColors(idx Mod (UBound(cfgTabColors) + 1))
        If .Color = CLR_NONE Then
            ActiveSheet.Tab.ColorIndex = xlColorIndexNone
        Else
            ActiveSheet.Tab.Color = .Color
        End If
    End With

    idx = idx + 1
    RegisterUndo "Tab Color"
End Sub





