Attribute VB_Name = "modFormatCycles"
Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Module: modFormatCycles (Fxx)
' Purpose:
'   Number/date/percent/currency/other cycles, decimals, scale, and sign toggle.
'
' Bound shortcuts (see modBindings):
'   F01 - CycleNumberFormat      Ctrl+Shift+1
'   F02 - CycleDateFormat        Ctrl+Shift+3
'   F03 - CyclePercentFormat     Ctrl+Shift+5
'   F04 - CycleCurrencyFormat    Ctrl+Shift+4
'   F05 - CycleOtherNumbers      Ctrl+Shift+8
'   F06 - IncreaseDecimal        Ctrl+Shift+.
'   F07 - DecreaseDecimal        Ctrl+Shift+,
'   F08 - ScaleUp                Alt+Shift+<
'   F09 - ScaleDown              Alt+Shift+>
'   F10 - ToggleSign             Ctrl+Alt+Shift+\
'   F11 - DivideByHundred        Ctrl+Alt+2
'   F12 - MultiplyByHundred      Ctrl+Alt+Shift+2
'==============================================================================

'------------------------------------------------------------------------------
' F01-F05 - Format cycles
'     Each press applies the next format from the active preset's list
'     ([NumberFormats], [DateFormats], ... in the .ini). The cycle restarts
'     at step 1 whenever the selection changes.
'------------------------------------------------------------------------------

' F01 CycleNumberFormat  (Ctrl+Shift+1)
Sub CycleNumberFormat()
    ApplyFormatCycle "Num", FormatNumIndex, PrevNumAddress, "NumFmt", "Number Format"
End Sub

' F02 CycleDateFormat  (Ctrl+Shift+3)
Sub CycleDateFormat()
    ApplyFormatCycle "Date", FormatDateIndex, PrevDateAddress, "DateFmt", "Date Format"
End Sub

' F03 CyclePercentFormat  (Ctrl+Shift+5)
Sub CyclePercentFormat()
    ApplyFormatCycle "Pct", FormatPctIndex, PrevPctAddress, "PctFmt", "Percent Format"
End Sub

' F04 CycleCurrencyFormat  (Ctrl+Shift+4)
Sub CycleCurrencyFormat()
    ApplyFormatCycle "Cur", FormatCurIndex, PrevCurAddress, "CurFmt", "Currency Format"
End Sub

' F05 CycleOtherNumbers  (Ctrl+Shift+8)
Sub CycleOtherNumbers()
    ApplyFormatCycle "Other", FormatOtherIndex, PrevOtherAddress, "OtherFmt", "Other Numbers Format"
End Sub

' F01-F05 direct pick (ribbon galleries): apply step i of a cycle. The next
' key press carries on from the step after it.
Public Sub ApplyFormatAt(ByVal which As String, ByVal i As Long)
    Select Case which
        Case "Num":   ApplyFormatCycle which, FormatNumIndex, PrevNumAddress, "NumFmt", "Number Format", i
        Case "Date":  ApplyFormatCycle which, FormatDateIndex, PrevDateAddress, "DateFmt", "Date Format", i
        Case "Pct":   ApplyFormatCycle which, FormatPctIndex, PrevPctAddress, "PctFmt", "Percent Format", i
        Case "Cur":   ApplyFormatCycle which, FormatCurIndex, PrevCurAddress, "CurFmt", "Currency Format", i
        Case "Other": ApplyFormatCycle which, FormatOtherIndex, PrevOtherAddress, "OtherFmt", "Other Numbers Format", i
    End Select
End Sub

' Shared engine for F01-F05. The list is looked up after BeginMacroWithUndo,
' which loads the preset if needed - so a preset switch applies immediately.
' pick >= 0 jumps straight to that step (ribbon gallery).
Private Sub ApplyFormatCycle(ByVal which As String, ByRef index As Integer, ByRef prevAddr As String, _
                             ByVal logName As String, ByVal undoName As String, Optional ByVal pick As Long = -1)
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim fmts As Variant, ni As Long
    Select Case which
        Case "Num":   fmts = cfgNumFormats
        Case "Date":  fmts = cfgDateFormats
        Case "Pct":   fmts = cfgPctFormats
        Case "Cur":   fmts = cfgCurFormats
        Case "Other": fmts = cfgOtherFormats
    End Select
    If pick >= 0 Then
        If pick > UBound(fmts) Then Exit Sub
        index = pick
    ElseIf Selection.Address(False, False) <> prevAddr Then
        index = 0
    End If
    prevAddr = Selection.Address(False, False)
    ni = index Mod (UBound(fmts) + 1)
    Selection.NumberFormat = fmts(ni)
    index = index + 1
    LogAction logName & (ni + 1), prevAddr
    RegisterUndo undoName
End Sub

'------------------------------------------------------------------------------
' F06 IncreaseDecimal (Ctrl+Shift+.)
'------------------------------------------------------------------------------
Sub IncreaseDecimal()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Application.ScreenUpdating = False
    AdjustDecimalsInSelection Selection, 1
    Application.ScreenUpdating = True
    LogAction "IncreaseDecimal", Selection.Address(False, False)
    RegisterUndo "Increase Decimal"
End Sub

'------------------------------------------------------------------------------
' F07 DecreaseDecimal (Ctrl+Shift+,)
'------------------------------------------------------------------------------
Sub DecreaseDecimal()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Application.ScreenUpdating = False
    AdjustDecimalsInSelection Selection, -1
    Application.ScreenUpdating = True
    LogAction "DecreaseDecimal", Selection.Address(False, False)
    RegisterUndo "Decrease Decimal"
End Sub

' Helper: cache formats so we don't recompute for every single cell
Private Sub AdjustDecimalsInSelection(ByVal rng As Range, ByVal delta As Long)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim c As Range, fmt As String, newFmt As String

    For Each c In LoopCells(rng).Cells               ' whole columns: used cells only
        fmt = CStr(c.NumberFormat)
        If Not dict.Exists(fmt) Then dict(fmt) = AdjustFormatDecimals(fmt, delta)
        newFmt = dict(fmt)
        If c.NumberFormat <> newFmt Then c.NumberFormat = newFmt
    Next c
End Sub

'------------------------------------------------------------------------------
' F08-F12 - Scale, sign, x100
'     Numbers are changed in place; formulas are wrapped, e.g. =(A1+B1)/1000.
'     Blanks, text, dates and errors are left alone. Only used cells are
'     visited, so a whole-column selection is instant.
'------------------------------------------------------------------------------

' F08 ScaleUp  (Alt+Shift+<) - divide by 1,000 (show in thousands)
Sub ScaleUp()
    TransformNumbers "/1000", "Scale Up"
End Sub

' F09 ScaleDown  (Alt+Shift+>) - multiply by 1,000
Sub ScaleDown()
    TransformNumbers "*1000", "Scale Down"
End Sub

' F10 ToggleSign  (Ctrl+Alt+Shift+\) - positive <-> negative; a formula
'     wrapped as =-(...) is unwrapped again on the next press
Sub ToggleSign()
    TransformNumbers "neg", "Toggle Sign"
End Sub

' F11 DivideByHundred  (Ctrl+Alt+2) - e.g. 5 -> 0.05 for 5%
Sub DivideByHundred()
    TransformNumbers "/100", "Divide by 100"
End Sub

' F12 MultiplyByHundred  (Ctrl+Alt+Shift+2) - e.g. 0.05 -> 5
Sub MultiplyByHundred()
    TransformNumbers "*100", "Multiply by 100"
End Sub

' Shared engine for F08-F12. op = "/1000", "*1000", "/100", "*100" or "neg".
Private Sub TransformNumbers(ByVal op As String, ByVal undoName As String)
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range, work As Range, f As String
    Set work = WorkCells(Selection)                 ' used cells only
    If Not work Is Nothing Then
        Application.ScreenUpdating = False
        For Each c In work.Cells
            If c.HasFormula Then
                f = FormulaOf(c)                    ' Formula2: keeps dynamic arrays intact
                If op <> "neg" Then
                    f = "=(" & Mid$(f, 2) & ")" & op
                ElseIf Left$(f, 3) = "=-(" And Right$(f, 1) = ")" Then
                    f = "=" & Mid$(f, 4, Len(f) - 4)                ' unwrap
                Else
                    f = "=-(" & Mid$(f, 2) & ")"                    ' wrap
                End If
                SetFormula c, f
            ElseIf IsNumberCell(c) Then
                Select Case op
                    Case "/1000": c.Value = c.Value / 1000
                    Case "*1000": c.Value = c.Value * 1000
                    Case "/100":  c.Value = c.Value / 100
                    Case "*100":  c.Value = c.Value * 100
                    Case "neg":   c.Value = -c.Value
                End Select
            End If
        Next c
        Application.ScreenUpdating = True
    End If
    LogAction Replace(undoName, " ", ""), Selection.Address(False, False)
    RegisterUndo undoName
End Sub

'------------------------------------------------------------------------------
' AdjustFormatDecimals - Split multi-section format by ";" and adjust each
' section independently, then rejoin. This correctly handles formats like:
'   #,##0_);(#,##0);""--"";@
'------------------------------------------------------------------------------
Private Function AdjustFormatDecimals(ByVal fmt As String, ByVal delta As Long) As String
    ' Split on ";" but only outside quoted strings
    Dim sections() As String
    sections = SplitFormatSections(fmt)
    
    Dim i As Long
    For i = 0 To UBound(sections)
        ' Don't touch text/placeholder sections like "--", "@", ""--""
        If Not IsPureLiteralSection(sections(i)) Then
            sections(i) = AdjustSectionDecimalsOne(sections(i), delta)
        End If
    Next i
    
    AdjustFormatDecimals = Join(sections, ";")
End Function

'------------------------------------------------------------------------------
' SplitFormatSections - Split format string by ";" while respecting quoted text
'------------------------------------------------------------------------------
Private Function SplitFormatSections(ByVal fmt As String) As String()
    Dim result(0 To 3) As String
    Dim count As Long, i As Long, ch As String
    Dim inQuote As Boolean, buf As String
    
    count = 0
    inQuote = False
    buf = ""
    
    For i = 1 To Len(fmt)
        ch = Mid$(fmt, i, 1)
        If ch = """" Then
            inQuote = Not inQuote
            buf = buf & ch
        ElseIf ch = ";" And Not inQuote Then
            If count <= 3 Then result(count) = buf
            count = count + 1
            buf = ""
        Else
            buf = buf & ch
        End If
    Next i
    
    If count <= 3 Then result(count) = buf
    
    ' Resize to actual count
    Dim final() As String
    ReDim final(0 To count)
    For i = 0 To count
        final(i) = result(i)
    Next i
    
    SplitFormatSections = final
End Function

'------------------------------------------------------------------------------
' IsPureLiteralSection - Returns True if a section is purely a literal/text
' placeholder with no adjustable digit placeholders (e.g. "@", """--""", "")
'------------------------------------------------------------------------------
Private Function IsPureLiteralSection(ByVal sec As String) As Boolean
    If Len(Trim$(sec)) = 0 Then IsPureLiteralSection = True: Exit Function
    If sec = "@" Then IsPureLiteralSection = True: Exit Function
    ' If it has no 0 or # outside of quotes, it's a literal
    Dim inQ As Boolean, i As Long, ch As String
    For i = 1 To Len(sec)
        ch = Mid$(sec, i, 1)
        If ch = """" Then
            inQ = Not inQ
        ElseIf Not inQ Then
            If ch = "0" Or ch = "#" Then
                IsPureLiteralSection = False
                Exit Function
            End If
        End If
    Next i
    IsPureLiteralSection = True
End Function

'------------------------------------------------------------------------------
' AdjustSectionDecimalsOne - Adjust decimals in a SINGLE section only
'------------------------------------------------------------------------------
Private Function AdjustSectionDecimalsOne(ByVal sec As String, ByVal delta As Long) As String
    Dim lastDigit As Long, dotPos As Long, p As Long, lastDec As Long
    lastDigit = LastDigitIndex(sec)
    If lastDigit = 0 Then
        AdjustSectionDecimalsOne = sec
        Exit Function
    End If
    dotPos = InStrRev(sec, ".", lastDigit)
    If delta > 0 Then
        If dotPos > 0 Then
            AdjustSectionDecimalsOne = Left$(sec, lastDigit) & "0" & Mid$(sec, lastDigit + 1)
        Else
            AdjustSectionDecimalsOne = Left$(sec, lastDigit) & ".0" & Mid$(sec, lastDigit + 1)
        End If
    Else
        If dotPos = 0 Then
            AdjustSectionDecimalsOne = sec
        Else
            p = dotPos + 1
            Do While p <= Len(sec) And (Mid$(sec, p, 1) = "0" Or Mid$(sec, p, 1) = "#")
                lastDec = p
                p = p + 1
            Loop
            If lastDec = 0 Then
                AdjustSectionDecimalsOne = sec
            ElseIf lastDec = dotPos + 1 Then
                AdjustSectionDecimalsOne = Left$(sec, dotPos - 1) & Mid$(sec, dotPos + 2)
            Else
                AdjustSectionDecimalsOne = Left$(sec, lastDec - 1) & Mid$(sec, lastDec + 1)
            End If
        End If
    End If
End Function

Private Function LastDigitIndex(ByVal sec As String) As Long
    Dim p0 As Long, pH As Long
    p0 = InStrRev(sec, "0")
    pH = InStrRev(sec, "#")
    If p0 >= pH Then LastDigitIndex = p0 Else LastDigitIndex = pH
End Function





