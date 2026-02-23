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
'   F01 – CycleNumberFormat      Ctrl+Shift+1
'   F02 – CycleDateFormat        Ctrl+Shift+3
'   F03 – CyclePercentFormat     Ctrl+Shift+5
'   F04 – CycleCurrencyFormat    Ctrl+Shift+4
'   F05 – CycleOtherNumbers      Ctrl+Shift+8
'   F06 – IncreaseDecimal        Ctrl+Shift+.
'   F07 – DecreaseDecimal        Ctrl+Shift+,
'   F08 – ScaleUp                Alt+Shift+<
'   F09 – ScaleDown              Alt+Shift+>
'   F10 – ToggleSign             Ctrl+Alt+Shift+\
'   F11 – DivideByHundred        Ctrl+Alt+2
'   F12 – MultiplyByHundred      Ctrl+Alt+Shift+2
'==============================================================================

'------------------------------------------------------------------------------
' F01 CycleNumberFormat  (Ctrl+Shift+1)
'------------------------------------------------------------------------------
Sub CycleNumberFormat()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevNumAddress Then FormatNumIndex = 0
    PrevNumAddress = Selection.Address(False, False)
    Dim fmts As Variant, ni As Long
    fmts = Array( _
      "#,##0_);(#,##0);""--"";@", _
      "#,##0,_);(#,##0,);""--"";@", _
      "#,##0,""K""_);(#,##0,""K"");""--"";@", _
      "#,##0.0,,_);(#,##0.0,,);""--"";@", _
      "#,##0.0,,""M""_);(#,##0.0,,""M"");""--"";@" _
    )
    ni = FormatNumIndex Mod (UBound(fmts) + 1)
    Selection.NumberFormat = fmts(ni)
    FormatNumIndex = FormatNumIndex + 1
    LogAction "NumFmt" & (ni + 1), PrevNumAddress
    RegisterUndo "Number Format"
End Sub

'------------------------------------------------------------------------------
' F02 CycleDateFormat  (Ctrl+Shift+3)
'------------------------------------------------------------------------------
Sub CycleDateFormat()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevDateAddress Then FormatDateIndex = 0
    PrevDateAddress = Selection.Address(False, False)
    Dim dates As Variant, ni As Long
    dates = Array("m/d/yyyy", "m/d/yy", "mmm-yy", "d-mmm-yy;d-mmm-yy;-")
    ni = FormatDateIndex Mod (UBound(dates) + 1)
    Selection.NumberFormat = dates(ni)
    FormatDateIndex = FormatDateIndex + 1
    LogAction "DateFmt" & (ni + 1), PrevDateAddress
    RegisterUndo "Date Format"
End Sub

'------------------------------------------------------------------------------
' F03 CyclePercentFormat  (Ctrl+Shift+5)
'------------------------------------------------------------------------------
Sub CyclePercentFormat()
    On Error GoTo CleanFail
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevPctAddress Then FormatPctIndex = 0
    PrevPctAddress = Selection.Address(False, False)
    Dim fmts(1 To 5) As String
    fmts(1) = "0.0%;(0.0%);""—"";@"
    fmts(2) = "0%;(0%);""—"";@"
    fmts(3) = "+0.0%;-0.0%;""—"";@"
    fmts(4) = "[<=-0.0005](0.0%);[>=0.0005]0.0%;"""";@"
    fmts(5) = "0.0%;(0.0%);"""";@"
    Dim ni As Long
    ni = (FormatPctIndex Mod UBound(fmts)) + 1
    Selection.NumberFormat = fmts(ni)
    FormatPctIndex = FormatPctIndex + 1
    LogAction "PctFmt:" & CStr(ni), PrevPctAddress
    RegisterUndo "Percent Format"
CleanExit:
    Exit Sub
CleanFail:
    Resume CleanExit
End Sub

'------------------------------------------------------------------------------
' F04 CycleCurrencyFormat  (Ctrl+Shift+4)
'------------------------------------------------------------------------------
Sub CycleCurrencyFormat()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevCurAddress Then FormatCurIndex = 0
    PrevCurAddress = Selection.Address(False, False)
    Dim fmts As Variant, ni As Long
    fmts = Array( _
      "$#,##0_);($#,##0);""--"";@", _
      "$#,##0,_);($#,##0,);""--"";@", _
      "$#,##0,""K""_);($#,##0,""K"");""--"";@", _
      "$#,##0.0,,_);($#,##0.0,,);""--"";@", _
      "$#,##0.0,,""M""_);($#,##0.0,,""M"");""--"";@" _
    )
    ni = FormatCurIndex Mod (UBound(fmts) + 1)
    Selection.NumberFormat = fmts(ni)
    FormatCurIndex = FormatCurIndex + 1
    LogAction "CurFmt" & (ni + 1), PrevCurAddress
    RegisterUndo "Currency Format"
End Sub

'------------------------------------------------------------------------------
' F05 CycleOtherNumbers  (Ctrl+Shift+8)
'------------------------------------------------------------------------------
Sub CycleOtherNumbers()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevOtherAddress Then FormatOtherIndex = 0
    PrevOtherAddress = Selection.Address(False, False)
    Dim fmts As Variant, ni As Long
    fmts = Array("0\A", "0\B", "0\F", """Q""#", "0\P", "0\E", "0.0""x""")
    ni = FormatOtherIndex Mod (UBound(fmts) + 1)
    Selection.NumberFormat = fmts(ni)
    FormatOtherIndex = FormatOtherIndex + 1
    LogAction "OtherFmt" & (ni + 1), PrevOtherAddress
    RegisterUndo "Other Numbers Format"
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

    For Each c In rng.Cells
        fmt = CStr(c.NumberFormat)
        If Not dict.Exists(fmt) Then dict(fmt) = AdjustFormatDecimals(fmt, delta)
        newFmt = dict(fmt)
        If c.NumberFormat <> newFmt Then c.NumberFormat = newFmt
    Next c
End Sub

'------------------------------------------------------------------------------
' F08 ScaleUp  (Alt+Shift+<)
'     Divides values by 1000 (converts to thousands)
'------------------------------------------------------------------------------
Sub ScaleUp()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    Application.ScreenUpdating = False
    For Each c In Selection.Cells
        If c.HasFormula Then
            c.Formula = "=(" & Mid(c.Formula, 2) & ")/1000"
        ElseIf IsNumeric(c.Value) Then
            c.Value = c.Value / 1000
        End If
    Next c
    Application.ScreenUpdating = True
    LogAction "ScaleUp", Selection.Address(False, False)
    RegisterUndo "Scale Up"
End Sub

'------------------------------------------------------------------------------
' F09 ScaleDown  (Alt+Shift+>)
'     Multiplies values by 1000 (converts from thousands)
'------------------------------------------------------------------------------
Sub ScaleDown()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    Application.ScreenUpdating = False
    For Each c In Selection.Cells
        If c.HasFormula Then
            c.Formula = "=(" & Mid(c.Formula, 2) & ")*1000"
        ElseIf IsNumeric(c.Value) Then
            c.Value = c.Value * 1000
        End If
    Next c
    Application.ScreenUpdating = True
    LogAction "ScaleDown", Selection.Address(False, False)
    RegisterUndo "Scale Down"
End Sub

'------------------------------------------------------------------------------
' F10 ToggleSign  (Ctrl+Alt+Shift+\)
'     Reverses the sign of numeric values in selection (positive <-> negative)
'------------------------------------------------------------------------------
Sub ToggleSign()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range, f As String
    For Each c In Selection.Cells
        If c.HasFormula Then
            f = c.Formula
            If Left(f, 3) = "=-(" And Right(f, 1) = ")" Then
                c.Formula = "=" & Mid(f, 4, Len(f) - 4)  ' unwrap
            Else
                c.Formula = "=-(" & Mid(f, 2) & ")"       ' wrap
            End If
        ElseIf IsNumeric(c.Value) Then
            c.Value = -c.Value
        End If
    Next c
    LogAction "ToggleSign", Selection.Address(False, False)
    RegisterUndo "Toggle Sign"
End Sub


'------------------------------------------------------------------------------
' F11 DivideByHundred  (Ctrl+Alt+2)
'     Divides values by 100 (e.g., convert 5 to 0.05 for 5%)
'------------------------------------------------------------------------------
Sub DivideByHundred()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection.Cells
        If c.HasFormula Then
            c.Formula = "=(" & Mid(c.Formula, 2) & ")/100"
        ElseIf IsNumeric(c.Value) Then
            c.Value = c.Value / 100
        End If
    Next c
    LogAction "DivideByHundred", Selection.Address(False, False)
    RegisterUndo "Divide by 100"
End Sub


'------------------------------------------------------------------------------
' F12 MultiplyByHundred  (Ctrl+Alt+Shift+2)
'     Multiplies values by 100 (e.g., convert 0.05 to 5)
'------------------------------------------------------------------------------
Sub MultiplyByHundred()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection.Cells
        If c.HasFormula Then
            c.Formula = "=(" & Mid(c.Formula, 2) & ")*100"
        ElseIf IsNumeric(c.Value) Then
            c.Value = c.Value * 100
        End If
    Next c
    LogAction "MultiplyByHundred", Selection.Address(False, False)
    RegisterUndo "Multiply by 100"
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



