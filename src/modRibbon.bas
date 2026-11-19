Attribute VB_Name = "modRibbon"

Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Module: modRibbon (Rxx)
' Purpose:
'   Callbacks for the ILLXL ribbon tab (src\ribbon.xml).
'   - Buttons carry the macro to run in their tag (Rib_Action) and show the
'     ACTIVE preset's shortcut in their tooltip (Rib_Screentip).
'   - Galleries list the active preset's steps - formats, styles, colors,
'     fonts, tags, presets - and apply the one you click (Gal_*).
'   - Style and color galleries get small swatch pictures, drawn on the fly
'     as tiny .bmp files in %TEMP% (VBA can load those with LoadPicture).
'
' The Gallery* functions do the real work and take plain ids, so the tests
' can call them without a ribbon.
'==============================================================================

Private Rib As IRibbonUI                    ' set by Rib_OnLoad; lost after a VBA reset (harmless)

Private Const SW_W As Long = 40             ' swatch size - matches itemWidth/itemHeight in ribbon.xml
Private Const SW_H As Long = 20

'------------------------------------------------------------------------------
' R01 - Ribbon load / refresh
'------------------------------------------------------------------------------
Public Sub Rib_OnLoad(ribbon As IRibbonUI)
    Set Rib = ribbon
End Sub

' Re-read every label and tooltip (after a preset switch)
Public Sub RefreshRibbon()
    On Error Resume Next
    If Not Rib Is Nothing Then Rib.Invalidate
End Sub

Public Function RibbonLoaded() As Boolean            ' test hook
    RibbonLoaded = Not Rib Is Nothing
End Function

Public Sub ShowRibbonTab()                           ' test hook (screenshots)
    On Error Resume Next
    Rib.ActivateTab "tabILLXL"
End Sub

'------------------------------------------------------------------------------
' R02 - Buttons
'------------------------------------------------------------------------------
Public Sub Rib_Action(control As IRibbonControl)
    RunCommand control.Tag
End Sub

Public Sub Rib_Screentip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ScreentipFor(control.Tag)
End Sub

' "Cycle Number Format  (Ctrl+Shift+1)" - the key comes from the active preset
Public Function ScreentipFor(ByVal macro As String) As String
    Dim keys As Object, label As String
    label = CommandInfo(macro)(1)
    Set keys = KeyAssignments()
    If keys.Exists(macro) Then
        If Len(keys(macro)) > 0 Then label = label & "   (" & keys(macro) & ")"
    End If
    ScreentipFor = label
End Function

Private Sub RunCommand(ByVal macro As String)
    On Error GoTo Failed
    Application.Run "'" & ThisWorkbook.Name & "'!" & macro
    Exit Sub
Failed:
    MsgBox "ILLXL couldn't run '" & macro & "':" & vbCrLf & Err.Description, vbExclamation, "ILLXL"
End Sub

'------------------------------------------------------------------------------
' R03 - Gallery callbacks (thin wrappers - see Gallery* below)
'------------------------------------------------------------------------------
Public Sub Gal_Count(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GalleryCount(control.Id)
End Sub

Public Sub Gal_Label(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = GalleryLabel(control.Id, index)
End Sub

Public Sub Gal_Screentip(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = GalleryTip(control.Id, index)
End Sub

Public Sub Gal_Image(control As IRibbonControl, index As Integer, ByRef returnedVal)
    On Error Resume Next
    Set returnedVal = LoadPicture(GallerySwatchFile(control.Id, index))
End Sub

Public Sub Gal_Action(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    GalleryPick control.Id, selectedIndex
End Sub

'==============================================================================
' GALLERY CONTENT
'==============================================================================

' Raw items for a gallery (0-based array)
Private Function GalleryItems(ByVal id As String) As Variant
    Dim i As Long, out() As String
    EnsureSettings
    Select Case id
        Case "galPreset":      GalleryItems = PresetNames()
        Case "galNum":         GalleryItems = cfgNumFormats
        Case "galCur":         GalleryItems = cfgCurFormats
        Case "galPct":         GalleryItems = cfgPctFormats
        Case "galDate":        GalleryItems = cfgDateFormats
        Case "galOther":       GalleryItems = cfgOtherFormats
        Case "galFont":        GalleryItems = cfgFonts
        Case "galTagValue":    GalleryItems = cfgTagsValue
        Case "galTagDuration": GalleryItems = cfgTagsDuration
        Case "galTagRate":     GalleryItems = cfgTagsRate
        Case "galFill"
            ReDim out(0 To UBound(cfgFills)): For i = 0 To UBound(out): out(i) = cfgFills(i).Name: Next i
            GalleryItems = out
        Case "galFontColor"
            ReDim out(0 To UBound(cfgFontColors)): For i = 0 To UBound(out): out(i) = cfgFontColors(i).Name: Next i
            GalleryItems = out
        Case "galTab"
            ReDim out(0 To UBound(cfgTabColors)): For i = 0 To UBound(out): out(i) = cfgTabColors(i).Name: Next i
            GalleryItems = out
        Case "galInput"
            ReDim out(0 To UBound(cfgInputStyles)): For i = 0 To UBound(out): out(i) = cfgInputStyles(i).Name: Next i
            GalleryItems = out
        Case "galHeader"
            ReDim out(0 To UBound(cfgHeaderStyles)): For i = 0 To UBound(out): out(i) = cfgHeaderStyles(i).Name: Next i
            GalleryItems = out
        Case Else:             GalleryItems = Array()
    End Select
End Function

Public Function GalleryCount(ByVal id As String) As Long
    GalleryCount = UBound(GalleryItems(id)) + 1
End Function

' What the item shows: formats as a formatted sample, everything else by name
Public Function GalleryLabel(ByVal id As String, ByVal i As Long) As String
    Dim items As Variant
    items = GalleryItems(id)
    If i < 0 Or i > UBound(items) Then GalleryLabel = "": Exit Function
    Select Case id
        Case "galPreset"
            GalleryLabel = items(i) & IIf(StrComp(items(i), ActivePreset, vbTextCompare) = 0, _
                                          "   " & ChrW(&H2713) & " active", "")
        Case "galNum", "galCur": GalleryLabel = Sample(1234567.891, items(i)) & "      " & Sample(-1234.5, items(i))
        Case "galPct":           GalleryLabel = Sample(0.1234, items(i)) & "      " & Sample(-0.05, items(i))
        Case "galDate":          GalleryLabel = Sample(CDbl(Date), items(i))
        Case "galOther":         GalleryLabel = Sample(2026, items(i)) & "      " & Sample(12.5, items(i))
        Case Else:               GalleryLabel = items(i)
    End Select
End Function

' Tooltip: the exact number-format code, or the item name
Public Function GalleryTip(ByVal id As String, ByVal i As Long) As String
    Dim items As Variant
    items = GalleryItems(id)
    If i < 0 Or i > UBound(items) Then GalleryTip = "": Exit Function
    GalleryTip = items(i)
End Function

' Apply the clicked item
Public Sub GalleryPick(ByVal id As String, ByVal i As Long)
    If i < 0 Or i >= GalleryCount(id) Then Exit Sub
    Select Case id
        Case "galPreset":      ApplyPreset GalleryItems(id)(i)
        Case "galNum":         ApplyFormatAt "Num", i
        Case "galCur":         ApplyFormatAt "Cur", i
        Case "galPct":         ApplyFormatAt "Pct", i
        Case "galDate":        ApplyFormatAt "Date", i
        Case "galOther":       ApplyFormatAt "Other", i
        Case "galFont":        ApplyFontAt i
        Case "galFill":        ApplyFillAt i
        Case "galFontColor":   ApplyFontColorAt i
        Case "galTab":         ApplyTabColorAt i
        Case "galInput":       ApplyInputAt i
        Case "galHeader":      ApplyHeaderAt i
        Case "galTagValue":    ApplyTagAt "Value", i
        Case "galTagDuration": ApplyTagAt "Duration", i
        Case "galTagRate":     ApplyTagAt "Rate", i
    End Select
End Sub

Private Function Sample(ByVal v As Double, ByVal fmt As String) As String
    On Error Resume Next
    Sample = Trim$(Application.WorksheetFunction.Text(v, fmt))
    If Err.Number <> 0 Or Len(Sample) = 0 Then Sample = fmt
End Function

'==============================================================================
' SWATCHES  (small .bmp pictures for the style / color galleries)
'==============================================================================

' Path of the swatch picture for one gallery item (drawn once, then reused)
Public Function GallerySwatchFile(ByVal id As String, ByVal i As Long) As String
    Dim fill As Long, bar As Long, frame As Long, kind As String
    EnsureSettings
    frame = RGB(191, 191, 191): kind = "solid": bar = -1
    Select Case id
        Case "galFill":       fill = cfgFills(i).Color
        Case "galTab":        fill = cfgTabColors(i).Color
        Case "galFontColor":  fill = RGB(255, 255, 255): bar = cfgFontColors(i).Color: kind = "text"
        Case "galInput":      fill = cfgInputStyles(i).Fill: bar = cfgInputStyles(i).Font
                              frame = cfgInputStyles(i).Border: kind = "text"
        Case "galHeader":     fill = cfgHeaderStyles(i).Fill: bar = cfgHeaderStyles(i).Font: kind = "text"
        Case Else:            Exit Function
    End Select
    If fill = CLR_NONE Then kind = "none": fill = RGB(255, 255, 255)
    If fill = CLR_STRIPE Then kind = "stripe": fill = RGB(255, 255, 255)

    GallerySwatchFile = Environ$("TEMP") & "\ILLXL_sw_" & kind & "_" & Hex$(fill) & "_" & Hex$(bar And &HFFFFFF) & _
                        "_" & Hex$(frame) & ".bmp"
    If Len(Dir$(GallerySwatchFile)) = 0 Then WriteSwatch GallerySwatchFile, fill, bar, frame, kind
End Function

' 24-bit BMP, SW_W x SW_H:  solid fill | "text" = fill + a bar of the font color
'                           | "none" = white with a red slash | "stripe" = gray diagonals
Private Sub WriteSwatch(ByVal path As String, ByVal fill As Long, ByVal bar As Long, ByVal frame As Long, ByVal kind As String)
    Dim rowBytes As Long, size As Long, b() As Byte, x As Long, y As Long, p As Long, c As Long, f As Integer
    rowBytes = ((SW_W * 3 + 3) \ 4) * 4
    size = 54 + rowBytes * SW_H
    ReDim b(0 To size - 1)
    b(0) = 66: b(1) = 77                                    ' "BM"
    PutLong b, 2, size: PutLong b, 10, 54: PutLong b, 14, 40
    PutLong b, 18, SW_W: PutLong b, 22, SW_H
    b(26) = 1: b(28) = 24                                   ' planes, bits per pixel
    PutLong b, 34, rowBytes * SW_H
    For y = 0 To SW_H - 1                                   ' BMP rows run bottom-up
        For x = 0 To SW_W - 1
            c = SwatchPixel(x, SW_H - 1 - y, fill, bar, frame, kind)
            p = 54 + y * rowBytes + x * 3
            b(p) = (c \ &H10000) And &HFF: b(p + 1) = (c \ &H100) And &HFF: b(p + 2) = c And &HFF
        Next x
    Next y
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, , b
    Close #f
End Sub

Private Function SwatchPixel(ByVal x As Long, ByVal y As Long, ByVal fill As Long, ByVal bar As Long, _
                             ByVal frame As Long, ByVal kind As String) As Long
    If x = 0 Or y = 0 Or x = SW_W - 1 Or y = SW_H - 1 Then SwatchPixel = frame: Exit Function
    SwatchPixel = fill
    Select Case kind
        Case "text"                                         ' a "line of text" in the font color
            If y >= SW_H \ 2 - 2 And y <= SW_H \ 2 + 1 And x >= SW_W \ 5 And x <= SW_W - SW_W \ 5 Then SwatchPixel = bar
        Case "none"
            If Abs(x * (SW_H - 1) \ (SW_W - 1) - (SW_H - 1 - y)) <= 1 Then SwatchPixel = RGB(220, 38, 38)
        Case "stripe"
            If (x + y) Mod 6 < 2 Then SwatchPixel = RGB(166, 166, 166)
    End Select
End Function

Private Sub PutLong(ByRef b() As Byte, ByVal at As Long, ByVal v As Long)
    b(at) = v And &HFF: b(at + 1) = (v \ &H100) And &HFF
    b(at + 2) = (v \ &H10000) And &HFF: b(at + 3) = (v \ &H1000000) And &HFF
End Sub
