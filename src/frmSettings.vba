Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Form: frmSettings  (Ctrl+Alt+Shift+S)
' Purpose:
'   The ILLXL Settings window. Edits one preset at a time - a complete copy
'   held in memory (Ed) - and writes it back on Save. Saving the active
'   preset reloads it and rebinds every shortcut immediately.
'
'   Every control is created in code (BuildUi) so the whole window lives in
'   this text file. Events arrive through clsUiEvents and land in OnUi.
'
'   Pages (sidebar):
'     Presets         pick / activate / duplicate / rename / delete / import
'     Number Formats  the format cycles, with live validation and preview
'     Colors & Styles fills, font colors, tabs, input + header styles, Auto-Color
'     Fonts & Borders font cycle and border-weight cycle
'     Tags & Headers  unit tags, header defaults, disabled keys
'     Shortcuts       press a key combination to assign it; conflicts flagged
'
'   Control Ids:  btnX = button   lstX/cboX/txtX = inputs   nav:page
'                 bind:Section:Key = text box saved straight into the preset
'                 swa:n / hex:n = color row n   pal:n = palette color n
'==============================================================================

'==== Layout (points) ====
Private Const FW As Single = 670        ' inside width
Private Const FH As Single = 462        ' inside height
Private Const SIDE_W As Single = 150    ' sidebar
Private Const CL As Single = 168        ' content left
Private Const CW As Single = 486        ' content width
Private Const PAGE_T As Single = 62     ' page frames top
Private Const PAGE_H As Single = 350

'==== Section lists ====
Private Const FMT_SECS As String = "NumberFormats|CurrencyFormats|PercentFormats|DateFormats|OtherFormats|TimestampFormats"
Private Const FMT_NAMES As String = "Number|Currency|Percent|Date|Other (A / B / F / x)|Timestamp"
Private Const CLR_SECS As String = "Fills|FontColors|TabColors|InputStyles|HeaderStyles|AutoColor|ZeroCheck|Headers"
Private Const CLR_NAMES As String = "Fill cycle|Font color cycle|Tab colors|Input styles|Header styles|Auto-Color convention|Zero-check colors|Inserted headers"
Private Const BORDER_STEPS As String = "thin|medium|thick|hairline|double|none"

'==== Theme colors (set in InitTheme) ====
Private cBg As Long, cText As Long, cMuted As Long, cLine As Long, cAccent As Long
Private cAccentHover As Long, cSide As Long, cSideSel As Long, cSideText As Long
Private cBtnHover As Long, cError As Long, cWarn As Long, cOk As Long, cSoft As Long

'==== UI registry ====
Private Hooks   As Collection           ' keeps every clsUiEvents alive
Private Ctl     As Object               ' id -> control
Private BtnBg   As Object               ' id -> background label of a flat button
Private BtnCap  As Object               ' id -> caption label of a flat button
Private BtnKind As Object               ' id -> "P" primary, "B" button, "N" nav, "S" swatch
Private BtnOff  As Object               ' id -> True while disabled
Private Pages   As Object               ' page key -> frame
Private HoverId As String
Private CurPage As String
Private Busy    As Boolean              ' True while filling controls (ignore their events)

'==== Editing state ====
Private Ed      As Object               ' the preset being edited (full settings tree)
Private EdName  As String
Private Dirty   As Boolean
Private ClrRow  As Long                 ' color row the palette writes to
Private KeyMacros As Variant            ' macro name per row of lstKeys
Private PaletteHex As Variant

'==== Test hooks (SelfTest) ====
Public Quiet As Boolean                 ' no message boxes / input boxes
Public QuietAnswer As VbMsgBoxResult
Public QuietText As String

'==============================================================================
' OPEN / CLOSE
'==============================================================================
Private Sub UserForm_Initialize()
    InitTheme
    BuildUi
    LoadEdit ActivePreset
    ShowPage "presets"
End Sub

' InsideWidth is only reliable once the window exists, so trim the frame here
Private Sub UserForm_Activate()
    Me.Width = FW + (Me.Width - Me.InsideWidth)
    Me.Height = FH + (Me.Height - Me.InsideHeight)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If Not ConfirmLeave() Then Cancel = True
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    OnHover ""
End Sub

'==============================================================================
' EVENT ROUTER  (called by clsUiEvents)
'==============================================================================
Public Sub OnUi(ByVal id As String, ByVal evt As String)
    If Busy Then Exit Sub
    If BtnOff.Exists(id) Then Exit Sub                  ' disabled button

    If Left$(id, 4) = "nav:" Then ShowPage Mid$(id, 5): Exit Sub
    If Left$(id, 5) = "bind:" Then BoundTextChanged id: Exit Sub
    If Left$(id, 4) = "pal:" Then PaletteClick CLng(Mid$(id, 5)): Exit Sub
    If Left$(id, 4) = "swa:" Then SetColorRow CLng(Mid$(id, 5)): Exit Sub
    If Left$(id, 4) = "hex:" Then HexChanged CLng(Mid$(id, 5)): Exit Sub

    Select Case id
        '---- footer
        Case "btnSave":       SaveEdits
        Case "btnClose":      If ConfirmLeave() Then Unload Me
        '---- presets
        Case "lstPresets":    PresetPicked
        Case "btnActivate":   PresetActivate
        Case "btnDuplicate":  PresetDuplicate
        Case "btnRename":     PresetRename
        Case "btnDelete":     PresetDelete
        Case "btnImport":     PresetImport
        Case "btnExport":     PresetExport
        Case "btnReset":      PresetReset
        Case "btnFolder":     OpenPresetFolder
        '---- number formats
        Case "cboFmt":        RefreshFormats 0
        Case "lstFmt":        FormatPicked
        Case "txtFmt":        FormatTyped
        Case "btnFmtAdd":     FormatAdd
        Case "btnFmtDel":     FormatRemove
        Case "btnFmtUp":      FormatMove -1
        Case "btnFmtDown":    FormatMove 1
        '---- colors
        Case "cboClr":        RefreshColors 0
        Case "lstClr":        ColorPicked
        Case "txtClrName":    ColorNameTyped
        Case "btnClrAdd":     ColorAdd
        Case "btnClrDel":     ColorRemove
        Case "btnClrUp":      ColorMove -1
        Case "btnClrDown":    ColorMove 1
        Case "btnClrNone":    SetHexText ClrRow, "none"
        Case "btnClrStripe":  SetHexText ClrRow, "stripe"
        '---- fonts & borders
        Case "lstFonts":      FontPicked
        Case "btnFontAdd":    FontAdd
        Case "btnFontDel":    ListRemove "Fonts", "lstFonts"
        Case "btnFontUp":     ListMove "Fonts", "lstFonts", -1
        Case "btnFontDown":   ListMove "Fonts", "lstFonts", 1
        Case "btnBdrAdd":     BorderEdit "add"
        Case "btnBdrDel":     BorderEdit "del"
        Case "btnBdrUp":      BorderEdit "up"
        Case "btnBdrDown":    BorderEdit "down"
        '---- shortcuts
        Case "lstKeys":       KeyPicked
        Case "btnKeyClear":   KeyClear
        Case "btnKeyDefault": KeyDefault
    End Select
End Sub

Public Sub OnKeyDown(ByVal id As String, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If id <> "txtKeyCap" Then Exit Sub
    Dim combo As String
    combo = ComboText(KeyCode, Shift)
    KeyCode = 0                                         ' never type into the box
    If Len(combo) = 0 Then Exit Sub                     ' a modifier on its own
    KeyCaptured combo, Shift
End Sub

'---- Hover: one flat button highlighted at a time ------------------------------
Public Sub OnHover(ByVal id As String)
    If id = HoverId Then Exit Sub
    If Len(HoverId) > 0 Then PaintButton HoverId, False
    HoverId = ""
    If Len(id) > 0 Then
        If BtnBg.Exists(id) Then
            If Not BtnOff.Exists(id) Then HoverId = id: PaintButton id, True
        End If
    End If
End Sub

'==============================================================================
' BUILDING THE WINDOW
'==============================================================================
Private Sub InitTheme()
    cBg = RGB(255, 255, 255):       cText = RGB(17, 24, 39):      cMuted = RGB(107, 114, 128)
    cLine = RGB(209, 213, 219):     cSoft = RGB(243, 244, 246)
    cAccent = RGB(37, 99, 235):     cAccentHover = RGB(29, 78, 216)
    cSide = RGB(17, 24, 39):        cSideSel = RGB(31, 41, 55):   cSideText = RGB(209, 213, 219)
    cBtnHover = RGB(239, 246, 255): cError = RGB(220, 38, 38):    cWarn = RGB(217, 119, 6)
    cOk = RGB(22, 163, 74)
    PaletteHex = Split( _
        "#FFFFFF #000000 #E7E6E6 #44546A #4472C4 #ED7D31 #A5A5A5 #FFC000 #5B9BD5 #70AD47 " & _
        "#C00000 #FF0000 #FFC000 #FFFF00 #92D050 #00B050 #00B0F0 #0070C0 #002060 #7030A0 " & _
        "#FFFFCC #FFF2CC #FCE4D6 #FFCC99 #FFC7CE #C6EFCE #DDEBF7 #D9E1F2 #EDEDED #D9D9D9 " & _
        "#0000FF #0070C0 #008000 #00B0B0 #0085B2 #7F7F7F #1F4E79 #00008B #112547 #0D2535 " & _
        "#1F3864 #44546A #F8F4EC #F2F2F2 #BFBFBF #808080 #595959 #FFFF99 #A6A6A6 #FF00FF", " ")
End Sub

Private Sub BuildUi()
    Set Hooks = New Collection
    Set Ctl = NewDict(): Set BtnBg = NewDict(): Set BtnCap = NewDict()
    Set BtnKind = NewDict(): Set BtnOff = NewDict(): Set Pages = NewDict()

    Me.Caption = "ILLXL Settings"
    Me.BackColor = cBg
    Me.Width = FW + 12: Me.Height = FH + 29         ' close enough until shown - see FitWindow
    Me.Font.Name = "Segoe UI": Me.Font.Size = 9

    '---- Sidebar -------------------------------------------------------------
    NewLabel Me, "", 0, 0, SIDE_W, FH, , , , cSide          ' sidebar background
    NewLabel Me, "ILLXL", 18, 16, SIDE_W - 20, 26, 18, True, RGB(255, 255, 255)
    NewLabel Me, "Settings   v" & ILLXL_VERSION, 19, 42, SIDE_W - 20, 14, 9, False, cSideText
    Dim navs As Variant, i As Long
    navs = Array("presets|Presets", "formats|Number Formats", "colors|Colors & Styles", _
                 "fonts|Fonts & Borders", "tags|Tags & Headers", "keys|Shortcuts")
    For i = 0 To UBound(navs)
        NewButton Me, "nav:" & Split(navs(i), "|")(0), Split(navs(i), "|")(1), 0, 76 + i * 30, SIDE_W, 28, "N"
    Next i
    NewLabel Me, "EDITING", 19, FH - 58, SIDE_W - 20, 12, 7.5, True, cMuted
    Set Ctl("lblEditing") = NewLabel(Me, "", 19, FH - 44, SIDE_W - 24, 30, 9.5, True, RGB(255, 255, 255))

    '---- Header + footer -------------------------------------------------------
    Set Ctl("lblTitle") = NewLabel(Me, "", CL, 14, CW, 24, 15, True)
    Set Ctl("lblSub") = NewLabel(Me, "", CL, 38, CW, 16, 8.5, False, cMuted)
    NewLabel Me, "", CL, FH - 48, CW, 1, , , , cLine        ' footer divider
    Set Ctl("lblStatus") = NewLabel(Me, "", CL, FH - 32, CW - 190, 16, 8.5, False, cMuted)
    NewButton Me, "btnSave", "Save", FW - 186, FH - 37, 84, 24, "P"
    NewButton Me, "btnClose", "Close", FW - 96, FH - 37, 80, 24, "B"

    '---- Pages ---------------------------------------------------------------
    BuildPresetsPage NewPage("presets")
    BuildFormatsPage NewPage("formats")
    BuildColorsPage NewPage("colors")
    BuildFontsPage NewPage("fonts")
    BuildTagsPage NewPage("tags")
    BuildKeysPage NewPage("keys")
End Sub

Private Sub BuildPresetsPage(ByVal pg As Object)
    Set Ctl("lstPresets") = NewList(pg, "lstPresets", 0, 0, 300, 112, 2, "200;90")
    NewButton pg, "btnActivate", "Make Active", 316, 0, 170, 24, "P"
    NewButton pg, "btnDuplicate", "Duplicate...", 316, 30, 170, 24
    NewButton pg, "btnRename", "Rename...", 316, 60, 170, 24
    NewButton pg, "btnDelete", "Delete", 316, 90, 170, 24
    NewButton pg, "btnImport", "Import...", 0, 124, 115, 24
    NewButton pg, "btnExport", "Export...", 123, 124, 115, 24
    NewButton pg, "btnReset", "Reset to Shipped", 246, 124, 115, 24
    NewButton pg, "btnFolder", "Open Folder", 371, 124, 115, 24
    NewLabel pg, "Description", 0, 164, CW, 14, 8.5, True, cMuted
    Set Ctl("bind:Info:Description") = NewText(pg, "bind:Info:Description", 0, 180, CW, 40, True)
    NewLabel pg, "Presets live in %APPDATA%\ILLXL\presets as plain .ini files - share one by " & _
                 "exporting it. Save writes the preset you are editing; if it is the active one, " & _
                 "your shortcuts update immediately.", 0, 236, CW, 44, 8.5, False, cMuted
End Sub

Private Sub BuildFormatsPage(ByVal pg As Object)
    Set Ctl("cboFmt") = NewCombo(pg, "cboFmt", 0, 0, 230, Split(FMT_NAMES, "|"))
    Set Ctl("lblFmtKey") = NewLabel(pg, "", 240, 3, 246, 16, 8.5, False, cMuted)
    Set Ctl("lstFmt") = NewList(pg, "lstFmt", 0, 28, CW, 140)
    ListButtons pg, "btnFmt", 0, 176, 110
    NewLabel pg, "Format code", 0, 212, CW, 14, 8.5, True, cMuted
    Set Ctl("txtFmt") = NewText(pg, "txtFmt", 0, 228, CW, 20)
    Set Ctl("lblFmtErr") = NewLabel(pg, "Excel can't use this format - the last valid version is kept until you fix it.", _
                                    0, 252, CW, 14, 8.5, False, cError)
    NewLabel pg, "Preview", 0, 274, CW, 14, 8.5, True, cMuted
    Set Ctl("lblFmtPrev") = NewLabel(pg, "", 0, 290, CW, 50, 10, False, cText, cSoft)
    Ctl("lblFmtPrev").Font.Name = "Consolas"
End Sub

Private Sub BuildColorsPage(ByVal pg As Object)
    Dim r As Long, k As Long
    Set Ctl("cboClr") = NewCombo(pg, "cboClr", 0, 0, 236, Split(CLR_NAMES, "|"))
    Set Ctl("lstClr") = NewList(pg, "lstClr", 0, 28, 236, 226, 2, "130;90")
    ListButtons pg, "btnClr", 0, 262, 56
    Set Ctl("lblClrPrev") = NewLabel(pg, "", 0, 294, 236, 46, 11, False)
    Ctl("lblClrPrev").TextAlign = fmTextAlignCenter

    NewLabel pg, "Name", 252, 28, 234, 14, 8.5, True, cMuted
    Set Ctl("txtClrName") = NewText(pg, "txtClrName", 252, 44, 234, 20)
    For r = 0 To 2
        Set Ctl("lblRow" & r) = NewLabel(pg, "", 252, 78 + r * 30, 50, 16, 9, False)
        NewButton pg, "swa:" & r, "", 306, 74 + r * 30, 30, 24, "S"
        Set Ctl("hex:" & r) = NewText(pg, "hex:" & r, 344, 76 + r * 30, 90, 20)
    Next r
    NewButton pg, "btnClrNone", "No color", 252, 166, 78, 22
    NewButton pg, "btnClrStripe", "Stripe", 336, 166, 78, 22
    NewLabel pg, "Quick colors - click a swatch above, then a color", 252, 196, 234, 14, 8.5, True, cMuted
    For k = 0 To UBound(PaletteHex)
        NewButton pg, "pal:" & k, "", 252 + (k Mod 10) * 23.4, 214 + (k \ 10) * 20, 20, 17, "S"
        PaintSwatch "pal:" & k, PaletteHex(k)
    Next k
End Sub

Private Sub BuildFontsPage(ByVal pg As Object)
    NewLabel pg, "Font cycle", 0, 0, 236, 14, 8.5, True, cMuted
    Set Ctl("lstFonts") = NewList(pg, "lstFonts", 0, 16, 236, 160)
    Set Ctl("cboFontAdd") = NewCombo(pg, "cboFontAdd", 0, 184, 176, InstalledFonts(), True)
    NewButton pg, "btnFontAdd", "Add", 182, 183, 54, 22
    ListButtons pg, "btnFont", 0, 214, 56, False
    Set Ctl("lblFontPrev") = NewLabel(pg, "AaBbCc  1,234.56", 0, 246, 236, 40, 14, False, cText, cSoft)

    NewLabel pg, "Border cycle (Ctrl+Alt+Shift+arrows)", 252, 0, 234, 14, 8.5, True, cMuted
    Set Ctl("lstBorders") = NewList(pg, "lstBorders", 252, 16, 234, 160)
    Set Ctl("cboBdrAdd") = NewCombo(pg, "cboBdrAdd", 252, 184, 174, Split(BORDER_STEPS, "|"))
    NewButton pg, "btnBdrAdd", "Add", 432, 183, 54, 22
    ListButtons pg, "btnBdr", 252, 214, 56, False
    NewLabel pg, "Each press moves the selected edge to the next step, " & _
                 "starting again at step 1 when the selection changes.", 252, 246, 234, 40, 8.5, False, cMuted
End Sub

Private Sub BuildTagsPage(ByVal pg As Object)
    Dim f As Variant, i As Long, y As Single
    f = Array("UnitTags:Value|Value tags  (Ctrl+Alt+Shift+T)", _
              "UnitTags:Duration|Duration tags  (Ctrl+Alt+Shift+O)", _
              "UnitTags:Rate|Rate tags  (Ctrl+Alt+Shift+P)", _
              "Headers:PromptDefault|Insert Headers - default answer", _
              "Headers:VarianceLabels|Variance headers - labels", _
              "General:DisabledKeys|Keys switched off (e.g. F1, ScrollLock, NumLock, Insert)")
    For i = 0 To UBound(f)
        y = i * 46
        NewLabel pg, Split(f(i), "|")(1) & "   - comma-separated", 0, y, CW, 14, 8.5, True, cMuted
        Set Ctl("bind:" & Split(f(i), "|")(0)) = NewText(pg, "bind:" & Split(f(i), "|")(0), 0, y + 16, CW, 20)
    Next i
End Sub

Private Sub BuildKeysPage(ByVal pg As Object)
    Set Ctl("lstKeys") = NewList(pg, "lstKeys", 0, 0, CW, 222, 3, "70;250;140")
    Set Ctl("lblKeyCmd") = NewLabel(pg, "Pick a command above", 0, 230, CW, 16, 10, True)
    NewLabel pg, "New shortcut - click the box, then press the keys (Ctrl and/or Alt + a key)", _
             0, 250, CW, 14, 8.5, False, cMuted
    Set Ctl("txtKeyCap") = NewText(pg, "txtKeyCap", 0, 268, 200, 22)
    Ctl("txtKeyCap").Font.Size = 10: Ctl("txtKeyCap").Font.Bold = True
    Ctl("txtKeyCap").TabKeyBehavior = True: Ctl("txtKeyCap").EnterKeyBehavior = True
    NewButton pg, "btnKeyClear", "No shortcut", 210, 267, 110, 24
    NewButton pg, "btnKeyDefault", "Shipped default", 328, 267, 110, 24
    Set Ctl("lblKeyMsg") = NewLabel(pg, "", 0, 298, CW, 44, 8.5, False, cMuted)
End Sub

'---- Page switching -------------------------------------------------------------
Private Function NewPage(ByVal key As String) As Object
    Dim f As MSForms.Frame
    Set f = Me.Controls.Add("Forms.Frame.1", "page_" & key)
    With f
        .Left = CL: .Top = PAGE_T: .Width = CW + 2: .Height = PAGE_H
        .BorderStyle = fmBorderStyleNone: .SpecialEffect = fmSpecialEffectFlat
        .Caption = "": .BackColor = cBg: .Visible = False
        .Font.Name = "Segoe UI": .Font.Size = 9
    End With
    Hook f, "page_" & key
    Set Pages(key) = f
    Set NewPage = f
End Function

Private Sub ShowPage(ByVal key As String)
    Dim k As Variant, titles As Variant, subs As Variant, idx As Long
    CurPage = key
    For Each k In Pages.Keys
        Pages(k).Visible = (k = key)
        PaintButton "nav:" & k, False
    Next k
    titles = Array("presets|Presets", "formats|Number Formats", "colors|Colors & Styles", _
                   "fonts|Fonts & Borders", "tags|Tags & Headers", "keys|Shortcuts")
    subs = Array("Up to 3. The active preset drives every shortcut. Click one to edit it.", _
                 "Each press of the shortcut moves to the next format in the list.", _
                 "Cycle colors, input and header styles, and the Auto-Color convention.", _
                 "Which fonts and border weights the cycles step through.", _
                 "Unit tag cycles, header defaults and keys ILLXL switches off.", _
                 "Click a command, then press the new key combination.")
    For idx = 0 To UBound(titles)
        If Split(titles(idx), "|")(0) = key Then
            Ctl("lblTitle").Caption = Split(titles(idx), "|")(1)
            Ctl("lblSub").Caption = subs(idx)
        End If
    Next idx
End Sub

'==============================================================================
' CONTROL FACTORY  (flat, modern look on standard MSForms controls)
'==============================================================================
Private Function NewLabel(ByVal parent As Object, ByVal caption As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single, Optional ByVal size As Single = 9, _
                          Optional ByVal bold As Boolean = False, Optional ByVal fore As Long = -1, _
                          Optional ByVal back As Long = -1) As MSForms.Label
    Dim c As MSForms.Label
    Set c = parent.Controls.Add("Forms.Label.1")
    With c
        .Caption = caption: .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Name = "Segoe UI": .Font.Size = size: .Font.Bold = bold
        .ForeColor = IIf(fore = -1, cText, fore)
        .WordWrap = True
        If back = -1 Then
            .BackStyle = fmBackStyleTransparent
        Else
            .BackStyle = fmBackStyleOpaque: .BackColor = back
        End If
    End With
    Set NewLabel = c
End Function

' Flat button = background label + vertically centred caption label (MSForms
' labels can't centre text vertically). Both forward clicks under one id.
Private Sub NewButton(ByVal parent As Object, ByVal id As String, ByVal caption As String, ByVal l As Single, _
                      ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal kind As String = "B")
    Dim bg As MSForms.Label, cap As MSForms.Label
    Set bg = NewLabel(parent, "", l, t, w, h, , , , cBg)
    bg.BorderStyle = fmBorderStyleSingle
    Set cap = NewLabel(parent, caption, l + IIf(kind = "N", 20, 2), t + (h - 14) / 2, w - IIf(kind = "N", 22, 4), 14, _
                       IIf(kind = "N", 9.5, 9), kind = "P")
    If kind <> "N" Then cap.TextAlign = fmTextAlignCenter
    cap.WordWrap = False
    Set BtnBg(id) = bg: Set BtnCap(id) = cap: BtnKind(id) = kind
    Hook bg, id: Hook cap, id
    PaintButton id, False
End Sub

Private Sub PaintButton(ByVal id As String, ByVal hover As Boolean)
    Dim bg As MSForms.Label, cap As MSForms.Label, off As Boolean, selected As Boolean
    If Not BtnBg.Exists(id) Then Exit Sub
    Set bg = BtnBg(id): Set cap = BtnCap(id): off = BtnOff.Exists(id)
    Select Case BtnKind(id)
        Case "P"
            bg.BackColor = IIf(off, cLine, IIf(hover, cAccentHover, cAccent)): bg.BorderColor = bg.BackColor
            cap.ForeColor = RGB(255, 255, 255)
        Case "N"                                        ' sidebar item: current page stays lit
            selected = (CurPage = Mid$(id, 5))
            bg.BorderStyle = fmBorderStyleNone
            bg.BackColor = IIf(selected, cSideSel, IIf(hover, RGB(26, 34, 48), cSide))
            cap.ForeColor = IIf(selected, RGB(255, 255, 255), cSideText)
            cap.Font.Bold = selected
        Case "S"
            bg.BorderColor = IIf(hover, cText, cLine)
        Case Else
            bg.BackColor = IIf(hover, cBtnHover, cBg)
            bg.BorderColor = IIf(hover, cAccent, cLine)
            cap.ForeColor = IIf(off, cLine, cText)
    End Select
End Sub

Private Sub EnableButton(ByVal id As String, ByVal enabled As Boolean)
    If enabled Then
        If BtnOff.Exists(id) Then BtnOff.Remove id
    Else
        BtnOff(id) = True
    End If
    PaintButton id, False
End Sub

Private Sub PaintSwatch(ByVal id As String, ByVal hexValue As String)
    Dim bg As MSForms.Label, cap As MSForms.Label, clr As Long
    Set bg = BtnBg(id): Set cap = BtnCap(id)
    cap.Font.Size = 7: cap.ForeColor = cMuted: cap.Caption = ""
    cap.Top = bg.Top + (bg.Height - 11) / 2: cap.Height = 11  ' small text, re-centred
    Select Case LCase$(Trim$(hexValue))
        Case "none":   bg.BackColor = cBg: cap.Caption = "none"
        Case "stripe": bg.BackColor = RGB(230, 230, 230): cap.Caption = "////"
        Case Else
            If HexToColor(hexValue, clr) Then bg.BackColor = clr Else bg.BackColor = cBg: cap.Caption = "?"
    End Select
End Sub

Private Function NewList(ByVal parent As Object, ByVal id As String, ByVal l As Single, ByVal t As Single, _
                         ByVal w As Single, ByVal h As Single, Optional ByVal cols As Long = 1, _
                         Optional ByVal widths As String = "") As MSForms.ListBox
    Dim c As MSForms.ListBox
    Set c = parent.Controls.Add("Forms.ListBox.1")
    With c
        .Left = l: .Top = t: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat: .BorderStyle = fmBorderStyleSingle: .BorderColor = cLine
        .Font.Name = "Segoe UI": .Font.Size = 9: .ForeColor = cText
        .ColumnCount = cols
        If Len(widths) > 0 Then .ColumnWidths = widths
    End With
    Hook c, id
    Set NewList = c
End Function

Private Function NewText(ByVal parent As Object, ByVal id As String, ByVal l As Single, ByVal t As Single, _
                         ByVal w As Single, ByVal h As Single, Optional ByVal multi As Boolean = False) As MSForms.TextBox
    Dim c As MSForms.TextBox
    Set c = parent.Controls.Add("Forms.TextBox.1")
    With c
        .Left = l: .Top = t: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat: .BorderStyle = fmBorderStyleSingle: .BorderColor = cLine
        .Font.Name = "Segoe UI": .Font.Size = 9: .ForeColor = cText
        .MultiLine = multi: .WordWrap = multi
    End With
    Hook c, id
    Set NewText = c
End Function

Private Function NewCombo(ByVal parent As Object, ByVal id As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal items As Variant, Optional ByVal editable As Boolean = False) As MSForms.ComboBox
    Dim c As MSForms.ComboBox, i As Long
    Set c = parent.Controls.Add("Forms.ComboBox.1")
    With c
        .Left = l: .Top = t: .Width = w: .Height = 20
        .SpecialEffect = fmSpecialEffectFlat: .BorderStyle = fmBorderStyleSingle: .BorderColor = cLine
        .Font.Name = "Segoe UI": .Font.Size = 9: .ForeColor = cText
        .Style = IIf(editable, fmStyleDropDownCombo, fmStyleDropDownList)
        .ListRows = 14
        For i = LBound(items) To UBound(items): .AddItem items(i): Next i
        If Not editable And .ListCount > 0 Then Busy = True: .ListIndex = 0: Busy = False
    End With
    Hook c, id
    Set NewCombo = c
End Function

' Add / Remove / Up / Down row under a list (ids: prefix & Add, Del, Up, Down)
Private Sub ListButtons(ByVal parent As Object, ByVal prefix As String, ByVal l As Single, ByVal t As Single, _
                        ByVal w As Single, Optional ByVal withAdd As Boolean = True)
    Dim caps As Variant, ids As Variant, i As Long, n As Long
    caps = Array("Add", "Remove", "Move Up", "Move Down")
    ids = Array("Add", "Del", "Up", "Down")
    For i = IIf(withAdd, 0, 1) To 3
        NewButton parent, prefix & ids(i), caps(i), l + n * (w + 4), t, w, 22
        n = n + 1
    Next i
End Sub

Private Sub Hook(ByVal c As Object, ByVal id As String)
    Dim h As New clsUiEvents
    Set h.Owner = Me: h.Id = id
    Select Case TypeName(c)
        Case "Label":   Set h.Lbl = c
        Case "ListBox": Set h.Lst = c
        Case "ComboBox": Set h.Cbo = c
        Case "TextBox": Set h.Txt = c
        Case "Frame":   Set h.Fra = c
    End Select
    Hooks.Add h
End Sub

'==============================================================================
' LOADING / SAVING THE PRESET BEING EDITED
'==============================================================================
Private Sub LoadEdit(ByVal presetName As String)
    Set Ed = EditLoad(presetName)
    EdName = presetName
    Dirty = False
    RefreshAll
    SetStatus "", cMuted
End Sub

Private Sub RefreshAll()
    RefreshPresets
    RefreshFormats 0
    RefreshColors 0
    RefreshFonts
    RefreshBorders 0
    RefreshBoundTexts
    RefreshKeys 0
    Ctl("lblEditing").Caption = EdName & IIf(IsActive(EdName), vbLf & "(active)", "")
End Sub

Private Sub SaveEdits()
    EditSave EdName, Ed
    Dirty = False
    RefreshPresets
    If Len(SettingsWarnings()) > 0 And IsActive(EdName) Then
        SetStatus "Saved - some settings were skipped (see the message).", cWarn
    ElseIf IsActive(EdName) Then
        SetStatus "Saved - shortcuts updated.", cOk
    Else
        SetStatus "Saved. Click Make Active to use this preset.", cOk
    End If
End Sub

' True when it is OK to move away from the current edits
Private Function ConfirmLeave() As Boolean
    If Not Dirty Then ConfirmLeave = True: Exit Function
    Select Case Ask("Save changes to '" & EdName & "' first?", vbYesNoCancel + vbQuestion)
        Case vbYes:    SaveEdits: ConfirmLeave = True
        Case vbNo:     Dirty = False: ConfirmLeave = True
        Case Else:     ConfirmLeave = False
    End Select
End Function

Private Sub MarkDirty()
    Dirty = True
    SetStatus "Unsaved changes to '" & EdName & "'", cWarn
End Sub

Private Sub SetStatus(ByVal msg As String, ByVal clr As Long)
    Ctl("lblStatus").Caption = msg
    Ctl("lblStatus").ForeColor = clr
End Sub

'==============================================================================
' PAGE: PRESETS
'==============================================================================
Private Sub RefreshPresets()
    Dim names As Variant, i As Long, lst As MSForms.ListBox
    Set lst = Ctl("lstPresets")
    names = PresetNames()
    Busy = True
    lst.Clear
    For i = 0 To UBound(names)
        lst.AddItem names(i)
        lst.List(i, 1) = IIf(IsActive(names(i)), "Active", "")
        If StrComp(names(i), EdName, vbTextCompare) = 0 Then lst.ListIndex = i
    Next i
    Ctl("bind:Info:Description").Text = IniGet(Ed, "Info", "Description")
    Busy = False
    EnableButton "btnActivate", Not IsActive(EdName)
    EnableButton "btnDuplicate", PresetCount() < MAX_PRESETS
    EnableButton "btnDelete", PresetCount() > 1
    EnableButton "btnReset", IsShippedPreset(EdName)
    Ctl("lblEditing").Caption = EdName & IIf(IsActive(EdName), vbLf & "(active)", "")
End Sub

Private Sub PresetPicked()
    Dim pick As String
    pick = Ctl("lstPresets").Value
    If StrComp(pick, EdName, vbTextCompare) = 0 Then Exit Sub
    If ConfirmLeave() Then LoadEdit pick Else RefreshPresets
End Sub

Private Sub PresetActivate()
    If Not ConfirmLeave() Then Exit Sub
    ApplyPreset EdName
    RefreshAll
    SetStatus "'" & EdName & "' is now active - its shortcuts are live.", cOk
End Sub

Private Sub PresetDuplicate()
    Dim nm As String
    If Not ConfirmLeave() Then Exit Sub
    nm = Trim$(AskText("Name for the copy of '" & EdName & "':", EdName & " copy"))
    If Len(nm) = 0 Then Exit Sub
    If DuplicatePreset(EdName, nm) Then
        LoadEdit nm
        SetStatus "Created '" & nm & "'.", cOk
    Else
        SetStatus "Could not create '" & nm & "' - the name is taken or you already have " & MAX_PRESETS & ".", cError
    End If
End Sub

Private Sub PresetRename()
    Dim nm As String
    If Not ConfirmLeave() Then Exit Sub
    nm = Trim$(AskText("New name for '" & EdName & "':", EdName))
    If Len(nm) = 0 Or StrComp(nm, EdName, vbBinaryCompare) = 0 Then Exit Sub
    If RenamePreset(EdName, nm) Then
        LoadEdit nm
        SetStatus "Renamed to '" & nm & "'.", cOk
    Else
        SetStatus "A preset called '" & nm & "' already exists.", cError
    End If
End Sub

Private Sub PresetDelete()
    If Ask("Delete preset '" & EdName & "'? This can't be undone.", vbOKCancel + vbExclamation) <> vbOK Then Exit Sub
    Dirty = False
    DeletePreset EdName
    LoadEdit ActivePreset
    SetStatus "Deleted.", cOk
End Sub

Private Sub PresetImport()
    Dim nm As String
    If Not ConfirmLeave() Then Exit Sub
    nm = ImportPreset()
    If Len(nm) > 0 Then LoadEdit nm: SetStatus "Imported '" & nm & "'.", cOk
End Sub

Private Sub PresetExport()
    If Not ConfirmLeave() Then Exit Sub
    If ExportPreset(EdName) Then SetStatus "Exported '" & EdName & "'.", cOk
End Sub

Private Sub PresetReset()
    If Ask("Put '" & EdName & "' back to how it shipped with ILLXL?", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    Dirty = False
    ResetPresetToShipped EdName
    LoadEdit EdName
    SetStatus "'" & EdName & "' reset to shipped settings.", cOk
End Sub

'==============================================================================
' PAGE: NUMBER FORMATS
'==============================================================================
Private Function FmtSection() As String
    FmtSection = Split(FMT_SECS, "|")(Ctl("cboFmt").ListIndex)
End Function

Private Sub RefreshFormats(ByVal keep As Long)
    Dim items As Variant, i As Long, lst As MSForms.ListBox, macros As Variant
    Set lst = Ctl("lstFmt")
    items = SectionList(Ed, FmtSection())
    Busy = True
    lst.Clear
    For i = 0 To UBound(items): lst.AddItem items(i): Next i
    If lst.ListCount > 0 Then lst.ListIndex = Application.Max(0, Application.Min(keep, lst.ListCount - 1))
    Busy = False
    macros = Array("CycleNumberFormat", "CycleCurrencyFormat", "CyclePercentFormat", _
                   "CycleDateFormat", "CycleOtherNumbers", "InsertStaticNow")
    Ctl("lblFmtKey").Caption = "Shortcut: " & NoneIfBlank(IniGet(Ed, "Keys", macros(Ctl("cboFmt").ListIndex)))
    FormatPicked
End Sub

Private Sub FormatPicked()
    Busy = True
    If Ctl("lstFmt").ListIndex >= 0 Then Ctl("txtFmt").Text = Ctl("lstFmt").Value Else Ctl("txtFmt").Text = ""
    Busy = False
    Ctl("lblFmtErr").Visible = False
    Ctl("txtFmt").BorderColor = cLine
    FormatPreview Ctl("txtFmt").Text
End Sub

Private Sub FormatTyped()
    Dim items As Variant, i As Long, t As String
    i = Ctl("lstFmt").ListIndex
    t = Ctl("txtFmt").Text
    If i < 0 Then Exit Sub
    If IsValidFormat(t) Then
        items = SectionList(Ed, FmtSection())
        items(i) = t
        SetSectionList Ed, FmtSection(), items
        Busy = True: Ctl("lstFmt").List(i) = t: Busy = False
        Ctl("lblFmtErr").Visible = False: Ctl("txtFmt").BorderColor = cLine
        FormatPreview t
        MarkDirty
    Else
        Ctl("lblFmtErr").Visible = True: Ctl("txtFmt").BorderColor = cError
    End If
End Sub

Private Sub FormatAdd()
    Dim t As String
    t = Ctl("txtFmt").Text
    If Not IsValidFormat(t) Then t = "General"
    SetSectionList Ed, FmtSection(), ArrAppend(SectionList(Ed, FmtSection()), t)
    RefreshFormats Ctl("lstFmt").ListCount
    MarkDirty
End Sub

Private Sub FormatRemove()
    Dim items As Variant
    items = SectionList(Ed, FmtSection())
    If UBound(items) < 1 Then SetStatus "A cycle needs at least one format.", cError: Exit Sub
    SetSectionList Ed, FmtSection(), ArrRemove(items, Ctl("lstFmt").ListIndex)
    RefreshFormats Ctl("lstFmt").ListIndex
    MarkDirty
End Sub

Private Sub FormatMove(ByVal delta As Long)
    Dim i As Long, items As Variant
    i = Ctl("lstFmt").ListIndex
    items = SectionList(Ed, FmtSection())
    If i < 0 Or i + delta < 0 Or i + delta > UBound(items) Then Exit Sub
    SetSectionList Ed, FmtSection(), ArrSwap(items, i, i + delta)
    RefreshFormats i + delta
    MarkDirty
End Sub

' Sample values run through Excel's TEXT() - what the cell would show
Private Sub FormatPreview(ByVal fmt As String)
    Dim samples As Variant, s As Variant, out As String
    Select Case FmtSection()
        Case "PercentFormats":   samples = Array(0.1234, -0.0567, 0)
        Case "DateFormats":      samples = Array(CDbl(Date))
        Case "TimestampFormats": samples = Array(CDbl(Now))
        Case "OtherFormats":     samples = Array(2026, 12.5)
        Case Else:               samples = Array(1234567.891, -98765.4, 0)
    End Select
    For Each s In samples
        out = out & "  " & Format$(s, "General Number") & "  ->  " & TextOf(s, fmt) & vbLf
    Next s
    Ctl("lblFmtPrev").Caption = out
End Sub

Private Function TextOf(ByVal v As Variant, ByVal fmt As String) As String
    On Error Resume Next
    TextOf = Application.WorksheetFunction.Text(v, fmt)
    If Err.Number <> 0 Then TextOf = "(invalid)"
End Function

'==============================================================================
' PAGE: COLORS & STYLES
'==============================================================================
' Kinds 0-4 are lists of "Name | color(s)"; 5-7 are fixed named settings.
Private Function ClrKind() As Long
    ClrKind = Ctl("cboClr").ListIndex
End Function

Private Function ClrSection() As String
    ClrSection = Split(CLR_SECS, "|")(ClrKind())
End Function

Private Function ClrIsList() As Boolean
    ClrIsList = (ClrKind() <= 4)
End Function

Private Function ClrRows() As Long
    Select Case ClrKind()
        Case 3: ClrRows = 3                 ' input style: fill, font, border
        Case 4: ClrRows = 2                 ' header style: fill, font
        Case Else: ClrRows = 1
    End Select
End Function

Private Function ClrAllowsNone() As Boolean
    ClrAllowsNone = (ClrKind() = 0 Or ClrKind() = 2)  ' fills, tab colors
End Function

' Fixed settings: "key|label" pairs per kind
Private Function ClrFixedKeys() As Variant
    Select Case ClrKind()
        Case 5: ClrFixedKeys = Array("Hardcode|Hardcodes (typed numbers)", "Formula|Formulas", _
                                     "SameSheetLink|Links on the same sheet", "CrossSheetLink|Links to other sheets", _
                                     "ExternalLink|Links to other files")
        Case 6: ClrFixedKeys = Array("Zero|Check = 0  (balanced)", "NonZero|Check <> 0  (out)")
        Case 7: ClrFixedKeys = Array("InsertFill|Header fill", "InsertFont|Header font")
    End Select
End Function

' The editable items as "Name | color | color ..." strings
Private Function ClrItems() As Variant
    Dim keys As Variant, i As Long, out() As Variant
    If ClrIsList() Then ClrItems = SectionList(Ed, ClrSection()): Exit Function
    keys = ClrFixedKeys()
    ReDim out(0 To UBound(keys))
    For i = 0 To UBound(keys)
        out(i) = Split(keys(i), "|")(1) & " | " & IniGet(Ed, ClrSection(), Split(keys(i), "|")(0), "#000000")
    Next i
    ClrItems = out
End Function

Private Function ItemParts(ByVal item As String) As Variant
    Dim p As Variant, i As Long
    p = Split(item, "|")
    For i = 0 To UBound(p): p(i) = Trim$(p(i)): Next i
    If UBound(p) = 0 Then p = Array(p(0), p(0))   ' bare "#RRGGBB" line: name = value
    ItemParts = p
End Function

Private Sub RefreshColors(ByVal keep As Long)
    Dim items As Variant, i As Long, lst As MSForms.ListBox, p As Variant, j As Long, cols As String
    Set lst = Ctl("lstClr")
    items = ClrItems()
    Busy = True
    lst.Clear
    For i = 0 To UBound(items)
        p = ItemParts(items(i))
        cols = ""
        For j = 1 To UBound(p): cols = cols & IIf(j > 1, " / ", "") & p(j): Next j
        lst.AddItem p(0)
        lst.List(i, 1) = cols
    Next i
    If lst.ListCount > 0 Then lst.ListIndex = Application.Max(0, Application.Min(keep, lst.ListCount - 1))
    Busy = False
    For Each p In Array("btnClrAdd", "btnClrDel", "btnClrUp", "btnClrDown")
        EnableButton p, ClrIsList()
    Next p
    EnableButton "btnClrNone", ClrAllowsNone()
    EnableButton "btnClrStripe", (ClrKind() = 0)
    Ctl("txtClrName").Enabled = ClrIsList()
    ColorPicked
End Sub

Private Sub ColorPicked()
    Dim r As Long, p As Variant, caps As Variant, i As Long
    i = Ctl("lstClr").ListIndex
    If i < 0 Then Exit Sub
    p = ItemParts(ClrItems()(i))
    Select Case ClrKind()
        Case 3: caps = Array("Fill", "Font", "Border")
        Case 4: caps = Array("Fill", "Font", "")
        Case Else: caps = Array("Color", "", "")
    End Select
    Busy = True
    Ctl("txtClrName").Text = p(0)
    For r = 0 To 2
        Ctl("lblRow" & r).Caption = caps(r)
        Ctl("lblRow" & r).Visible = (r < ClrRows())
        Ctl("hex:" & r).Visible = (r < ClrRows())
        BtnBg("swa:" & r).Visible = (r < ClrRows()): BtnCap("swa:" & r).Visible = (r < ClrRows())
        If r < ClrRows() Then
            If r + 1 <= UBound(p) Then Ctl("hex:" & r).Text = p(r + 1) Else Ctl("hex:" & r).Text = "#000000"
            PaintSwatch "swa:" & r, Ctl("hex:" & r).Text
            Ctl("hex:" & r).BorderColor = cLine
        End If
    Next r
    Busy = False
    SetColorRow 0
    ColorPreview
End Sub

Private Sub SetColorRow(ByVal r As Long)
    Dim k As Long
    ClrRow = r
    For k = 0 To 2
        Ctl("lblRow" & k).Font.Bold = (k = r)
        BtnBg("swa:" & k).BorderColor = IIf(k = r, cAccent, cLine)
    Next k
End Sub

Private Sub SetHexText(ByVal r As Long, ByVal v As String)
    Ctl("hex:" & r).Text = v                        ' fires HexChanged
End Sub

Private Sub PaletteClick(ByVal k As Long)
    SetHexText ClrRow, PaletteHex(k)
End Sub

Private Sub HexChanged(ByVal r As Long)
    Dim clr As Long
    SetColorRow r
    If HexToColor(Ctl("hex:" & r).Text, clr, ClrAllowsNone() And r = 0) Then
        Ctl("hex:" & r).BorderColor = cLine
        PaintSwatch "swa:" & r, Ctl("hex:" & r).Text
        CommitColorItem
    Else
        Ctl("hex:" & r).BorderColor = cError          ' wait until it is a valid #RRGGBB
    End If
End Sub

Private Sub ColorNameTyped()
    If Not ClrIsList() Then Exit Sub
    If InStr(Ctl("txtClrName").Text, "|") > 0 Then Exit Sub
    CommitColorItem
End Sub

' Write the editor back into the preset (only when every color is valid)
Private Sub CommitColorItem()
    Dim i As Long, r As Long, clr As Long, s As String, items As Variant, keys As Variant
    i = Ctl("lstClr").ListIndex
    If i < 0 Then Exit Sub
    For r = 0 To ClrRows() - 1
        If Not HexToColor(Ctl("hex:" & r).Text, clr, ClrAllowsNone() And r = 0) Then Exit Sub
    Next r
    If ClrIsList() Then
        s = Trim$(Ctl("txtClrName").Text)
        If Len(s) = 0 Then s = "Color " & (i + 1)
        For r = 0 To ClrRows() - 1: s = s & " | " & NormalHex(Ctl("hex:" & r).Text): Next r
        items = SectionList(Ed, ClrSection())
        items(i) = s
        SetSectionList Ed, ClrSection(), items
    Else
        keys = ClrFixedKeys()
        IniSet Ed, ClrSection(), Split(keys(i), "|")(0), NormalHex(Ctl("hex:0").Text)
    End If
    Dim p As Variant, cols As String
    p = ItemParts(ClrItems()(i))
    For r = 1 To UBound(p): cols = cols & IIf(r > 1, " / ", "") & p(r): Next r
    Busy = True
    Ctl("lstClr").List(i, 0) = p(0)
    Ctl("lstClr").List(i, 1) = cols
    Busy = False
    ColorPreview
    MarkDirty
End Sub

Private Function NormalHex(ByVal s As String) As String
    s = Trim$(s)
    If LCase$(s) = "none" Or LCase$(s) = "stripe" Then NormalHex = LCase$(s) Else NormalHex = UCase$(s)
End Function

Private Sub ColorPreview()
    Dim lbl As MSForms.Label, c0 As Long, c1 As Long, c2 As Long, h0 As String
    Set lbl = Ctl("lblClrPrev")
    h0 = LCase$(Trim$(Ctl("hex:0").Text))
    lbl.BackStyle = fmBackStyleOpaque: lbl.BackColor = cBg: lbl.ForeColor = cText
    lbl.BorderStyle = fmBorderStyleSingle: lbl.BorderColor = cLine
    lbl.Font.Bold = False: lbl.Caption = vbLf & "1,234.56     Revenue"
    Call HexToColor(Ctl("hex:0").Text, c0): Call HexToColor(Ctl("hex:1").Text, c1)
    Call HexToColor(Ctl("hex:2").Text, c2)
    Select Case ClrKind()
        Case 0, 2, 6                                    ' fill-type colors
            If h0 = "stripe" Then
                lbl.BackColor = RGB(230, 230, 230)
            ElseIf h0 <> "none" Then
                lbl.BackColor = c0
            End If
        Case 1, 5: lbl.ForeColor = c0                   ' font-type colors
        Case 3: lbl.BackColor = c0: lbl.ForeColor = c1: lbl.BorderColor = c2
        Case 4: lbl.BackColor = c0: lbl.ForeColor = c1: lbl.Font.Bold = True
        Case 7
            Call HexToColor(IniGet(Ed, "Headers", "InsertFill"), c0)
            Call HexToColor(IniGet(Ed, "Headers", "InsertFont"), c1)
            lbl.BackColor = c0: lbl.ForeColor = c1: lbl.Font.Bold = True
    End Select
End Sub

Private Sub ColorAdd()
    Dim t As String
    Select Case ClrKind()
        Case 3: t = "New style | #FFFFCC | #0000FF | #000000"
        Case 4: t = "New header | #1F3864 | #FFFFFF"
        Case Else: t = "New color | #4472C4"
    End Select
    SetSectionList Ed, ClrSection(), ArrAppend(SectionList(Ed, ClrSection()), t)
    RefreshColors Ctl("lstClr").ListCount
    MarkDirty
End Sub

Private Sub ColorRemove()
    Dim items As Variant
    items = SectionList(Ed, ClrSection())
    If UBound(items) < 1 Then SetStatus "A cycle needs at least one entry.", cError: Exit Sub
    SetSectionList Ed, ClrSection(), ArrRemove(items, Ctl("lstClr").ListIndex)
    RefreshColors Ctl("lstClr").ListIndex
    MarkDirty
End Sub

Private Sub ColorMove(ByVal delta As Long)
    Dim i As Long, items As Variant
    i = Ctl("lstClr").ListIndex
    items = SectionList(Ed, ClrSection())
    If i < 0 Or i + delta < 0 Or i + delta > UBound(items) Then Exit Sub
    SetSectionList Ed, ClrSection(), ArrSwap(items, i, i + delta)
    RefreshColors i + delta
    MarkDirty
End Sub

'==============================================================================
' PAGE: FONTS & BORDERS
'==============================================================================
Private Sub RefreshFonts(Optional ByVal keep As Long = 0)
    Dim items As Variant, i As Long, lst As MSForms.ListBox
    Set lst = Ctl("lstFonts")
    items = SectionList(Ed, "Fonts")
    Busy = True
    lst.Clear
    For i = 0 To UBound(items): lst.AddItem items(i): Next i
    If lst.ListCount > 0 Then lst.ListIndex = Application.Max(0, Application.Min(keep, lst.ListCount - 1))
    Busy = False
    FontPicked
End Sub

Private Sub FontPicked()
    On Error Resume Next                            ' font not installed: keep the current one
    If Ctl("lstFonts").ListIndex >= 0 Then Ctl("lblFontPrev").Font.Name = Ctl("lstFonts").Value
End Sub

Private Sub FontAdd()
    Dim f As String
    f = Trim$(Ctl("cboFontAdd").Text)
    If Len(f) = 0 Then Exit Sub
    SetSectionList Ed, "Fonts", ArrAppend(SectionList(Ed, "Fonts"), f)
    RefreshFonts Ctl("lstFonts").ListCount
    MarkDirty
End Sub

' Generic remove / move for the simple numbered lists (Fonts)
Private Sub ListRemove(ByVal sec As String, ByVal lstId As String)
    Dim items As Variant
    items = SectionList(Ed, sec)
    If UBound(items) < 1 Then SetStatus "A cycle needs at least one entry.", cError: Exit Sub
    SetSectionList Ed, sec, ArrRemove(items, Ctl(lstId).ListIndex)
    RefreshFonts Ctl(lstId).ListIndex
    MarkDirty
End Sub

Private Sub ListMove(ByVal sec As String, ByVal lstId As String, ByVal delta As Long)
    Dim i As Long, items As Variant
    i = Ctl(lstId).ListIndex
    items = SectionList(Ed, sec)
    If i < 0 Or i + delta < 0 Or i + delta > UBound(items) Then Exit Sub
    SetSectionList Ed, sec, ArrSwap(items, i, i + delta)
    RefreshFonts i + delta
    MarkDirty
End Sub

' Border steps are one comma list: [Borders] CycleSteps
Private Sub RefreshBorders(ByVal keep As Long)
    Dim steps As Variant, i As Long, lst As MSForms.ListBox
    Set lst = Ctl("lstBorders")
    steps = Split(IniGet(Ed, "Borders", "CycleSteps", "thin,none,medium,hairline"), ",")
    Busy = True
    lst.Clear
    For i = 0 To UBound(steps): lst.AddItem (i + 1) & ".  " & Trim$(steps(i)): Next i
    If lst.ListCount > 0 Then lst.ListIndex = Application.Max(0, Application.Min(keep, lst.ListCount - 1))
    Busy = False
End Sub

Private Sub BorderEdit(ByVal action As String)
    Dim steps As Variant, i As Long, j As Long
    steps = Split(IniGet(Ed, "Borders", "CycleSteps", "thin,none,medium,hairline"), ",")
    For j = 0 To UBound(steps): steps(j) = Trim$(steps(j)): Next j
    i = Ctl("lstBorders").ListIndex
    Select Case action
        Case "add":  steps = ArrAppend(steps, Ctl("cboBdrAdd").Value): i = UBound(steps)
        Case "del"
            If UBound(steps) < 1 Then SetStatus "A cycle needs at least one step.", cError: Exit Sub
            steps = ArrRemove(steps, i)
        Case "up", "down"
            j = i + IIf(action = "up", -1, 1)
            If i < 0 Or j < 0 Or j > UBound(steps) Then Exit Sub
            steps = ArrSwap(steps, i, j): i = j
    End Select
    IniSet Ed, "Borders", "CycleSteps", Join(steps, ",")
    RefreshBorders i
    MarkDirty
End Sub

Private Function InstalledFonts() As Variant
    ' Excel's own font-name box lists every installed font. It isn't always
    ' available (e.g. no ribbon), so any failure falls back to a short list.
    Dim box As Object, i As Long, n As Long, out() As String
    On Error GoTo Fallback
    Set box = Application.CommandBars.FindControl(Id:=1728)
    n = box.ListCount
    If n = 0 Then GoTo Fallback
    ReDim out(0 To n - 1)
    For i = 1 To n: out(i - 1) = box.List(i): Next i
    InstalledFonts = out
    Exit Function
Fallback:
    InstalledFonts = Array("Aptos", "Aptos Narrow", "Arial", "Calibri", "Cambria", "Consolas", _
                           "Georgia", "Poppins", "Segoe UI", "Tahoma", "Times New Roman", "Verdana")
End Function

'==============================================================================
' PAGES: TAGS & HEADERS + preset description  (text boxes saved as typed)
'==============================================================================
Private Sub RefreshBoundTexts()
    Dim k As Variant, p As Variant
    Busy = True
    For Each k In Ctl.Keys
        If Left$(k, 5) = "bind:" Then
            p = Split(k, ":")
            Ctl(k).Text = IniGet(Ed, p(1), p(2))
        End If
    Next k
    Busy = False
End Sub

Private Sub BoundTextChanged(ByVal id As String)
    Dim p As Variant
    p = Split(id, ":")
    IniSet Ed, p(1), p(2), Replace(Replace(Ctl(id).Text, vbCr, " "), vbLf, " ")
    MarkDirty
End Sub

'==============================================================================
' PAGE: SHORTCUTS
'==============================================================================
Private Sub RefreshKeys(ByVal keep As Long)
    Dim i As Long, info As Variant, lst As MSForms.ListBox
    Set lst = Ctl("lstKeys")
    KeyMacros = CommandNames()
    Busy = True
    lst.Clear
    For i = 0 To UBound(KeyMacros)
        info = CommandInfo(KeyMacros(i))
        lst.AddItem info(0)
        lst.List(i, 1) = info(1)
        lst.List(i, 2) = NoneIfBlank(IniGet(Ed, "Keys", KeyMacros(i)))
    Next i
    If lst.ListCount > 0 Then lst.ListIndex = Application.Max(0, Application.Min(keep, lst.ListCount - 1))
    Busy = False
    KeyPicked
End Sub

Private Function CurMacro() As String
    If Ctl("lstKeys").ListIndex >= 0 Then CurMacro = KeyMacros(Ctl("lstKeys").ListIndex)
End Function

Private Sub KeyPicked()
    If Len(CurMacro()) = 0 Then Exit Sub
    Ctl("lblKeyCmd").Caption = CommandInfo(CurMacro())(1)
    Busy = True: Ctl("txtKeyCap").Text = IniGet(Ed, "Keys", CurMacro()): Busy = False
    Ctl("lblKeyMsg").Caption = ""
End Sub

Private Sub KeyCaptured(ByVal combo As String, ByVal Shift As Integer)
    If (Shift And 6) = 0 And Not (combo Like "F#" Or combo Like "F1#") Then
        Ctl("lblKeyMsg").ForeColor = cError
        Ctl("lblKeyMsg").Caption = combo & " on its own would stop that key typing normally - add Ctrl and/or Alt."
        Exit Sub
    End If
    AssignKey CurMacro(), combo
End Sub

' Give macro a key. If another command has it, ask before moving it.
Public Function AssignKey(ByVal macro As String, ByVal combo As String) As Boolean
    Dim code As String, m As Variant, other As String, note As String
    If Len(macro) = 0 Then Exit Function
    code = KeyToOnKey(combo)
    If Len(combo) > 0 And Len(code) = 0 Then
        Ctl("lblKeyMsg").ForeColor = cError: Ctl("lblKeyMsg").Caption = "ILLXL can't use " & combo & "."
        Exit Function
    End If
    If Len(code) > 0 Then
        For Each m In KeyMacros
            If m <> macro Then
                If KeyToOnKey(IniGet(Ed, "Keys", m)) = code Then other = m
            End If
        Next m
    End If
    If Len(other) > 0 Then
        If Ask(combo & " is already used by '" & CommandInfo(other)(1) & "'." & vbLf & vbLf & _
               "Move it to '" & CommandInfo(macro)(1) & "'?  ('" & CommandInfo(other)(1) & _
               "' will have no shortcut.)", vbYesNo + vbQuestion) <> vbYes Then Exit Function
        IniSet Ed, "Keys", other, ""
    End If
    IniSet Ed, "Keys", macro, combo
    RefreshKeys Ctl("lstKeys").ListIndex
    note = ExcelShortcutNote(code)
    Ctl("lblKeyMsg").ForeColor = IIf(Len(note) > 0, cWarn, cOk)
    Ctl("lblKeyMsg").Caption = IIf(Len(combo) > 0, combo & " now runs '" & CommandInfo(macro)(1) & "'.", _
                                   "'" & CommandInfo(macro)(1) & "' has no shortcut.") & note
    MarkDirty
    AssignKey = True
End Function

Private Sub KeyClear()
    AssignKey CurMacro(), ""
End Sub

Private Sub KeyDefault()
    AssignKey CurMacro(), ShippedKey(CurMacro())
End Sub

' A heads-up when a key replaces one of Excel's everyday shortcuts
Private Function ExcelShortcutNote(ByVal code As String) As String
    Dim known As Variant, k As Variant
    known = Array("^c|Copy", "^v|Paste", "^x|Cut", "^z|Undo", "^y|Redo", "^s|Save", "^p|Print", _
                  "^f|Find", "^h|Replace", "^a|Select All", "^b|Bold", "^i|Italic", "^u|Underline", _
                  "^n|New", "^o|Open", "^w|Close", "^d|Fill Down", "^r|Fill Right", "^k|Hyperlink", _
                  "^1|Format Cells", "^+L|Filter", "^+{=}|Insert", "^{-}|Delete", "^;|Today", "^+:|Now")
    For Each k In known
        If StrComp(Split(k, "|")(0), code, vbBinaryCompare) = 0 Then
            ExcelShortcutNote = vbLf & "Note: this replaces Excel's own " & Split(k, "|")(1) & " shortcut."
        End If
    Next k
End Function

' KeyDown -> "Ctrl+Alt+Shift+K" (empty for a bare modifier or an unsupported key)
Public Function ComboText(ByVal keyCode As Long, ByVal Shift As Long) As String
    Dim k As String
    Select Case keyCode
        Case 65 To 90, 48 To 57: k = Chr$(keyCode)
        Case 112 To 123:         k = "F" & (keyCode - 111)
        Case 37: k = "Left"
        Case 38: k = "Up"
        Case 39: k = "Right"
        Case 40: k = "Down"
        Case 8:  k = "Backspace"
        Case 46: k = "Delete"
        Case 45: k = "Insert"
        Case 36: k = "Home"
        Case 35: k = "End"
        Case 33: k = "PgUp"
        Case 34: k = "PgDn"
        Case 13: k = "Enter"
        Case 186: k = ";"
        Case 187: k = "="
        Case 188: k = ","
        Case 189: k = "-"
        Case 190: k = "."
        Case 191: k = "/"
        Case 192: k = "`"
        Case 219: k = "["
        Case 220: k = "\"
        Case 221: k = "]"
        Case 222: k = "'"
        Case Else: Exit Function
    End Select
    ComboText = IIf(Shift And 2, "Ctrl+", "") & IIf(Shift And 4, "Alt+", "") & IIf(Shift And 1, "Shift+", "") & k
End Function

'==============================================================================
' HELPERS
'==============================================================================
Private Function IsActive(ByVal presetName As String) As Boolean
    IsActive = (StrComp(presetName, ActivePreset, vbTextCompare) = 0)
End Function

Private Function NoneIfBlank(ByVal s As String) As String
    NoneIfBlank = IIf(Len(s) = 0, "-", s)
End Function

Private Function Ask(ByVal msg As String, ByVal buttons As VbMsgBoxStyle) As VbMsgBoxResult
    If Quiet Then Ask = QuietAnswer Else Ask = MsgBox(msg, buttons, "ILLXL Settings")
End Function

Private Function AskText(ByVal msg As String, ByVal default As String) As String
    If Quiet Then AskText = QuietText Else AskText = InputBox(msg, "ILLXL Settings", default)
End Function

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
    NewDict.CompareMode = vbTextCompare
End Function

'---- Arrays (0-based copies, so any array from Split / Array works) -------------
Private Function ArrAppend(ByVal a As Variant, ByVal v As Variant) As Variant
    Dim r() As Variant, i As Long, n As Long
    n = UBound(a) - LBound(a) + 1
    ReDim r(0 To n)
    For i = 0 To n - 1: r(i) = a(LBound(a) + i): Next i
    r(n) = v
    ArrAppend = r
End Function

Private Function ArrRemove(ByVal a As Variant, ByVal idx As Long) As Variant
    Dim r() As Variant, i As Long, n As Long
    If UBound(a) - LBound(a) < 1 Then ArrRemove = Array(): Exit Function
    ReDim r(0 To UBound(a) - LBound(a) - 1)
    For i = LBound(a) To UBound(a)
        If i - LBound(a) <> idx Then r(n) = a(i): n = n + 1
    Next i
    ArrRemove = r
End Function

Private Function ArrSwap(ByVal a As Variant, ByVal i As Long, ByVal j As Long) As Variant
    Dim r As Variant, t As Variant
    r = ArrAppend(a, Empty)
    ReDim Preserve r(0 To UBound(r) - 1)
    t = r(i): r(i) = r(j): r(j) = t
    ArrSwap = r
End Function

'==============================================================================
' SELF-TEST  (run by tests\run-tests.ps1 without showing the window)
'==============================================================================
Public Sub GoToPage(ByVal key As String)
    If Pages.Exists(key) Then ShowPage key
End Sub

Public Function SelfTest() As String
    Dim fails As String, items As Variant, n As Long, k As Variant, txt As String
    fails = ""                                      ' real empty string (COM-safe)
    Quiet = True: QuietAnswer = vbYes

    '-- every page builds and opens
    For Each k In Array("presets", "formats", "colors", "fonts", "tags", "keys")
        ShowPage CStr(k)
        If Not Pages(k).Visible Then fails = fails & "page " & k & " not shown; "
    Next k
    If EdName <> ActivePreset Then fails = fails & "not editing the active preset; "

    '-- number formats: edit, reject invalid, add, move, remove
    ShowPage "formats"
    Ctl("cboFmt").ListIndex = 0
    Ctl("lstFmt").ListIndex = 0
    Ctl("txtFmt").Text = "0.000"
    If SectionList(Ed, "NumberFormats")(0) <> "0.000" Then fails = fails & "format edit not stored; "
    Ctl("txtFmt").Text = "0.0"""
    If Not Ctl("lblFmtErr").Visible Then fails = fails & "invalid format not flagged; "
    If SectionList(Ed, "NumberFormats")(0) <> "0.000" Then fails = fails & "invalid format was stored; "
    n = UBound(SectionList(Ed, "NumberFormats"))
    FormatAdd
    If UBound(SectionList(Ed, "NumberFormats")) <> n + 1 Then fails = fails & "format add; "
    FormatMove -1
    FormatRemove
    If UBound(SectionList(Ed, "NumberFormats")) <> n Then fails = fails & "format remove; "
    If InStr(Ctl("lblFmtPrev").Caption, "->") = 0 Then fails = fails & "no preview; "

    '-- colors: hex edit, palette, input style rows
    ShowPage "colors"
    Ctl("cboClr").ListIndex = 0
    Ctl("lstClr").ListIndex = 1
    Ctl("hex:0").Text = "#123456"
    If InStr(SectionList(Ed, "Fills")(1), "#123456") = 0 Then fails = fails & "fill hex edit; "
    Ctl("hex:0").Text = "#12"
    If InStr(SectionList(Ed, "Fills")(1), "#123456") = 0 Then fails = fails & "half-typed hex was stored; "
    PaletteClick 10
    If InStr(SectionList(Ed, "Fills")(1), PaletteHex(10)) = 0 Then fails = fails & "palette click; "
    Ctl("cboClr").ListIndex = 3
    If Not Ctl("hex:2").Visible Then fails = fails & "input styles need 3 color rows; "
    Ctl("cboClr").ListIndex = 5
    Ctl("lstClr").ListIndex = 0
    Ctl("hex:0").Text = "#00AA00"
    If IniGet(Ed, "AutoColor", "Hardcode") <> "#00AA00" Then fails = fails & "auto-color edit; "

    '-- fonts / borders / bound text
    ShowPage "fonts"
    Ctl("cboFontAdd").Text = "Verdana"
    FontAdd
    items = SectionList(Ed, "Fonts")
    If items(UBound(items)) <> "Verdana" Then fails = fails & "font add; "
    Ctl("cboBdrAdd").ListIndex = 4
    BorderEdit "add"
    If Right$(IniGet(Ed, "Borders", "CycleSteps"), 6) <> "double" Then fails = fails & "border add; "
    Ctl("bind:UnitTags:Rate").Text = "[a],[b]"
    If IniGet(Ed, "UnitTags", "Rate") <> "[a],[b]" Then fails = fails & "bound text; "

    '-- shortcuts: capture text, conflict hand-over
    If ComboText(75, 3) <> "Ctrl+Shift+K" Then fails = fails & "combo Ctrl+Shift+K; "
    If ComboText(190, 6) <> "Ctrl+Alt+." Then fails = fails & "combo Ctrl+Alt+.; "
    If ComboText(17, 2) <> "" Then fails = fails & "bare Ctrl should be ignored; "
    RefreshKeys 0
    If Not AssignKey("CycleFont", "Ctrl+Shift+K") Then fails = fails & "assign key; "
    If IniGet(Ed, "Keys", "CycleFont") <> "Ctrl+Shift+K" Then fails = fails & "key not stored; "
    If IniGet(Ed, "Keys", "CycleFill") <> "" Then fails = fails & "conflicting key not cleared; "

    '-- save round-trip
    SaveEdits
    If Dirty Then fails = fails & "still dirty after save; "
    txt = PresetToText(EditLoad(EdName))
    For Each k In Array("0.000", PaletteHex(10), "#00AA00", "Verdana", "double", "[a],[b]", "CycleFont=Ctrl+Shift+K")
        If InStr(txt, k) = 0 Then fails = fails & "saved file missing " & k & "; "
    Next k
    If cfgClrHardcode <> RGB(0, 170, 0) Then fails = fails & "active preset not reloaded on save; "

    '-- duplicate / rename / delete
    QuietText = "UI Test Copy"
    PresetDuplicate
    If EdName <> "UI Test Copy" Then fails = fails & "duplicate; "
    QuietText = "UI Test Renamed"
    PresetRename
    If Not PresetExists("UI Test Renamed") Or PresetExists("UI Test Copy") Then fails = fails & "rename; "
    QuietAnswer = vbOK
    PresetDelete
    If PresetExists("UI Test Renamed") Then fails = fails & "delete; "

    SelfTest = fails
End Function
