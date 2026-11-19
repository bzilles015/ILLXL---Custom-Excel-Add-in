Attribute VB_Name = "modSettings"

Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Module: modSettings (Pxx)
' Purpose:
'   Presets. Every cycle list, color convention and shortcut key lives in a
'   plain-text preset file instead of being hard-coded in the macros.
'
'   Files (UTF-8, editable in Notepad, easy to share):
'     %APPDATA%\ILLXL\presets\<Name>.ini   up to MAX_PRESETS of them
'     %APPDATA%\ILLXL\settings.ini         remembers the active preset
'
'   The shipped presets (Nexcess, Bruce Home, Banker) are also stored on the
'   add-in's hidden _Presets sheet. They are written to disk on first run,
'   and the FIRST shipped preset (the edition's default - Nexcess or Bruce
'   Home) is the fallback for anything a preset
'   leaves out or gets wrong - so a broken file can never break a shortcut.
'
'   Macros read the cfg* variables below (loaded once, fast per keystroke).
'
' Public entry points:
'   P01 - LoadSettings / EnsureSettings     P02 - ApplyPreset
'   P03 - PresetNames                       P04 - ShowSettings  (Ctrl+Alt+Shift+S)
'   P05 - ImportPreset                      P06 - OpenPresetFolder
'   P07 - KeyToOnKey                        P08 - KeyAssignments
'==============================================================================

'==== Constants ====
Public Const MAX_PRESETS As Long = 3
Public Const CLR_NONE    As Long = -1           ' "none"   (no fill / no tab color)
Public Const CLR_STRIPE  As Long = -2           ' "stripe" (diagonal stripe fill)
Public Const BDR_NONE    As Long = 0            ' border cycle step: no line
Public Const BDR_DOUBLE  As Long = 99           ' border cycle step: double line
Private Const PRESET_SHEET As String = "_Presets"

'==== Types ====
Public Type ColorItem                   ' Fills, FontColors, TabColors
    Name  As String
    Color As Long                       ' RGB, CLR_NONE or CLR_STRIPE
End Type

Public Type StyleItem                   ' InputStyles (border used) / HeaderStyles
    Name   As String
    Fill   As Long
    Font   As Long
    Border As Long
End Type

'==== Loaded settings (read by the macros) ====
Public cfgNumFormats    As Variant      ' 0-based String arrays, never empty
Public cfgCurFormats    As Variant
Public cfgPctFormats    As Variant
Public cfgDateFormats   As Variant
Public cfgOtherFormats  As Variant
Public cfgStampFormats  As Variant
Public cfgFonts         As Variant
Public cfgBorderSteps   As Variant      ' Long: xlThin, xlMedium, ... BDR_NONE, BDR_DOUBLE
Public cfgTagsValue     As Variant
Public cfgTagsDuration  As Variant
Public cfgTagsRate      As Variant
Public cfgVarianceLabels As Variant
Public cfgHeaderPrompt  As String

Public cfgFills()        As ColorItem
Public cfgFontColors()   As ColorItem
Public cfgTabColors()    As ColorItem
Public cfgInputStyles()  As StyleItem
Public cfgHeaderStyles() As StyleItem

Public cfgClrHardcode   As Long         ' Auto-Color convention
Public cfgClrFormula    As Long
Public cfgClrSameSheet  As Long
Public cfgClrCrossSheet As Long
Public cfgClrExternal   As Long
Public cfgClrZeroOK     As Long         ' Zero-check CF
Public cfgClrZeroBad    As Long
Public cfgHdrInsertFill As Long         ' Insert Headers / Variance Headers
Public cfgHdrInsertFont As Long

Public ActivePreset     As String

'==== Load state ====
Private Loaded       As Boolean
Private Ini          As Object          ' merged settings: section -> (key -> value)
Private BaseIni      As Object          ' first shipped preset (fallback)
Private Warnings     As String          ' problems found while loading, shown once
Private SettingsRoot As String          ' override for tests (blank = %APPDATA%\ILLXL)
Private QuietLoad    As Boolean         ' tests: no message boxes

'------------------------------------------------------------------------------
' P01 - LoadSettings
'     First shipped preset (the default), then the active preset on top of it.
'     List sections (1=, 2=, ...) replace the whole list; other sections
'     replace key by key, so new ILLXL settings always get a default.
'------------------------------------------------------------------------------
Public Sub LoadSettings()
    Dim path As String
    Warnings = ""
    EnsurePresetFiles

    Set BaseIni = ParseIni(BuiltInText(BasePreset()))
    Set Ini = ParseIni(BuiltInText(BasePreset()))

    ActivePreset = ReadActivePreset()
    path = PresetPath(ActivePreset)
    If Len(Dir$(path)) > 0 Then
        OverlayIni Ini, ParseIni(ReadUtf8(path))
    ElseIf StrComp(ActivePreset, BasePreset(), vbTextCompare) <> 0 Then
        AddWarning "Preset '" & ActivePreset & "' was not found - using " & BasePreset() & "."
        ActivePreset = BasePreset()
    End If

    ReadAllSettings
    Loaded = True
    ResetCycleState

    If Len(Warnings) > 0 And Not QuietLoad Then
        MsgBox "Some settings in preset '" & ActivePreset & "' could not be used." & vbCrLf & _
               "ILLXL defaults were used for these:" & vbCrLf & vbCrLf & Warnings, _
               vbExclamation, "ILLXL Presets"
    End If
End Sub

Public Sub EnsureSettings()
    If Not Loaded Then LoadSettings
End Sub

Public Function SettingsWarnings() As String
    SettingsWarnings = Warnings
End Function

'------------------------------------------------------------------------------
' P02 - ApplyPreset - switch preset, reload, rebind keys
'------------------------------------------------------------------------------
Public Sub ApplyPreset(ByVal presetName As String)
    WriteUtf8 RootDir() & "\settings.ini", "[General]" & vbCrLf & "ActivePreset=" & presetName & vbCrLf
    UnbindAllKeys                                   ' release the old preset's keys first
    LoadSettings
    BindAllKeys
    FlashStatus "ILLXL preset: " & ActivePreset
End Sub

'------------------------------------------------------------------------------
' P03 - PresetNames - preset files on disk (0-based array, may be empty)
'------------------------------------------------------------------------------
Public Function PresetNames() As Variant
    Dim f As String, names As String
    EnsurePresetFiles
    f = Dir$(PresetDir() & "\*.ini")
    Do While Len(f) > 0
        names = names & "|" & Left$(f, Len(f) - 4)
        f = Dir$
    Loop
    If Len(names) = 0 Then PresetNames = Array() Else PresetNames = Split(Mid$(names, 2), "|")
End Function

'------------------------------------------------------------------------------
' P04 - ShowSettings  (Ctrl+Alt+Shift+S) - opens the Settings window
'------------------------------------------------------------------------------
Public Sub ShowSettings()
    EnsureSettings
    frmSettings.Show
End Sub

'------------------------------------------------------------------------------
' P05 - ImportPreset - copy a shared .ini into the presets folder
'     Uses the [Info] Name inside the file. At MAX_PRESETS, asks which to replace.
'     Returns the imported preset's name ("" if cancelled).
'------------------------------------------------------------------------------
Public Function ImportPreset() As String
    Dim src As Variant, txt As String, nm As String, names As Variant, ans As String, i As Long
    ImportPreset = ""
    src = Application.GetOpenFilename("ILLXL preset (*.ini),*.ini", , "Import ILLXL preset")
    If VarType(src) = vbBoolean Then Exit Function   ' cancelled

    txt = ReadUtf8(CStr(src))
    nm = IniValue(ParseIni(txt), "Info", "Name", "")
    If Len(nm) = 0 Then nm = Mid$(src, InStrRev(src, "\") + 1): nm = Left$(nm, Len(nm) - 4)
    nm = SafeName(nm)

    names = PresetNames()
    If Len(Dir$(PresetPath(nm))) > 0 Then
        If MsgBox("Replace your existing preset '" & nm & "'?", vbOKCancel + vbQuestion, _
                  "ILLXL Presets") = vbCancel Then Exit Function
    ElseIf UBound(names) + 1 >= MAX_PRESETS Then
        Dim msg As String
        msg = "ILLXL keeps up to " & MAX_PRESETS & " presets. Which one should '" & nm & "' replace?" & vbCrLf & vbCrLf
        For i = 0 To UBound(names): msg = msg & "   " & (i + 1) & "   " & names(i) & vbCrLf: Next i
        ans = Trim$(InputBox(msg, "ILLXL Presets"))
        If Not IsNumeric(ans) Then Exit Function
        i = CLng(ans) - 1
        If i < 0 Or i > UBound(names) Then Exit Function
        DeletePreset names(i)
    End If

    WriteUtf8 PresetPath(nm), txt
    ImportPreset = nm
End Function

'------------------------------------------------------------------------------
' P06 - OpenPresetFolder
'------------------------------------------------------------------------------
Public Sub OpenPresetFolder()
    EnsurePresetFiles
    Shell "explorer.exe """ & PresetDir() & """", vbNormalFocus
End Sub

'------------------------------------------------------------------------------
' P07 - KeyToOnKey - "Ctrl+Alt+Shift+M" -> "^%+M"   ("" if not valid)
'------------------------------------------------------------------------------
Public Function KeyToOnKey(ByVal friendly As String) As String
    Dim parts As Variant, i As Long, mods As String, k As String
    Dim useCtrl As Boolean, useAlt As Boolean, useShift As Boolean
    KeyToOnKey = ""                                 ' real empty string (safe for COM callers)
    friendly = Trim$(friendly)
    If Len(friendly) = 0 Then Exit Function
    parts = Split(friendly, "+")
    For i = 0 To UBound(parts) - 1
        Select Case LCase$(Trim$(parts(i)))
            Case "ctrl":  useCtrl = True
            Case "alt":   useAlt = True
            Case "shift": useShift = True
            Case Else:    Exit Function                 ' unknown modifier
        End Select
    Next i
    ' fixed order (^ % +) so "Shift+Alt+X" and "Alt+Shift+X" match
    mods = IIf(useCtrl, "^", "") & IIf(useAlt, "%", "") & IIf(useShift, "+", "")

    k = Trim$(parts(UBound(parts)))
    Select Case LCase$(k)
        Case "up", "down", "left", "right", "backspace", "delete", "insert", "home", "end", _
             "pgup", "pgdn", "enter", "tab", "esc", "scrolllock", "numlock", "capslock", "break"
            k = "{" & UCase$(k) & "}"
        Case Else
            If LCase$(k) Like "f#" Or LCase$(k) Like "f1#" Then
                k = "{" & UCase$(k) & "}"               ' F1..F15
            ElseIf Len(k) <> 1 Then
                Exit Function                           ' unknown key name
            ElseIf InStr("+^%~(){}[]=-_", k) > 0 Then
                k = "{" & k & "}"                       ' characters OnKey needs in braces
            ElseIf k Like "[A-Za-z]" Then
                k = IIf(InStr(mods, "+") > 0, UCase$(k), LCase$(k))
            End If
    End Select
    KeyToOnKey = mods & k
End Function

'------------------------------------------------------------------------------
' P08 - KeyAssignments - [Keys] section: macro name -> friendly key
'       DisabledKeyList - [General] DisabledKeys as OnKey codes
'------------------------------------------------------------------------------
Public Function KeyAssignments() As Object
    EnsureSettings
    If Ini.Exists("Keys") Then Set KeyAssignments = Ini("Keys") Else Set KeyAssignments = NewDict()
End Function

Public Function DisabledKeyList() As Variant
    Dim raw As Variant, i As Long, out As String, code As String
    EnsureSettings
    raw = SplitList(IniValue(Ini, "General", "DisabledKeys", ""))
    For i = 0 To UBound(raw)
        code = KeyToOnKey(raw(i))
        If Len(code) > 0 Then out = out & "|" & code         ' bad ones reported by CheckKeys
    Next i
    If Len(out) = 0 Then DisabledKeyList = Array() Else DisabledKeyList = Split(Mid$(out, 2), "|")
End Function

'------------------------------------------------------------------------------
' Test hook - drive the Settings window without showing it ("" = all passed)
'------------------------------------------------------------------------------
Public Function SettingsUiSelfTest() As String
    Dim f As frmSettings
    EnsureSettings
    Set f = New frmSettings                         ' builds every page; nothing is shown
    SettingsUiSelfTest = f.SelfTest()
    Unload f
End Function

' Test hook - open the window on one page without blocking (for screenshots)
Public Sub DemoSettings(ByVal page As String)
    Dim f As Object
    EnsureSettings
    For Each f In VBA.UserForms: Unload f: Next f
    If page = "__close" Then Exit Sub
    Set f = New frmSettings
    f.GoToPage page
    f.Show vbModeless
End Sub

'------------------------------------------------------------------------------
' Test hook - point ILLXL at a temp folder and silence message boxes
'------------------------------------------------------------------------------
Public Sub SetSettingsRoot(ByVal folder As String, Optional ByVal quiet As Boolean = True)
    SettingsRoot = folder
    QuietLoad = quiet
    Loaded = False
End Sub

'==============================================================================
' PRESET EDITING  (used by the Settings window - frmSettings)
'==============================================================================
' The window edits a complete in-memory copy of one preset (EditLoad),
' then writes it back as a tidy, commented .ini (EditSave). Saving the
' active preset reloads it and rebinds the keys straight away.
'==============================================================================

' P09 - EditLoad - the preset as a full settings tree (missing parts filled
'       from the default preset), detached from the live settings.
Public Function EditLoad(ByVal presetName As String) As Object
    Dim data As Object
    Set data = ParseIni(BuiltInText(BasePreset()))
    If Len(Dir$(PresetPath(presetName))) > 0 Then OverlayIni data, ParseIni(ReadUtf8(PresetPath(presetName)))
    IniSet data, "Info", "Name", presetName
    Set EditLoad = data
End Function

' P10 - EditSave - write the tree as presets\<name>.ini; reload if it is live
Public Sub EditSave(ByVal presetName As String, ByVal data As Object)
    IniSet data, "Info", "Name", presetName
    WriteUtf8 PresetPath(presetName), PresetToText(data)
    If StrComp(presetName, ActivePreset, vbTextCompare) = 0 Then ApplyPreset presetName
End Sub

' P11 - Preset file management -------------------------------------------------
Public Function PresetExists(ByVal presetName As String) As Boolean
    PresetExists = Len(Dir$(PresetPath(presetName))) > 0
End Function

Public Function IsShippedPreset(ByVal presetName As String) As Boolean
    IsShippedPreset = Len(BuiltInText(presetName)) > 0
End Function

Public Function PresetCount() As Long
    PresetCount = UBound(PresetNames()) + 1
End Function

' Copy a preset under a new name. False if the name is taken or the limit is hit.
Public Function DuplicatePreset(ByVal srcName As String, ByVal newName As String) As Boolean
    Dim data As Object
    newName = SafeName(newName)
    If PresetExists(newName) Or PresetCount() >= MAX_PRESETS Then Exit Function
    Set data = EditLoad(srcName)
    IniSet data, "Info", "Name", newName
    WriteUtf8 PresetPath(newName), PresetToText(data)
    DuplicatePreset = True
End Function

Public Function RenamePreset(ByVal oldName As String, ByVal newName As String) As Boolean
    Dim data As Object, wasActive As Boolean
    newName = SafeName(newName)
    If PresetExists(newName) Or Not PresetExists(oldName) Then Exit Function
    wasActive = (StrComp(oldName, ActivePreset, vbTextCompare) = 0)
    Set data = EditLoad(oldName)
    IniSet data, "Info", "Name", newName
    WriteUtf8 PresetPath(newName), PresetToText(data)
    Kill PresetPath(oldName)
    If wasActive Then ApplyPreset newName
    RenamePreset = True
End Function

' Deleting the active preset switches to the default (or the first one left).
Public Sub DeletePreset(ByVal presetName As String)
    Dim wasActive As Boolean, names As Variant
    If Not PresetExists(presetName) Then Exit Sub
    wasActive = (StrComp(presetName, ActivePreset, vbTextCompare) = 0)
    Kill PresetPath(presetName)
    If wasActive Then
        names = PresetNames()
        If PresetExists(BasePreset()) Then
            ApplyPreset BasePreset()
        ElseIf UBound(names) >= 0 Then
            ApplyPreset names(0)
        Else
            EnsurePresetFiles: ApplyPreset BasePreset()   ' never leave zero presets
        End If
    End If
End Sub

' Put a shipped preset back to how it came with ILLXL
Public Sub ResetPresetToShipped(ByVal presetName As String)
    If Not IsShippedPreset(presetName) Then Exit Sub
    WriteUtf8 PresetPath(presetName), BuiltInText(presetName)
    If StrComp(presetName, ActivePreset, vbTextCompare) = 0 Then ApplyPreset presetName
End Sub

Public Function ExportPreset(ByVal presetName As String) As Boolean
    Dim dest As Variant
    dest = Application.GetSaveAsFilename(presetName & ".ini", "ILLXL preset (*.ini),*.ini", , "Export ILLXL preset")
    If VarType(dest) = vbBoolean Then Exit Function
    WriteUtf8 CStr(dest), PresetToText(EditLoad(presetName))
    ExportPreset = True
End Function

' P12 - Small accessors for the window ------------------------------------------
Public Function IniGet(ByVal data As Object, ByVal sec As String, ByVal key As String, _
                       Optional ByVal fallback As String = "") As String
    IniGet = IniValue(data, sec, key, fallback)
End Function

Public Sub IniSet(ByVal data As Object, ByVal sec As String, ByVal key As String, ByVal value As String)
    If Not data.Exists(sec) Then Set data(sec) = NewDict()
    data(sec)(key) = value
End Sub

Public Function SectionList(ByVal data As Object, ByVal sec As String) As Variant
    SectionList = ListValues(data, sec)
End Function

' Replace a numbered list section with items (renumbered 1, 2, 3 ...)
Public Sub SetSectionList(ByVal data As Object, ByVal sec As String, ByVal items As Variant)
    Dim d As Object, i As Long
    Set d = NewDict()
    If IsArray(items) Then
        For i = LBound(items) To UBound(items): d(CStr(i - LBound(items) + 1)) = items(i): Next i
    End If
    Set data(sec) = d
End Sub

Public Function ShippedKey(ByVal macro As String) As String
    ShippedKey = IniValue(ParseIni(BuiltInText(BasePreset())), "Keys", macro, "")
End Function

Public Function IsValidFormat(ByVal fmt As String) As Boolean
    Dim probe As Range
    If Len(Trim$(fmt)) = 0 Then Exit Function
    Set probe = ThisWorkbook.Worksheets(PRESET_SHEET).Range("Z1")
    On Error Resume Next
    probe.NumberFormat = fmt
    IsValidFormat = (Err.Number = 0)
    On Error GoTo 0
    probe.NumberFormat = "General"
End Function

Public Function HexToColor(ByVal s As String, ByRef clr As Long, Optional ByVal allowNone As Boolean) As Boolean
    HexToColor = ParseColor(s, allowNone, clr)
End Function

Public Function ColorToHex(ByVal clr As Long) As String
    ColorToHex = "#" & Right$("0" & Hex$(clr And &HFF&), 2) & _
                       Right$("0" & Hex$((clr \ &H100&) And &HFF&), 2) & _
                       Right$("0" & Hex$((clr \ &H10000) And &HFF&), 2)
End Function

' P13 - PresetToText - a settings tree as a tidy, commented .ini -----------------
Public Function PresetToText(ByVal data As Object) As String
    Dim specs As Variant, s As Variant, parts As Variant, out As String, done As Object, sec As Variant
    Set done = NewDict()
    out = "; =============================================================" & vbLf & _
          ";  ILLXL preset  -  edit with Ctrl+Alt+Shift+S (Settings window)" & vbLf & _
          ";  or in Notepad.  Colors are hex #RRGGBB.  Lists run 1, 2, 3" & vbLf & _
          ";  in cycle order.  Anything left out uses the ILLXL defaults." & vbLf & _
          "; =============================================================" & vbLf

    ' Section | note shown above it (canonical order)
    specs = Array("Info|", "General|keys ILLXL switches off because they are hit by accident", _
        "NumberFormats|Ctrl+Shift+1", "CurrencyFormats|Ctrl+Shift+4", "PercentFormats|Ctrl+Shift+5", _
        "DateFormats|Ctrl+Shift+3", "OtherFormats|Ctrl+Shift+8", "TimestampFormats|Ctrl+Shift+N", _
        "Fonts|Ctrl+'", "Fills|Name | #color, none or stripe", "FontColors|Name | #color", _
        "AutoColor|Ctrl+Alt+A", "ZeroCheck|Ctrl+Alt+Shift+Z", _
        "InputStyles|Name | fill | font | dotted border", "HeaderStyles|Name | fill | font  (always bold)", _
        "Headers|Insert Headers / Variance Headers", "TabColors|Name | #color or none", _
        "Borders|thin, none, medium, hairline, thick, double", "UnitTags|comma-separated, in cycle order", _
        "Keys|Ctrl / Alt / Shift + key.  Blank = no shortcut.")
    For Each s In specs
        parts = Split(s, "|", 2)
        If data.Exists(parts(0)) Then
            out = out & SectionText(parts(0), parts(1), data(parts(0)))
            done(parts(0)) = True
        End If
    Next s
    For Each sec In data.Keys                       ' anything unrecognised is kept, not lost
        If Not done.Exists(sec) Then out = out & SectionText(sec, "", data(sec))
    Next sec
    PresetToText = out
End Function

Private Function SectionText(ByVal sec As String, ByVal note As String, ByVal d As Object) As String
    Dim k As Variant, out As String
    out = vbLf
    If Len(note) > 0 Then out = out & "; " & note & vbLf
    out = out & "[" & sec & "]" & vbLf
    For Each k In d.Keys
        out = out & k & "=" & d(k) & vbLf
    Next k
    SectionText = out
End Function

'==============================================================================
' READING SETTINGS INTO cfg* VARIABLES
'==============================================================================
Private Sub ReadAllSettings()
    cfgNumFormats = FormatList("NumberFormats")
    cfgCurFormats = FormatList("CurrencyFormats")
    cfgPctFormats = FormatList("PercentFormats")
    cfgDateFormats = FormatList("DateFormats")
    cfgOtherFormats = FormatList("OtherFormats")
    cfgStampFormats = FormatList("TimestampFormats")
    cfgFonts = TextList("Fonts")

    ColorList "Fills", True, cfgFills
    ColorList "FontColors", False, cfgFontColors
    ColorList "TabColors", True, cfgTabColors
    StyleList "InputStyles", cfgInputStyles
    StyleList "HeaderStyles", cfgHeaderStyles

    cfgClrHardcode = ColorSetting("AutoColor", "Hardcode")
    cfgClrFormula = ColorSetting("AutoColor", "Formula")
    cfgClrSameSheet = ColorSetting("AutoColor", "SameSheetLink")
    cfgClrCrossSheet = ColorSetting("AutoColor", "CrossSheetLink")
    cfgClrExternal = ColorSetting("AutoColor", "ExternalLink")
    cfgClrZeroOK = ColorSetting("ZeroCheck", "Zero")
    cfgClrZeroBad = ColorSetting("ZeroCheck", "NonZero")
    cfgHdrInsertFill = ColorSetting("Headers", "InsertFill")
    cfgHdrInsertFont = ColorSetting("Headers", "InsertFont")

    cfgHeaderPrompt = IniValue(Ini, "Headers", "PromptDefault", "")
    cfgVarianceLabels = SplitSetting("Headers", "VarianceLabels")
    cfgTagsValue = SplitSetting("UnitTags", "Value")
    cfgTagsDuration = SplitSetting("UnitTags", "Duration")
    cfgTagsRate = SplitSetting("UnitTags", "Rate")
    cfgBorderSteps = BorderSteps()
    CheckKeys
End Sub

'--- [Keys] / DisabledKeys sanity check: report problems, never fail -------------
Private Sub CheckKeys()
    Dim seen As Object, macro As Variant, code As String, k As Variant
    Set seen = NewDict()
    For Each k In SplitList(IniValue(Ini, "General", "DisabledKeys", ""))
        If Len(KeyToOnKey(k)) = 0 Then AddWarning "[General] DisabledKeys: '" & k & "' is not a key"
    Next k
    If Not Ini.Exists("Keys") Then Exit Sub
    For Each macro In Ini("Keys").Keys
        code = KeyToOnKey(Ini("Keys")(macro))
        If Not IsCommand(macro) Then
            AddWarning "[Keys] " & macro & ": no such ILLXL command"
        ElseIf Len(Ini("Keys")(macro)) > 0 And Len(code) = 0 Then
            AddWarning "[Keys] " & macro & "=" & Ini("Keys")(macro) & ": key not recognized"
        ElseIf Len(code) > 0 Then
            If seen.Exists(code) Then
                AddWarning "[Keys] " & Ini("Keys")(macro) & " is used by both " & seen(code) & _
                           " and " & macro & " (" & macro & " wins)"
            End If
            seen(code) = macro
        End If
    Next macro
End Sub

'--- Number formats: drops any format Excel rejects ----------------------------
Private Function FormatList(ByVal sec As String) As Variant
    Dim items As Variant, i As Long, out As String, probe As Range
    Set probe = ThisWorkbook.Worksheets(PRESET_SHEET).Range("Z1")
    items = ListValues(Ini, sec)
    For i = 0 To UBound(items)
        On Error Resume Next
        probe.NumberFormat = items(i)
        If Err.Number = 0 Then
            out = out & vbLf & items(i)
        Else
            AddWarning "[" & sec & "] " & (i + 1) & "=" & items(i)
        End If
        On Error GoTo 0
    Next i
    probe.NumberFormat = "General"
    If Len(out) = 0 Then FormatList = ListValues(BaseIni, sec) Else FormatList = Split(Mid$(out, 2), vbLf)
End Function

Private Function TextList(ByVal sec As String) As Variant
    TextList = ListValues(Ini, sec)
    If UBound(TextList) < 0 Then TextList = ListValues(BaseIni, sec)
End Function

'--- "Name | #color" lists ------------------------------------------------------
'     Bad lines are skipped (and reported). A list with nothing usable falls
'     back to the shipped list.
Private Sub ColorList(ByVal sec As String, ByVal allowNone As Boolean, ByRef out() As ColorItem, _
                      Optional ByVal src As Object = Nothing)
    Dim items As Variant, parts As Variant, i As Long, n As Long, clr As Long
    If src Is Nothing Then Set src = Ini
    items = ListValues(src, sec)
    If UBound(items) >= 0 Then
        ReDim out(0 To UBound(items))
        For i = 0 To UBound(items)
            parts = SplitPipe(items(i))
            If ParseColor(parts(UBound(parts)), allowNone, clr) Then
                out(n).Name = Trim$(parts(0)): out(n).Color = clr
                n = n + 1
            Else
                AddWarning "[" & sec & "] " & (i + 1) & "=" & items(i)
            End If
        Next i
    End If
    If n > 0 Then
        ReDim Preserve out(0 To n - 1)
    ElseIf Not src Is BaseIni Then
        ColorList sec, allowNone, out, BaseIni
    Else
        ReDim out(0 To 0): out(0).Name = "Black": out(0).Color = 0
    End If
End Sub

'--- "Name | fill | font | border" lists -------------------------------------
'     Border is optional (black if left out).
Private Sub StyleList(ByVal sec As String, ByRef out() As StyleItem, Optional ByVal src As Object = Nothing)
    Dim items As Variant, parts As Variant, i As Long, n As Long, f1 As Long, f2 As Long, f3 As Long
    If src Is Nothing Then Set src = Ini
    items = ListValues(src, sec)
    If UBound(items) >= 0 Then
        ReDim out(0 To UBound(items))
        For i = 0 To UBound(items)
            parts = SplitPipe(items(i) & " | | |")                 ' pad missing columns
            If ParseColor(parts(1), False, f1) And ParseColor(parts(2), False, f2) Then
                If Not ParseColor(parts(3), False, f3) Then f3 = 0
                out(n).Name = Trim$(parts(0)): out(n).Fill = f1: out(n).Font = f2: out(n).Border = f3
                n = n + 1
            Else
                AddWarning "[" & sec & "] " & (i + 1) & "=" & items(i)
            End If
        Next i
    End If
    If n > 0 Then
        ReDim Preserve out(0 To n - 1)
    ElseIf Not src Is BaseIni Then
        StyleList sec, out, BaseIni
    Else
        ReDim out(0 To 0): out(0).Name = "Plain": out(0).Fill = RGB(255, 255, 255)
    End If
End Sub

'--- Single color: preset value, else shipped value -----------------------------
Private Function ColorSetting(ByVal sec As String, ByVal key As String) As Long
    Dim v As String
    v = IniValue(Ini, sec, key, "")
    If ParseColor(v, False, ColorSetting) Then Exit Function
    If Len(v) > 0 Then AddWarning "[" & sec & "] " & key & "=" & v
    Call ParseColor(IniValue(BaseIni, sec, key, "#000000"), False, ColorSetting)
End Function

'--- Comma list: preset value, else shipped value --------------------------------
Private Function SplitSetting(ByVal sec As String, ByVal key As String) As Variant
    SplitSetting = SplitList(IniValue(Ini, sec, key, ""))
    If UBound(SplitSetting) < 0 Then SplitSetting = SplitList(IniValue(BaseIni, sec, key, ""))
End Function

'--- Border cycle steps -> Excel weights -----------------------------------------
Private Function BorderSteps() As Variant
    Dim raw As Variant, i As Long, out() As Variant, n As Long, w As Long, ok As Boolean
    raw = SplitSetting("Borders", "CycleSteps")
    ReDim out(0 To UBound(raw))
    For i = 0 To UBound(raw)
        ok = True
        Select Case LCase$(raw(i))
            Case "thin":     w = xlThin
            Case "medium":   w = xlMedium
            Case "thick":    w = xlThick
            Case "hairline": w = xlHairline
            Case "none":     w = BDR_NONE
            Case "double":   w = BDR_DOUBLE
            Case Else:       ok = False: AddWarning "[Borders] CycleSteps: " & raw(i)
        End Select
        If ok Then out(n) = w: n = n + 1
    Next i
    If n = 0 Then BorderSteps = Array(xlThin, BDR_NONE, xlMedium, xlHairline): Exit Function
    ReDim Preserve out(0 To n - 1)
    BorderSteps = out
End Function

'==============================================================================
' INI PARSING
'==============================================================================

' Text -> Dictionary(section -> Dictionary(key -> value)), case-insensitive.
' Lines starting with ; are notes. Values are everything after the first "=",
' so number formats full of ; and = are safe.
Private Function ParseIni(ByVal text As String) As Object
    Dim result As Object, sec As Object, ln As Variant, s As String, p As Long
    Set result = NewDict()
    text = Replace(Replace(text, ChrW(&HFEFF), ""), vbCrLf, vbLf)
    For Each ln In Split(text, vbLf)
        s = Trim$(ln)
        If Len(s) = 0 Or Left$(s, 1) = ";" Then
            ' blank or note
        ElseIf Left$(s, 1) = "[" And InStr(s, "]") > 2 Then
            Set sec = NewDict()
            Set result(Trim$(Mid$(s, 2, InStr(s, "]") - 2))) = sec
        ElseIf Not sec Is Nothing Then
            p = InStr(s, "=")
            If p > 1 Then sec(Trim$(Left$(s, p - 1))) = Trim$(Mid$(s, p + 1))
        End If
    Next ln
    Set ParseIni = result
End Function

' Lays a preset over the base. Numbered list sections replace the whole list.
Private Sub OverlayIni(ByVal base As Object, ByVal over As Object)
    Dim sec As Variant, k As Variant
    For Each sec In over.Keys
        If Not base.Exists(sec) Or IsListSection(over(sec)) Then
            Set base(sec) = over(sec)
        Else
            For Each k In over(sec).Keys
                base(sec)(k) = over(sec)(k)
            Next k
        End If
    Next sec
End Sub

Private Function IsListSection(ByVal sec As Object) As Boolean
    If sec.Count > 0 Then IsListSection = IsNumeric(sec.Keys()(0))
End Function

' Values of a numbered section in order 1, 2, 3 ... (gaps allowed). 0-based.
Private Function ListValues(ByVal src As Object, ByVal sec As String) As Variant
    Dim d As Object, i As Long, out As String, found As Boolean
    ListValues = Array()
    If Not src.Exists(sec) Then Exit Function
    Set d = src(sec)
    For i = 1 To 99
        If d.Exists(CStr(i)) Then
            If Len(d(CStr(i))) > 0 Then out = out & vbLf & d(CStr(i)): found = True
        End If
    Next i
    If found Then ListValues = Split(Mid$(out, 2), vbLf)
End Function

Private Function IniValue(ByVal src As Object, ByVal sec As String, ByVal key As String, _
                          ByVal fallback As String) As String
    IniValue = fallback
    If src.Exists(sec) Then
        If src(sec).Exists(key) Then IniValue = src(sec)(key)
    End If
End Function

'==============================================================================
' SMALL HELPERS
'==============================================================================

' "#RRGGBB" -> RGB Long. Also "none" / "stripe" when allowed. False if invalid.
Private Function ParseColor(ByVal s As String, ByVal allowNone As Boolean, ByRef clr As Long) As Boolean
    s = Trim$(s)
    Select Case LCase$(s)
        Case "none":   If allowNone Then clr = CLR_NONE: ParseColor = True
        Case "stripe": If allowNone Then clr = CLR_STRIPE: ParseColor = True
        Case Else
            If Len(s) = 7 And s Like "[#][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                clr = RGB(CLng("&H" & Mid$(s, 2, 2)), CLng("&H" & Mid$(s, 4, 2)), CLng("&H" & Mid$(s, 6, 2)))
                ParseColor = True
            End If
    End Select
End Function

Private Function SplitPipe(ByVal s As String) As Variant
    SplitPipe = Split(s, "|")
End Function

' "a, b ,c" -> Array("a","b","c")   ("" -> empty array)
Private Function SplitList(ByVal s As String) As Variant
    Dim parts As Variant, i As Long
    If Len(Trim$(s)) = 0 Then SplitList = Array(): Exit Function
    parts = Split(s, ",")
    For i = 0 To UBound(parts): parts(i) = Trim$(parts(i)): Next i
    SplitList = parts
End Function

Private Sub AddWarning(ByVal msg As String)
    If Len(Warnings) < 1500 Then Warnings = Warnings & "  - " & msg & vbCrLf
End Sub

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
    NewDict.CompareMode = vbTextCompare
End Function

Private Function SafeName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "")
    Next ch
    SafeName = Trim$(s)
    If Len(SafeName) = 0 Then SafeName = "Imported"
End Function

Private Sub FlashStatus(ByVal msg As String)
    Application.StatusBar = msg
    Application.OnTime Now + TimeSerial(0, 0, 4), "'" & ThisWorkbook.Name & "'!ClearStatus"
End Sub

Public Sub ClearStatus()
    Application.StatusBar = False
End Sub

'==============================================================================
' FILES
'==============================================================================
Private Function RootDir() As String
    If Len(SettingsRoot) > 0 Then RootDir = SettingsRoot Else RootDir = Environ$("APPDATA") & "\ILLXL"
End Function

Private Function PresetDir() As String
    PresetDir = RootDir() & "\presets"
End Function

Private Function PresetPath(ByVal presetName As String) As String
    PresetPath = PresetDir() & "\" & presetName & ".ini"
End Function

' First run: create the folders and write the shipped presets.
' (Only when the folder has no presets - a preset you deleted stays deleted.)
Private Sub EnsurePresetFiles()
    Dim ws As Worksheet, c As Long
    On Error Resume Next
    MkDir RootDir()
    MkDir PresetDir()
    On Error GoTo 0
    If Len(Dir$(PresetDir() & "\*.ini")) > 0 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(PRESET_SHEET)
    For c = 1 To MAX_PRESETS
        If Len(ws.Cells(1, c).Value) > 0 Then WriteUtf8 PresetPath(ws.Cells(1, c).Value), ws.Cells(2, c).Value
    Next c
End Sub

' The edition's default preset = first column of the _Presets sheet
' (Nexcess in the Nexcess build, Bruce Home in the public build)
Public Function BasePreset() As String
    BasePreset = ThisWorkbook.Worksheets(PRESET_SHEET).Cells(1, 1).Value
End Function

' Shipped preset text from the _Presets sheet (row 1 = name, row 2 = file text)
Private Function BuiltInText(ByVal presetName As String) As String
    Dim ws As Worksheet, c As Long
    Set ws = ThisWorkbook.Worksheets(PRESET_SHEET)
    For c = 1 To MAX_PRESETS
        If StrComp(ws.Cells(1, c).Value, presetName, vbTextCompare) = 0 Then
            BuiltInText = ws.Cells(2, c).Value
            Exit Function
        End If
    Next c
End Function

Private Function ReadActivePreset() As String
    Dim path As String
    path = RootDir() & "\settings.ini"
    ReadActivePreset = BasePreset()
    If Len(Dir$(path)) > 0 Then
        ReadActivePreset = IniValue(ParseIni(ReadUtf8(path)), "General", "ActivePreset", BasePreset())
    End If
End Function

Private Function ReadUtf8(ByVal path As String) As String
    With CreateObject("ADODB.Stream")
        .Type = 2: .Charset = "utf-8": .Open
        .LoadFromFile path
        ReadUtf8 = .ReadText
        .Close
    End With
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    With CreateObject("ADODB.Stream")
        .Type = 2: .Charset = "utf-8": .Open
        .WriteText Replace(Replace(text, vbCrLf, vbLf), vbLf, vbCrLf)
        .SaveToFile path, 2                         ' 2 = overwrite
        .Close
    End With
End Sub

