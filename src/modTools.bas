Attribute VB_Name = "modTools"

Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Module: modTools (Txx)
' Purpose:
'   Model tools - audit, send-ready, formula wrappers, sheet index, updates.
'
'   T01 - AuditHardcodes        formulas with a typed-in number  (=A1*1.05)
'   T02 - AuditInconsistent     formulas that break their row's pattern
'   T03 - AuditExternalLinks    report sheet of every link to another file
'   T04 - SendReady             every sheet to its top-left / freeze-pane cell
'   T05 - ToggleIfError         wrap / unwrap IFERROR(..., preset value)
'   T06 - ToggleRound           wrap / unwrap ROUND(..., preset digits)
'   T07 - BuildSheetIndex       hyperlinked table of contents
'   T08 - CheckForUpdates       newer release on GitHub?
'
'   Audits 1-2 work on the selection, or the whole used area of the sheet
'   when one cell is selected, and SELECT what they find - press Enter / Tab
'   to step through the hits.
'==============================================================================

Public Const ILLXL_REPO As String = "bzilles015/ILLXL---Custom-Excel-Add-in"
Private Const INDEX_TAG As String = "ILLXL Sheet Index"      ' custom property on our index sheet
Private Const LINKS_SHEET As String = "ILLXL Links"

'==============================================================================
' T01 - AuditHardcodes
'==============================================================================
Sub AuditHardcodes()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim f As Range, c As Range, hits As Range, n As Long
    Set f = SpecialCellsIn(AuditScope(), xlCellTypeFormulas)
    If f Is Nothing Then UiMsg "No formulas to check here.": Exit Sub
    For Each c In f.Cells
        If HasTypedNumber(FormulaOf(c)) Then Set hits = AddHit(hits, c): n = n + 1
    Next c
    ReportHits hits, n, "formula(s) contain a typed-in number (like =A1*1.05). 0 and 1 are ignored."
End Sub

' True when the formula holds a numeric constant - outside strings, quoted
' sheet names, table [brackets], cell references and names. 0 and 1 are OK.
Public Function HasTypedNumber(ByVal f As String) As Boolean
    Dim i As Long, j As Long, ch As String, prev As String, nxt As String, depth As Long
    i = 2                                               ' skip the "="
    Do While i <= Len(f)
        ch = Mid$(f, i, 1)
        Select Case ch
            Case """"                                   ' "text" (with "" escapes)
                i = i + 1
                Do While i <= Len(f)
                    If Mid$(f, i, 1) = """" Then
                        If Mid$(f, i + 1, 1) = """" Then i = i + 2 Else Exit Do
                    Else
                        i = i + 1
                    End If
                Loop
                i = i + 1
            Case "'"                                    ' 'Sheet name'!
                j = InStr(i + 1, f, "'!")
                If j = 0 Then Exit Function
                i = j + 2
            Case "["                                    ' table[[#This Row],[Col 2]]
                depth = 0
                Do While i <= Len(f)
                    If Mid$(f, i, 1) = "[" Then depth = depth + 1
                    If Mid$(f, i, 1) = "]" Then depth = depth - 1: If depth = 0 Then Exit Do
                    i = i + 1
                Loop
                i = i + 1
            Case "0" To "9", "."
                prev = Mid$(f, i - 1, 1)
                j = i
                Do While j <= Len(f)
                    If Not Mid$(f, j, 1) Like "[0-9.]" Then Exit Do
                    j = j + 1
                Loop
                If Mid$(f, j, 1) Like "[Ee]" And Mid$(f, j + 1, 1) Like "[+-]" Then  ' 1E+03
                    j = j + 2
                    Do While Mid$(f, j, 1) Like "[0-9]": j = j + 1: Loop
                End If
                nxt = Mid$(f, j, 1)
                ' part of A1, $B$2, Sheet1!, FY2026, 1:1 ... is a reference, not a number
                If Not (prev Like "[A-Za-z_$:!]" Or prev Like "[0-9.]" Or nxt = ":" Or nxt Like "[A-Za-z_(]") Then
                    If Val(Mid$(f, i, j - i)) <> 0 And Val(Mid$(f, i, j - i)) <> 1 Then
                        HasTypedNumber = True
                        Exit Function
                    End If
                End If
                i = j
            Case Else
                i = i + 1
        End Select
    Loop
End Function

'==============================================================================
' T02 - AuditInconsistent
'     Each row's runs of 3+ formulas side by side are compared in R1C1 form
'     (=RC[-1]*1.1 is the "same" formula in every column). Cells that differ
'     from the pattern most of the run uses are flagged.
'==============================================================================
Sub AuditInconsistent()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim area As Range, r As Long, col As Long, runStart As Long, hits As Range, n As Long, ws As Worksheet
    Set ws = ActiveSheet
    For Each area In AuditScope().Areas
        For r = area.Row To area.Row + area.Rows.Count - 1
            runStart = 0
            For col = area.Column To area.Column + area.Columns.Count - 1
                If ws.Cells(r, col).HasFormula Then
                    If runStart = 0 Then runStart = col
                ElseIf runStart > 0 Then
                    CheckRun ws.Range(ws.Cells(r, runStart), ws.Cells(r, col - 1)), hits, n
                    runStart = 0
                End If
            Next col
            If runStart > 0 Then CheckRun ws.Range(ws.Cells(r, runStart), ws.Cells(r, col - 1)), hits, n
        Next r
    Next area
    ReportHits hits, n, "formula(s) break the pattern of the formulas next to them."
End Sub

Private Sub CheckRun(ByVal run As Range, ByRef hits As Range, ByRef n As Long)
    Dim counts As Object, c As Range, k As Variant, best As String, bestN As Long
    If run.Cells.Count < 3 Then Exit Sub
    Set counts = CreateObject("Scripting.Dictionary")
    For Each c In run.Cells: counts(c.FormulaR1C1) = counts(c.FormulaR1C1) + 1: Next c
    For Each k In counts.Keys
        If counts(k) > bestN Then bestN = counts(k): best = k
    Next k
    If bestN < 2 Or bestN * 2 <= run.Cells.Count Then Exit Sub    ' no clear pattern
    For Each c In run.Cells
        If c.FormulaR1C1 <> best Then Set hits = AddHit(hits, c): n = n + 1
    Next c
End Sub

'==============================================================================
' T03 - AuditExternalLinks
'     Lists every cell (and defined name) that points at another file on a
'     report sheet "ILLXL Links", each with a link to jump to it.
'==============================================================================
Sub AuditExternalLinks()
    Dim wb As Workbook, ws As Worksheet, f As Range, c As Range, files As Variant, rep As Worksheet
    Dim rows As Collection, item As Variant, r As Long, nm As Name
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    files = ExternalLinkFiles(wb)
    If IsEmpty(files) Then UiMsg "No links to other files in '" & wb.Name & "'.": Exit Sub

    Set rows = New Collection
    For Each ws In wb.Worksheets
        If ws.Name <> LINKS_SHEET Then
            Set f = SpecialCellsIn(ws.UsedRange, xlCellTypeFormulas)
            If Not f Is Nothing Then
                For Each c In f.Cells
                    If IsExternalFormula(c.Formula, files) Then rows.Add Array(ws.Name, c.Address(False, False), c.Formula)
                Next c
            End If
        End If
    Next ws
    For Each nm In wb.Names
        If IsExternalFormula(nm.RefersTo, files) Then rows.Add Array("(defined name)", nm.Name, nm.RefersTo)
    Next nm

    ' (re)build the report sheet - it belongs to ILLXL, so it is replaced each run
    Application.ScreenUpdating = False
    Set rep = SheetIn(wb, LINKS_SHEET)
    If rep Is Nothing Then
        Set rep = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        rep.Name = LINKS_SHEET
    Else
        rep.Cells.Clear: rep.Hyperlinks.Delete
    End If
    rep.Activate                            ' adding hyperlinks while a chart sheet is active crashes Excel
    rep.Range("A1").Value = "External links in " & wb.Name
    rep.Range("A1").Font.Size = 14: rep.Range("A1").Font.Bold = True
    rep.Range("A2").Value = rows.Count & " link(s) to: " & Join(files, ", ") & "   -   built by ILLXL " & Format$(Now, "m/d/yyyy h:mm")
    rep.Range("A4:C4").Value = Array("Sheet", "Cell", "Formula")
    StyleHeaderRow rep.Range("A4:C4")
    r = 5
    For Each item In rows
        rep.Cells(r, 1).Value = item(0)
        rep.Cells(r, 2).Value = item(1)
        If item(0) <> "(defined name)" Then
            rep.Hyperlinks.Add rep.Cells(r, 2), "", "'" & Replace(item(0), "'", "''") & "'!" & item(1), , item(1)
        End If
        rep.Cells(r, 3).Value = "'" & item(2)                      ' show the formula as text
        r = r + 1
    Next item
    rep.Columns("A:B").AutoFit: rep.Columns("C").ColumnWidth = 90
    rep.Activate: rep.Range("A1").Select
    Application.ScreenUpdating = True
    LogAction "AuditLinks", CStr(rows.Count)
End Sub

'==============================================================================
' T04 - SendReady  (Ctrl+Alt+Shift+Home)
'     Every visible sheet: scrolled to the top, the first cell below / right
'     of its freeze panes selected (A1 when nothing is frozen), optional zoom
'     from the preset. Then the first sheet is shown - ready to email.
'==============================================================================
Sub SendReady()
    Dim wb As Workbook, sh As Object, first As Object
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    For Each sh In wb.Sheets
        If sh.Visible = xlSheetVisible Then
            sh.Activate
            If Len(cfgSendZoom) > 0 Then ActiveWindow.Zoom = CLng(cfgSendZoom)
            If TypeName(sh) = "Worksheet" Then TopOfSheet sh
            If first Is Nothing Then Set first = sh
        End If
    Next sh
    If Not first Is Nothing Then first.Activate
    Application.ScreenUpdating = True
    FlashStatus "ILLXL: every sheet is back at its top-left cell - ready to send."
    LogAction "SendReady", wb.Name
End Sub

Private Sub TopOfSheet(ByVal ws As Worksheet)
    Dim w As Window, home As Range
    Set w = ActiveWindow
    If w.FreezePanes Then
        ' first cell of the scrolling area = top-left of the frozen pane + frozen rows/cols
        Set home = ws.Cells(w.Panes(1).ScrollRow + w.SplitRow, w.Panes(1).ScrollColumn + w.SplitColumn)
        w.Panes(w.Panes.Count).ScrollRow = home.Row
        w.Panes(w.Panes.Count).ScrollColumn = home.Column
    Else
        Set home = ws.Range("A1")
        w.ScrollRow = 1: w.ScrollColumn = 1
    End If
    home.Select
End Sub

'==============================================================================
' T05 - ToggleIfError  (Ctrl+Alt+Shift+F)     T06 - ToggleRound
'     Formula cells in the selection. If every formula is already wrapped,
'     the wrapper comes off; otherwise the unwrapped ones get wrapped:
'       =A1/B1  <->  =IFERROR(A1/B1,0)         (value: preset [Formulas] IfErrorValue)
'       =A1/B1  <->  =ROUND(A1/B1,0)           (digits: preset [Formulas] RoundDigits)
'==============================================================================
Sub ToggleIfError()
    WrapToggle "IFERROR", cfgIfErrorValue, "IFERROR"
End Sub

Sub ToggleRound()
    WrapToggle "ROUND", cfgRoundDigits, "ROUND"
End Sub

Private Sub WrapToggle(ByVal fn As String, ByVal lastArg As String, ByVal undoName As String)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim formulas As Range, c As Range, allWrapped As Boolean, f As String
    Set formulas = SpecialCellsIn(WriteCells(Selection), xlCellTypeFormulas)
    If formulas Is Nothing Then Exit Sub
    allWrapped = True
    For Each c In formulas.Cells
        If Len(FirstArgOf(FormulaOf(c), fn)) = 0 Then allWrapped = False: Exit For
    Next c

    BeginMacroWithUndo formulas
    Application.ScreenUpdating = False
    For Each c In formulas.Cells
        f = FormulaOf(c)
        If allWrapped Then
            SetFormula c, "=" & FirstArgOf(f, fn)                           ' unwrap
        ElseIf Len(FirstArgOf(f, fn)) = 0 Then
            SetFormula c, "=" & fn & "(" & Mid$(f, 2) & "," & lastArg & ")" ' wrap
        End If
    Next c
    Application.ScreenUpdating = True
    LogAction undoName & IIf(allWrapped, "Off", "On"), Selection.Address(False, False)
    RegisterUndo IIf(allWrapped, "Remove ", "Add ") & undoName
End Sub

' "=IFERROR(A1/B1,0)" with fn "IFERROR" -> "A1/B1". "" when f isn't one call to fn.
Public Function FirstArgOf(ByVal f As String, ByVal fn As String) As String
    Dim i As Long, depth As Long, ch As String, inQ As Boolean, comma As Long
    If UCase$(Left$(f, Len(fn) + 2)) <> "=" & UCase$(fn) & "(" Then Exit Function
    depth = 0
    For i = Len(fn) + 2 To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" Then inQ = Not inQ
        If Not inQ Then
            Select Case ch
                Case "(", "{": depth = depth + 1
                Case ")", "}"
                    depth = depth - 1
                    If depth = 0 Then
                        If i <> Len(f) Or comma = 0 Then Exit Function   ' more after the call / no 2nd arg
                        FirstArgOf = Mid$(f, Len(fn) + 3, comma - Len(fn) - 3)
                        Exit Function
                    End If
                Case ","
                    If depth = 1 And comma = 0 Then comma = i
            End Select
        End If
    Next i
End Function

'==============================================================================
' T07 - BuildSheetIndex
'     A hyperlinked table of contents at the front of the workbook. Running
'     it again refreshes the same sheet (ILLXL tags it, so it never touches
'     a sheet of yours that happens to be called "Index").
'==============================================================================
Sub BuildSheetIndex()
    Dim wb As Workbook, idx As Worksheet, sh As Object, r As Long, n As Long, nm As String
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set idx = OurIndex(wb)
    Application.ScreenUpdating = False
    If idx Is Nothing Then
        Set idx = wb.Worksheets.Add(Before:=wb.Sheets(1))
        nm = "Index"
        If Not SheetIn(wb, nm) Is Nothing Then nm = "Sheet Index"
        If Not SheetIn(wb, nm) Is Nothing Then nm = "ILLXL Index"
        idx.Name = nm
        idx.CustomProperties.Add INDEX_TAG, "1"
    Else
        idx.Cells.Clear: idx.Hyperlinks.Delete
    End If
    idx.Activate                            ' adding hyperlinks while a chart sheet is active crashes Excel

    idx.Range("B2").Value = "Index - " & wb.Name
    idx.Range("B2").Font.Size = 16: idx.Range("B2").Font.Bold = True
    idx.Range("B3").Value = "Click a sheet name to jump to it. Built by ILLXL " & Format$(Now, "m/d/yyyy") & " - run Sheet Index again to refresh."
    idx.Range("B3").Font.Color = RGB(107, 114, 128)
    idx.Range("B5:E5").Value = Array("#", "Sheet", "Tab", "Note")
    StyleHeaderRow idx.Range("B5:E5")
    r = 6
    For Each sh In wb.Sheets
        If Not sh Is idx And sh.Name <> LINKS_SHEET Then
            n = n + 1
            idx.Cells(r, 2).Value = n
            If sh.Visible = xlSheetVisible And TypeName(sh) = "Worksheet" Then
                idx.Hyperlinks.Add idx.Cells(r, 3), "", "'" & Replace(sh.Name, "'", "''") & "'!A1", , sh.Name
            Else                                        ' hidden sheets and chart sheets can't be linked to
                idx.Cells(r, 3).Value = sh.Name
                idx.Cells(r, 5).Value = IIf(TypeName(sh) = "Chart", "chart sheet", "hidden")
                If sh.Visible <> xlSheetVisible Then idx.Cells(r, 3).Font.Color = RGB(156, 163, 175)
            End If
            If sh.Tab.ColorIndex <> xlColorIndexNone Then idx.Cells(r, 4).Interior.Color = sh.Tab.Color
            r = r + 1
        End If
    Next sh
    idx.Columns("A").ColumnWidth = 2: idx.Columns("B").ColumnWidth = 5
    idx.Columns("C").ColumnWidth = 40: idx.Columns("D").ColumnWidth = 6: idx.Columns("E").ColumnWidth = 14
    idx.Activate: idx.Range("C6").Select
    ActiveWindow.DisplayGridlines = False               ' (the index sheet only)
    Application.ScreenUpdating = True
    LogAction "SheetIndex", CStr(n)
End Sub

Private Function OurIndex(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet, p As Object
    For Each ws In wb.Worksheets
        For Each p In ws.CustomProperties
            If p.Name = INDEX_TAG Then Set OurIndex = ws: Exit Function
        Next p
    Next ws
End Function

'==============================================================================
' T08 - Updates
'     LatestRelease asks GitHub for the newest release tag (at most once a
'     day; cached in settings.ini). Nothing about you or your files is sent.
'     Switch off with  [General] CheckForUpdates=0  in settings.ini.
'==============================================================================
Sub CheckForUpdates()
    Dim latest As String
    latest = LatestRelease(True)
    If Len(latest) = 0 Then
        UiMsg "Couldn't reach GitHub to check for updates. You have ILLXL v" & ILLXL_VERSION & "."
    ElseIf IsNewerVersion(latest, ILLXL_VERSION) Then
        If UiMsg("ILLXL v" & latest & " is available (you have v" & ILLXL_VERSION & ")." & vbCrLf & vbCrLf & _
                 "Open the download page?", vbYesNo + vbQuestion) = vbYes Then
            ThisWorkbook.FollowHyperlink "https://github.com/" & ILLXL_REPO & "/releases/latest"
        End If
    Else
        UiMsg "You're up to date - ILLXL v" & ILLXL_VERSION & "."
    End If
End Sub

' Quiet check a few seconds after Excel starts (scheduled by Auto_Open)
Public Sub StartupUpdateCheck()
    On Error Resume Next
    If AppSetting("General", "CheckForUpdates", "1") = "0" Then Exit Sub
    If IsNewerVersion(LatestRelease(False), ILLXL_VERSION) Then
        FlashStatus "ILLXL v" & LatestRelease(False) & " is available - ILLXL tab > Check for Updates."
        RefreshRibbon
    End If
End Sub

' "" when unknown. force = ignore today's cached answer.
Public Function LatestRelease(ByVal force As Boolean) As String
    Dim http As Object, body As String, p As Long, q As Long
    LatestRelease = AppSetting("Update", "Latest", "")
    If UiIsTest() Then Exit Function                                ' tests never go online
    If Not force And AppSetting("Update", "Checked", "") = Format$(Date, "yyyy-mm-dd") Then Exit Function
    On Error GoTo Offline
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 3000, 3000, 3000, 3000
    http.Open "GET", "https://api.github.com/repos/" & ILLXL_REPO & "/releases/latest", False
    http.setRequestHeader "User-Agent", "ILLXL-update-check"
    http.setRequestHeader "Accept", "application/vnd.github+json"
    http.send
    If http.Status = 200 Then
        body = http.responseText
        p = InStr(body, """tag_name""")
        If p > 0 Then
            p = InStr(p + 10, body, """") + 1: q = InStr(p, body, """")
            LatestRelease = Replace(LCase$(Mid$(body, p, q - p)), "v", "")
        End If
    End If
    SetAppSetting "Update", "Latest", LatestRelease
    SetAppSetting "Update", "Checked", Format$(Date, "yyyy-mm-dd")
    Exit Function
Offline:
    ' no network / blocked - keep whatever we knew
End Function

' "2.1.0" > "2.0.10"? Numeric, part by part.
Public Function IsNewerVersion(ByVal a As String, ByVal b As String) As Boolean
    Dim pa As Variant, pb As Variant, i As Long, x As Long, y As Long
    If Len(a) = 0 Or Len(b) = 0 Or b = "dev" Then Exit Function
    pa = Split(Replace(LCase$(a), "v", ""), "."): pb = Split(Replace(LCase$(b), "v", ""), ".")
    For i = 0 To 2
        x = 0: y = 0
        If i <= UBound(pa) Then x = Val(pa(i))
        If i <= UBound(pb) Then y = Val(pb(i))
        If x <> y Then IsNewerVersion = (x > y): Exit Function
    Next i
End Function

'==============================================================================
' HELPERS
'==============================================================================
' Selection, or the sheet's used area when only one cell is selected
Private Function AuditScope() As Range
    If Selection.Cells.CountLarge = 1 Then
        Set AuditScope = ActiveSheet.UsedRange
    Else
        Set AuditScope = WorkCells(Selection)           ' whole columns: used area only
        If AuditScope Is Nothing Then Set AuditScope = Selection.Cells(1, 1)
    End If
End Function

Private Function AddHit(ByVal hits As Range, ByVal c As Range) As Range
    If hits Is Nothing Then Set AddHit = c Else Set AddHit = Union(hits, c)
End Function

Private Sub ReportHits(ByVal hits As Range, ByVal n As Long, ByVal what As String)
    If n = 0 Then
        UiMsg "All clear - nothing found."
    Else
        hits.Select
        UiMsg n & " " & what & vbCrLf & vbCrLf & "They're selected - press Enter or Tab to step through them."
    End If
End Sub

Private Function SheetIn(ByVal wb As Workbook, ByVal nm As String) As Object
    On Error Resume Next
    Set SheetIn = wb.Sheets(nm)
End Function

Private Sub StyleHeaderRow(ByVal rng As Range)
    rng.Interior.Color = cfgHdrInsertFill
    rng.Font.Color = cfgHdrInsertFont
    rng.Font.Bold = True
End Sub
