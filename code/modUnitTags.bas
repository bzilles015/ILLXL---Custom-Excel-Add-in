Attribute VB_Name = "modUnitTags"

Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Module: modUnitTags (Uxx)
' Purpose:
'   - Apply uniform bracket unit tags (e.g., [#], [%], [mln $]) to selections.
'   - Detect/replace/remove the last [...] tag in cell text.
'
' IMPROVEMENTS:
'   - Better error handling for empty selections
'   - Cleaner string operations
'   - More consistent behavior across functions
'   - Auto-center alignment when applying unit tags
'==============================================================================

'------------------------------------------------------------------------------
' U01 – CycleUnitTag_Value_Uniform  (Ctrl+Alt+Shift+T)
'     Cycles selection through: [#], [%], [$], [mln $], [thd $], [bn $], [x], [pp], [bps].
'------------------------------------------------------------------------------
Public Sub CycleUnitTag_Value_Uniform()
    ApplyUniformTagCycle Array("[#]", "[%]", "[$]", "[mln $]", "[thd $]", "[bn $]", "[x]", "[pp]", "[bps]"), _
                         "UnitTag_Value_Uniform"
End Sub

'------------------------------------------------------------------------------
' U02 – CycleUnitTag_Duration_Uniform  (Ctrl+Alt+Shift+O)
'     Cycles duration-style tags uniformly across the whole selection.
'------------------------------------------------------------------------------
Public Sub CycleUnitTag_Duration_Uniform()
    ApplyUniformTagCycle Array("[d]", "[m]", "[q]", "[y]"), _
                         "UnitTag_Duration_Uniform"
End Sub

'------------------------------------------------------------------------------
' U03 – CycleUnitTag_Rate_Uniform  (Ctrl+Alt+Shift+P)
'     Cycles rate-style tags uniformly across the whole selection.
'------------------------------------------------------------------------------
Public Sub CycleUnitTag_Rate_Uniform()
    ApplyUniformTagCycle Array("[%/y]", "[$/unit]", "[$/FTE]", "[$/yr]"), _
                         "UnitTag_Rate_Uniform"
End Sub

'------------------------------------------------------------------------------
' U04 – RemoveUnitTag  - (Ctrl+Alt+Shift+Backspace)
'     Remove the final [...] tag from each selected cell
'------------------------------------------------------------------------------
Public Sub RemoveUnitTag()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    
    Dim c As Range
    Application.ScreenUpdating = False
    
    For Each c In Selection.Cells
        ' IMPROVEMENT: Only process cells with actual content
        If Not c.HasFormula And Len(c.Value2) > 0 Then
            c.Value = StripLastBracketTag(CStr(c.Value))
        End If
    Next c
    
    Application.ScreenUpdating = True
    LogAction "UnitTag_Remove", Selection.Address(False, False)
    RegisterUndo "Remove Unit Tag"
End Sub

'================ CORE LOGIC ================

'------------------------------------------------------------------------------
' ApplyUniformTagCycle - Core cycling logic
' IMPROVEMENT: Added screen updating control and better empty cell handling
'------------------------------------------------------------------------------
Private Sub ApplyUniformTagCycle(ByVal tags As Variant, ByVal actionName As String)
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo

    ' 1) Determine the current tag from first nonblank, non-formula cell in selection
    Dim cur As String, idx As Long
    cur = DetectSelectionTag(Selection)
    idx = FindTagIndex(cur, tags)

    ' 2) Compute next tag for ENTIRE selection
    Dim n As Long, nextTag As String
    n = UBound(tags) - LBound(tags) + 1
    If idx = -1 Then
        nextTag = CStr(tags(LBound(tags)))
    Else
        nextTag = CStr(tags((idx - LBound(tags) + 1) Mod n + LBound(tags)))
    End If

    ' 3) Apply nextTag to all non-formula cells
    Dim c As Range
    Application.ScreenUpdating = False
    
    For Each c In Selection.Cells
        ' IMPROVEMENT: Only process cells that have content or will have content
        If Not c.HasFormula Then
            If Len(c.Value2) > 0 Then
                c.Value = ReplaceOrAppendTag(CStr(c.Value), nextTag)
            Else
                ' For empty cells, just add the tag
                c.Value = nextTag
            End If
            ' Center align the cell
            c.HorizontalAlignment = xlCenter
        End If
    Next c
    
    Application.ScreenUpdating = True
    LogAction actionName & ":" & nextTag, Selection.Address(False, False)
    RegisterUndo "Cycle Unit Tag (Uniform)"
End Sub

'------------------------------------------------------------------------------
' FindTagIndex - Return index of tag in list; -1 if not found
'------------------------------------------------------------------------------
Private Function FindTagIndex(ByVal tag As String, ByVal tags As Variant) As Long
    Dim i As Long
    If Len(tag) = 0 Then
        FindTagIndex = -1
        Exit Function
    End If
    
    For i = LBound(tags) To UBound(tags)
        If StrComp(tag, CStr(tags(i)), vbTextCompare) = 0 Then
            FindTagIndex = i
            Exit Function
        End If
    Next i
    
    FindTagIndex = -1
End Function

'------------------------------------------------------------------------------
' DetectSelectionTag - Detect tag from first nonblank, non-formula cell
' Returns "" if none found
'------------------------------------------------------------------------------
Private Function DetectSelectionTag(ByVal rg As Range) As String
    Dim c As Range, s As String
    
    For Each c In rg.Cells
        If Len(c.Value2) > 0 And Not c.HasFormula Then
            s = CStr(c.Value2)
            DetectSelectionTag = ExtractLastBracketTag(s)
            Exit Function
        End If
    Next c
    
    DetectSelectionTag = ""
End Function

'------------------------------------------------------------------------------
' ReplaceOrAppendTag - Replace last [...] or append new tag
' IMPROVEMENT: Better whitespace handling
'------------------------------------------------------------------------------
Private Function ReplaceOrAppendTag(ByVal s As String, ByVal newTag As String) As String
    Dim lb As Long, rb As Long
    
    lb = InStrRev(s, "[")
    rb = InStrRev(s, "]")
    
    If lb > 0 And rb > lb Then
        ' Replace existing tag
        ReplaceOrAppendTag = Trim$(Left$(s, lb - 1)) & " " & newTag & Mid$(s, rb + 1)
        ReplaceOrAppendTag = Trim$(ReplaceOrAppendTag)
    Else
        ' Append new tag
        If Len(Trim$(s)) = 0 Then
            ReplaceOrAppendTag = newTag
        Else
            ReplaceOrAppendTag = Trim$(s) & " " & newTag
        End If
    End If
End Function

'------------------------------------------------------------------------------
' ExtractLastBracketTag - Extract last [...] tag; "" if none
'------------------------------------------------------------------------------
Private Function ExtractLastBracketTag(ByVal s As String) As String
    Dim lb As Long, rb As Long
    
    lb = InStrRev(s, "[")
    rb = InStrRev(s, "]")
    
    If lb > 0 And rb > lb Then
        ExtractLastBracketTag = Mid$(s, lb, rb - lb + 1)
    Else
        ExtractLastBracketTag = ""
    End If
End Function

'------------------------------------------------------------------------------
' StripLastBracketTag - Remove last [...] tag
' IMPROVEMENT: Better whitespace cleanup
'------------------------------------------------------------------------------
Private Function StripLastBracketTag(ByVal s As String) As String
    Dim lb As Long, rb As Long
    
    lb = InStrRev(s, "[")
    rb = InStrRev(s, "]")
    
    If lb > 0 And rb > lb Then
        ' Remove tag and clean up whitespace
        StripLastBracketTag = Trim$(Left$(s, lb - 1) & Mid$(s, rb + 1))
    Else
        StripLastBracketTag = s
    End If
End Function


