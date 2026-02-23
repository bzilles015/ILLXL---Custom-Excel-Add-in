Attribute VB_Name = "modBorders"
Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+



'==============================================================================
' Module: modBorders (Bxx)
' Purpose:
'   Border cycling and quick border application for financial modeling.
'
' Bound shortcuts (see modBindings):
'   B01 – BorderTop              Ctrl+Alt+Shift+Up
'   B02 – BorderBottom           Ctrl+Alt+Shift+Down
'   B03 – BorderLeft             Ctrl+Alt+Shift+Left
'   B04 – BorderRight            Ctrl+Alt+Shift+Right
'   B05 – BordersOutlineInside   Ctrl+Alt+Shift+B
'   B06 – ApplySumBar            Ctrl+Alt+-
'   B07 – ApplyDoubleSumBar      Ctrl+Alt+_
'==============================================================================

'------------------------------------------------------------------------------
' Shared Border Cycle Helper
'------------------------------------------------------------------------------
Private Sub CycleBorder(ByVal edge As XlBordersIndex, _
                        ByRef index As Integer, _
                        ByRef prevAddr As String, _
                        ByVal edgeName As String)
    BeginMacroWithUndo
    If Selection.Address(False, False) <> prevAddr Then index = 0
    prevAddr = Selection.Address(False, False)
    
    Dim idx As Integer: idx = index Mod 4
    With Selection.Borders(edge)
        Select Case idx
          Case 0: .LineStyle = xlContinuous: .ColorIndex = 0: .TintAndShade = 0: .Weight = xlThin
          Case 1: .LineStyle = xlNone
          Case 2: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlMedium
          Case 3: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlHairline
        End Select
    End With
    
    index = index + 1
    LogAction "Bdr" & edgeName & "Cyc" & (idx + 1), prevAddr
    RegisterUndo edgeName & " Border"
End Sub

'------------------------------------------------------------------------------
' B01 – BorderTop  (Ctrl+Alt+Shift+Up)
'------------------------------------------------------------------------------
Sub BorderTop()
    CycleBorder xlEdgeTop, BorderTopIndex, PrevTopAddress, "Top"
End Sub

'------------------------------------------------------------------------------
' B02 – BorderBottom  (Ctrl+Alt+Shift+Down)
'------------------------------------------------------------------------------
Sub BorderBottom()
    CycleBorder xlEdgeBottom, BorderBottomIndex, PrevBottomAddress, "Bottom"
End Sub

'------------------------------------------------------------------------------
' B03 – BorderLeft  (Ctrl+Alt+Shift+Left)
'------------------------------------------------------------------------------
Sub BorderLeft()
    CycleBorder xlEdgeLeft, BorderLeftIndex, PrevLeftAddress, "Left"
End Sub

'------------------------------------------------------------------------------
' B04 – BorderRight  (Ctrl+Alt+Shift+Right)
'------------------------------------------------------------------------------
Sub BorderRight()
    CycleBorder xlEdgeRight, BorderRightIndex, PrevRightAddress, "Right"
End Sub

'------------------------------------------------------------------------------
' B05 – BordersOutlineInside  (Ctrl+Alt+Shift+B)
'     Applies outline + inside borders using medium outline, thin inside.
'------------------------------------------------------------------------------
Sub BordersOutlineInside()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    With Selection
        .Borders.LineStyle = xlNone
        With .Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = 0: End With
        With .Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = 0: End With
        With .Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = 0: End With
        With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = 0: End With
        With .Borders(xlInsideVertical):   .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 0: End With
        With .Borders(xlInsideHorizontal): .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 0: End With
    End With
    LogAction "BdrOutlineInside", Selection.Address(False, False)
    RegisterUndo "Outline Borders"
End Sub

'------------------------------------------------------------------------------
' B06 – ApplySumBar  (Ctrl+Alt+-)
'     Single top border for subtotals - the classic "sum bar"
'------------------------------------------------------------------------------
Sub ApplySumBar()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
    
    LogAction "SumBar", Selection.Address(False, False)
    RegisterUndo "Sum Bar"
End Sub

'------------------------------------------------------------------------------
' B07 – ApplyDoubleSumBar  (Ctrl+Alt+_)
'     Single top border + double bottom border for grand totals
'------------------------------------------------------------------------------
Sub ApplyDoubleSumBar()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
        .ColorIndex = 0
    End With
    
    LogAction "DoubleSumBar", Selection.Address(False, False)
    RegisterUndo "Double Sum Bar"
End Sub



