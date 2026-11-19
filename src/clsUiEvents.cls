VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUiEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  +=============================================================+
'  |   I  L  L  X  L   --   The Illest Excel Add-In             |
'  |   50+ Shortcuts  .  Zero Cost  .  All ILL                   |
'  |   github.com/bzilles015                                     |
'  +=============================================================+

'==============================================================================
' Class: clsUiEvents
' Purpose:
'   The Settings window builds its controls in code. VBA only delivers events
'   from such controls through a WithEvents variable, so each control gets one
'   of these wrappers. Every event is forwarded to the window with the
'   control's Id, and the window decides what to do (frmSettings.OnUi).
'==============================================================================

Public Owner As Object                      ' the frmSettings instance
Public Id As String                         ' e.g. "btnSave", "lstFmt", "nav:keys"

Public WithEvents Lbl As MSForms.Label      ' flat buttons, nav items, swatches
Attribute Lbl.VB_VarHelpID = -1
Public WithEvents Lst As MSForms.ListBox
Attribute Lst.VB_VarHelpID = -1
Public WithEvents Cbo As MSForms.ComboBox
Attribute Cbo.VB_VarHelpID = -1
Public WithEvents Txt As MSForms.TextBox
Attribute Txt.VB_VarHelpID = -1
Public WithEvents Fra As MSForms.Frame      ' page backgrounds (clear hover)
Attribute Fra.VB_VarHelpID = -1

'---- Labels (buttons) ---------------------------------------------------------
Private Sub Lbl_Click()
    Owner.OnUi Id, "Click"
End Sub

Private Sub Lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Owner.OnHover Id
End Sub

'---- Lists / combos / text boxes ----------------------------------------------
Private Sub Lst_Click()
    Owner.OnUi Id, "Click"
End Sub

Private Sub Cbo_Change()
    Owner.OnUi Id, "Change"
End Sub

Private Sub Txt_Change()
    Owner.OnUi Id, "Change"
End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Owner.OnKeyDown Id, KeyCode, Shift
End Sub

'---- Page background -----------------------------------------------------------
Private Sub Fra_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Owner.OnHover ""
End Sub
