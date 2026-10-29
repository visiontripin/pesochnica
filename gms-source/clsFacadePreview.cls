VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFacadePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Сохраняется в коллекции формы: события динамических полей не теряются.
Private WithEvents mText As MSForms.TextBox
Attribute mText.VB_VarHelpID = -1
Private WithEvents mCombo As MSForms.ComboBox
Attribute mCombo.VB_VarHelpID = -1
Private WithEvents mCheck As MSForms.CheckBox
Attribute mCheck.VB_VarHelpID = -1
Private mOwner As Object

Public Sub Attach(ByVal owner As Object, ByVal control As Object)
    Set mOwner = owner
    Select Case TypeName(control)
        Case "TextBox": Set mText = control
        Case "ComboBox": Set mCombo = control
        Case "CheckBox": Set mCheck = control
    End Select
End Sub

Public Sub Detach()
    Set mText = Nothing
    Set mCombo = Nothing
    Set mCheck = Nothing
    Set mOwner = Nothing
End Sub

Private Sub mText_Change()
    mOwner.RefreshPreview
End Sub
Private Sub mCombo_Change()
    mOwner.RefreshPreview
End Sub
Private Sub mCheck_Click()
    mOwner.RefreshPreview
End Sub
