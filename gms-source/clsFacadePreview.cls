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
Private WithEvents mScroll As MSForms.ScrollBar
Private WithEvents mButton As MSForms.CommandButton
Attribute mButton.VB_VarHelpID = -1
Private mOwner As Object

Public Sub Attach(ByVal owner As Object, ByVal control As Object)
    Set mOwner = owner
    Select Case TypeName(control)
        Case "TextBox": Set mText = control
        Case "ComboBox": Set mCombo = control
        Case "CheckBox": Set mCheck = control
        Case "CommandButton": Set mButton = control
        Case "ScrollBar": Set mScroll = control
    End Select
End Sub

Public Sub Detach()
    Set mText = Nothing
    Set mCombo = Nothing
    Set mCheck = Nothing
    Set mButton = Nothing
    Set mScroll = Nothing
    Set mOwner = Nothing
End Sub

Private Sub mText_Change()
    mOwner.InputChanged mText.Name
End Sub
Private Sub mCombo_Change()
    If mCombo.Name = "cmbTemplate" Then
        mOwner.TemplateSelected
    ElseIf mCombo.Name = "cmbCladding" Then
        mOwner.CladdingSelected
    ElseIf mCombo.Name = "cmbSectionNo" Or mCombo.Name = "cmbSectionType" Or _
           mCombo.Name = "cmbEntrance" Or mCombo.Name = "cmbSeed" Then
        mOwner.SectionChoiceChanged mCombo.Name
    Else
        mOwner.RefreshPreview
    End If
End Sub
Private Sub mButton_Click()
    mOwner.SaveCurrentTemplate
End Sub
Private Sub mCheck_Click()
    mOwner.RefreshPreview
End Sub

Private Sub mScroll_Change()
    mOwner.SliderChanged mScroll.Name
End Sub
