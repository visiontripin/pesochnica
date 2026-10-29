Attribute VB_Name = "modFormRuntime"
Option Explicit

' ============================================================
' modFormRuntime - диалоговое окно, собираемое во время запуска.
' В проекте нет design-time UserForm (не нужен бинарный .frx),
' форма создаётся через VBIDE и удаляется после закрытия.
' Если доступ к VBIDE запрещён, FasadPanel() переключается на
' пошаговый мастер на InputBox (modMain.WizardInputBoxes).
' ============================================================

Private Const FORM_NAME As String = "frmFasadDyn"
' Имя проекта в редакторе VBA пользователя; измените при переименовании проекта.
Private Const PROJECT_NAME As String = "FasadPanel100"

Public Function ShowFacadeForm() As Boolean
    Dim prj As Object, comp As Object, existing As Object, frm As Object
    Dim stage As String, problem As String
    Dim created As Boolean, shown As Boolean
    On Error GoTo fail

    stage = "поиск проекта " & PROJECT_NAME
    Set prj = FindProject()
    If prj Is Nothing Then
        Err.Raise vbObjectError + 101, "FasadPanel", _
            "Проект не найден. Проверьте PROJECT_NAME в modFormRuntime."
    End If

    ' Не удаляем автоматически одноимённую форму: это может быть чужой объект.
    stage = "проверка существующей формы"
    On Error Resume Next
    Set existing = prj.VBComponents(FORM_NAME)
    Err.Clear
    On Error GoTo fail
    If Not existing Is Nothing Then
        Err.Raise vbObjectError + 102, "FasadPanel", _
            "В проекте уже есть " & FORM_NAME & ". Проверьте её в редакторе VBA; " & _
            "если это оставшаяся временная форма фасада, удалите её вручную."
    End If

    stage = "создание временной формы"
    Set comp = prj.VBComponents.Add(3)   ' vbext_ct_MSForm
    created = True
    comp.Name = FORM_NAME
    comp.Properties("Caption").Value = "Генератор фасада панельного дома"
    comp.Properties("Width").Value = 472
    comp.Properties("Height").Value = 586

    stage = "создание элементов управления"
    BuildControls comp.Designer
    stage = "запись кода формы"
    comp.CodeModule.AddFromString FormCodeText()

    stage = "открытие формы"
    Set frm = VBA.UserForms.Add(comp.Name)
    frm.Show
    shown = True
    stage = "закрытие и удаление временной формы"
    Unload frm
    Set frm = Nothing
    prj.VBComponents.Remove comp
    created = False
    ShowFacadeForm = True
    Exit Function
fail:
    problem = "Этап: " & stage & vbCrLf & "Ошибка " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    If created Then prj.VBComponents.Remove comp
    On Error GoTo 0
    If shown Then
        MsgBox problem & vbCrLf & "Проверьте, не осталась ли временная форма " & FORM_NAME & _
               " в проекте.", vbExclamation, "Диалог фасада"
        ShowFacadeForm = True
    Else
        MsgBox problem, vbExclamation, "Диалог фасада: диагностика"
        ShowFacadeForm = False
    End If
End Function

Private Function FindProject() As Object
    Dim prj As Object
    For Each prj In Application.VBE.VBProjects
        If StrComp(prj.Name, PROJECT_NAME, vbTextCompare) = 0 Then
            Set FindProject = prj
            Exit Function
        End If
    Next prj
End Function

Private Sub BuildControls(d As Object)
    Dim y As Long
    y = 10

    Header d, "Размер панели (мм)", 10, y: y = y + 20
    AddLabel d, "lblW", "Ширина:", 10, y + 3, 70
    AddText d, "txtW", 85, y, 90, CStr(PanelW)
    AddLabel d, "lblH", "Высота:", 190, y + 3, 60
    AddText d, "txtH", 250, y, 90, CStr(PanelH): y = y + 26

    AddLabel d, "lblMargin", "Поле, мм:", 10, y + 3, 70
    AddText d, "txtMargin", 85, y, 90, NumStrPub(MarginMm)
    AddLabel d, "lblGap", "Зазор луча, мм:", 190, y + 3, 95
    AddText d, "txtGap", 292, y, 80, NumStrPub(GapMin): y = y + 30

    Header d, "Структура здания", 10, y: y = y + 20
    AddLabel d, "lblFloors", "Жилые этажи:", 10, y + 3, 90
    AddText d, "txtFloors", 105, y, 70, CStr(NFloors)
    AddSpin d, "spinFloors", 178, y, 14
    AddLabel d, "lblPlinth", "Цокольные этажи:", 205, y + 3, 110
    AddText d, "txtPlinth", 320, y, 52, CStr(NPlinth): y = y + 26

    AddLabel d, "lblSections", "Секции по ширине:", 10, y + 3, 110
    AddText d, "txtSections", 125, y, 60, CStr(NSections)
    AddLabel d, "lblEntrances", "Подъезды:", 205, y + 3, 70
    AddText d, "txtEntrances", 280, y, 60, CStr(NEntrances): y = y + 26

    AddCheck d, "chkFoundation", "Фундамент", 10, y + 2, 90
    AddLabel d, "lblFoundH", "высота, мм:", 120, y + 3, 80
    AddText d, "txtFoundH", 205, y, 70, NumStrPub(FoundationH): y = y + 30

    Header d, "Типы элементов", 10, y: y = y + 20
    AddLabel d, "lblRoof", "Крыша:", 10, y + 3, 120
    AddCombo d, "cmbRoof", 135, y, 237: y = y + 26
    AddLabel d, "lblWindow", "Форма окон:", 10, y + 3, 120
    AddCombo d, "cmbWindow", 135, y, 170
    AddCheck d, "chkRandom", "Смешивать стили", 320, y + 2, 130: y = y + 26
    AddLabel d, "lblTexture", "Фактура отделки:", 10, y + 3, 120
    AddCombo d, "cmbTexture", 135, y, 237: y = y + 26
    AddLabel d, "lblTA", "Размер A:", 10, y + 3, 60
    AddText d, "txtTexA", 75, y, 70, NumStrPub(TexA)
    AddLabel d, "lblTB", "Размер B:", 160, y + 3, 60
    AddText d, "txtTexB", 225, y, 70, NumStrPub(TexB)
    AddLabel d, "lblTG", "Зазор:", 305, y + 3, 45
    AddText d, "txtTexGap", 352, y, 45, NumStrPub(TexGap): y = y + 30

    Header d, "Детализация и стиль линий", 10, y: y = y + 20
    AddCheck d, "chkJitter", "Живые (неидеальные) линии", 10, y + 2, 180
    AddCheck d, "chkDetails", "Водостоки, кондиционеры, трещины", 205, y + 2, 240: y = y + 24
    AddCheck d, "chkCanopy", "Козырёк подъезда", 10, y + 2, 150
    AddLabel d, "lblSeed", "Seed (рисунок):", 205, y + 3, 110
    AddText d, "txtSeed", 320, y, 60, CStr(Seed): y = y + 32

    AddButton d, "btnPreset", "Пресет 165x320", 10, y, 130
    AddButton d, "btnOk", "Создать фасад", 180, y, 120
    AddButton d, "btnCancel", "Отмена", 320, y, 80
End Sub

Private Sub Header(d As Object, caption As String, x As Long, y As Long)
    Dim c As Object
    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblH" & CStr(y)
    c.Left = x: c.Top = y: c.Width = 420: c.Height = 16
    c.caption = caption
    c.Font.Bold = True
End Sub

Private Sub AddLabel(d As Object, nm As String, caption As String, _
                     x As Long, y As Long, w As Long)
    Dim c As Object
    Set c = d.Controls.Add("Forms.Label.1", nm)
    c.Left = x: c.Top = y: c.Width = w: c.Height = 16
    c.caption = caption
End Sub

Private Sub AddText(d As Object, nm As String, x As Long, y As Long, _
                    w As Long, value As String)
    Dim c As Object
    Set c = d.Controls.Add("Forms.TextBox.1", nm)
    c.Left = x: c.Top = y: c.Width = w: c.Height = 16
    c.Text = value
End Sub

Private Sub AddCombo(d As Object, nm As String, x As Long, y As Long, w As Long)
    Dim c As Object
    Set c = d.Controls.Add("Forms.ComboBox.1", nm)
    c.Left = x: c.Top = y: c.Width = w: c.Height = 18
    c.Style = 2   ' fmStyleDropDownList
End Sub

Private Sub AddCheck(d As Object, nm As String, caption As String, _
                     x As Long, y As Long, w As Long)
    Dim c As Object
    Set c = d.Controls.Add("Forms.CheckBox.1", nm)
    c.Left = x: c.Top = y: c.Width = w: c.Height = 16
    c.caption = caption
End Sub

Private Sub AddSpin(d As Object, nm As String, x As Long, y As Long, w As Long)
    Dim c As Object
    Set c = d.Controls.Add("Forms.SpinButton.1", nm)
    c.Left = x: c.Top = y: c.Width = w: c.Height = 18
End Sub

Private Sub AddButton(d As Object, nm As String, caption As String, _
                      x As Long, y As Long, w As Long)
    Dim c As Object
    Set c = d.Controls.Add("Forms.CommandButton.1", nm)
    c.Left = x: c.Top = y: c.Width = w: c.Height = 24
    c.caption = caption
End Sub

Private Function NumStrPub(v As Double) As String
    NumStrPub = Replace(CStr(v), ",", ".")
End Function

' код, который будет внедрён во временную форму
Private Function FormCodeText() As String
    Dim s As String
    s = "Option Explicit" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    cmbRoof.Clear" & vbCrLf
    s = s & "    cmbRoof.AddItem ""0 - Плоская (парапет, техэтаж)""" & vbCrLf
    s = s & "    cmbRoof.AddItem ""1 - Двускатная""" & vbCrLf
    s = s & "    cmbRoof.AddItem ""2 - Односкатная""" & vbCrLf
    s = s & "    cmbRoof.AddItem ""3 - Без крыши""" & vbCrLf
    s = s & "    cmbWindow.Clear" & vbCrLf
    s = s & "    cmbWindow.AddItem ""0 - Одностворчатое""" & vbCrLf
    s = s & "    cmbWindow.AddItem ""1 - Двустворчатое""" & vbCrLf
    s = s & "    cmbWindow.AddItem ""2 - Трёхстворчатое""" & vbCrLf
    s = s & "    cmbWindow.AddItem ""3 - С форточкой""" & vbCrLf
    s = s & "    cmbWindow.AddItem ""4 - Балконный блок""" & vbCrLf
    s = s & "    cmbTexture.Clear" & vbCrLf
    s = s & "    cmbTexture.AddItem ""0 - Кирпич (ложковая перевязка)""" & vbCrLf
    s = s & "    cmbTexture.AddItem ""1 - Доски (горизонтальная обшивка)""" & vbCrLf
    s = s & "    cmbTexture.AddItem ""2 - Облицовочные панели""" & vbCrLf
    s = s & "    cmbRoof.ListIndex = RoofType" & vbCrLf
    s = s & "    cmbWindow.ListIndex = WindowType" & vbCrLf
    s = s & "    cmbTexture.ListIndex = TextureType" & vbCrLf
    s = s & "    chkFoundation.Value = IIf(FoundationOn, 1, 0)" & vbCrLf
    s = s & "    chkRandom.Value = IIf(RandomStyles, 1, 0)" & vbCrLf
    s = s & "    chkJitter.Value = IIf(JitterOn, 1, 0)" & vbCrLf
    s = s & "    chkDetails.Value = IIf(DetailsOn, 1, 0)" & vbCrLf
    s = s & "    chkCanopy.Value = IIf(CanopyOn, 1, 0)" & vbCrLf
    s = s & "    txtFoundH.Enabled = FoundationOn" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub chkFoundation_Click()" & vbCrLf
    s = s & "    txtFoundH.Enabled = (chkFoundation.Value = 1)" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub spinFloors_SpinUp()" & vbCrLf
    s = s & "    txtFloors.Text = CStr(NumL(txtFloors.Text, 9) + 1)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub spinFloors_SpinDown()" & vbCrLf
    s = s & "    Dim v As Long: v = NumL(txtFloors.Text, 9)" & vbCrLf
    s = s & "    If v > 1 Then txtFloors.Text = CStr(v - 1)" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnPreset_Click()" & vbCrLf
    s = s & "    ApplyDefaultPreset" & vbCrLf
    s = s & "    UserForm_Initialize" & vbCrLf
    s = s & "    txtW.Text = Replace(CStr(PanelW), "","", ""."")" & vbCrLf
    s = s & "    txtH.Text = Replace(CStr(PanelH), "","", ""."")" & vbCrLf
    s = s & "    txtMargin.Text = Replace(CStr(MarginMm), "","", ""."")" & vbCrLf
    s = s & "    txtGap.Text = Replace(CStr(GapMin), "","", ""."")" & vbCrLf
    s = s & "    txtFloors.Text = CStr(NFloors): txtPlinth.Text = CStr(NPlinth)" & vbCrLf
    s = s & "    txtSections.Text = CStr(NSections): txtEntrances.Text = CStr(NEntrances)" & vbCrLf
    s = s & "    txtFoundH.Text = Replace(CStr(FoundationH), "","", ""."")" & vbCrLf
    s = s & "    txtTexA.Text = Replace(CStr(TexA), "","", ""."")" & vbCrLf
    s = s & "    txtTexB.Text = Replace(CStr(TexB), "","", ""."")" & vbCrLf
    s = s & "    txtTexGap.Text = Replace(CStr(TexGap), "","", ""."")" & vbCrLf
    s = s & "    txtSeed.Text = CStr(Seed)" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnOk_Click()" & vbCrLf
    s = s & "    PanelW = Num(txtW.Text, 165): PanelH = Num(txtH.Text, 320)" & vbCrLf
    s = s & "    MarginMm = Num(txtMargin.Text, 4): GapMin = Num(txtGap.Text, 0.25)" & vbCrLf
    s = s & "    NFloors = NumL(txtFloors.Text, 9): NPlinth = NumL(txtPlinth.Text, 1)" & vbCrLf
    s = s & "    NSections = NumL(txtSections.Text, 5)" & vbCrLf
    s = s & "    NEntrances = NumL(txtEntrances.Text, 1)" & vbCrLf
    s = s & "    FoundationOn = (chkFoundation.Value = 1)" & vbCrLf
    s = s & "    FoundationH = Num(txtFoundH.Text, 12)" & vbCrLf
    s = s & "    RoofType = cmbRoof.ListIndex" & vbCrLf
    s = s & "    WindowType = cmbWindow.ListIndex" & vbCrLf
    s = s & "    RandomStyles = (chkRandom.Value = 1)" & vbCrLf
    s = s & "    TextureType = cmbTexture.ListIndex" & vbCrLf
    s = s & "    TexA = Num(txtTexA.Text, 4): TexB = Num(txtTexB.Text, 1.7)" & vbCrLf
    s = s & "    TexGap = Num(txtTexGap.Text, 0.4)" & vbCrLf
    s = s & "    JitterOn = (chkJitter.Value = 1)" & vbCrLf
    s = s & "    DetailsOn = (chkDetails.Value = 1)" & vbCrLf
    s = s & "    CanopyOn = (chkCanopy.Value = 1)" & vbCrLf
    s = s & "    Seed = NumL(txtSeed.Text, 101)" & vbCrLf
    s = s & "    FormCancelled = False" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    FormCancelled = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        FormCancelled = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FormCodeText = s
End Function
