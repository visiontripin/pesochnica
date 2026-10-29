Attribute VB_Name = "modFormRuntime"
Option Explicit

' ============================================================
' modFormRuntime - элементы постоянной формы frmFasad.
' Форма и две кнопки создаются в редакторе VBA один раз.
' Остальные элементы добавляются при её открытии, без VBIDE.
' ============================================================

Public Function ShowFacadeForm() As Boolean
    Dim frm As Object
    On Error GoTo fail
    FormCancelled = True
    Set frm = VBA.UserForms.Add("frmFasad")
    frm.Show
    Unload frm
    Set frm = Nothing
    ShowFacadeForm = True
    Exit Function
fail:
    Dim message As String
    message = "Ошибка " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    On Error GoTo 0
    MsgBox "Не удалось открыть форму frmFasad." & vbCrLf & message, vbExclamation, "Генератор фасада"
    ShowFacadeForm = False
End Function

Public Sub BuildControls(d As Object)
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

    ' Подсказки доступны при наведении мыши на поля и списки.
    SetTip d, "txtW", "Ширина панели в мм; стандарт светильника — 165."
    SetTip d, "txtH", "Высота панели в мм; стандарт — 320."
    SetTip d, "txtMargin", "Отступ фасада от края панели в миллиметрах."
    SetTip d, "txtGap", "Минимальный технологический зазор для луча; проверьте пробой на материале."
    SetTip d, "txtFloors", "Число жилых этажей, не включая цоколь."
    SetTip d, "txtPlinth", "Число цокольных этажей под жилыми; подъезды располагаются на самом нижнем этаже."
    SetTip d, "txtSections", "Число вертикальных секций фасада (1–12)."
    SetTip d, "txtEntrances", "Количество подъездов: справа налево, не более числа секций."
    SetTip d, "chkFoundation", "Включить фундамент под этажами."
    SetTip d, "txtFoundH", "Высота фундамента в миллиметрах."
    SetTip d, "cmbRoof", "Вид крыши; схема справа показывает её условно."
    SetTip d, "cmbWindow", "Тип окон; при смешивании стилей тип может различаться по этажам."
    SetTip d, "cmbTexture", "Тип отделки; размеры элементов задаются полями A и B."
    SetTip d, "txtTexA", "Размер A в мм: длина кирпича, ширина доски или ширина панели."
    SetTip d, "txtTexB", "Размер B в мм: высота кирпича, длина доски или высота панели."
    SetTip d, "txtTexGap", "Зазор облицовочных панелей в миллиметрах."
    SetTip d, "chkRandom", "Смешивать стили окон; на схеме это показано условно."
    SetTip d, "chkJitter", "Небольшая неровность векторных линий."
    SetTip d, "chkDetails", "Добавить инженерные детали и трещины."
    SetTip d, "chkCanopy", "Добавить козырёк подъезда."
    SetTip d, "txtSeed", "Число для воспроизводимого расположения деталей."
    Dim preview As Object
    Set preview = d.Controls.Add("Forms.Frame.1", "fraPreview")
    preview.Caption = "Предпросмотр (схематично)"
    preview.Left = 440: preview.Top = 10: preview.Width = 290: preview.Height = 395
    Dim note As Object
    Set note = d.Controls.Add("Forms.Label.1", "lblPreviewInfo")
    note.Left = 440: note.Top = 408: note.Width = 290: note.Height = 36
    note.Caption = "Схема не заменяет проверку линий перед гравировкой."

    ' Эти две кнопки созданы в конструкторе UserForm, поэтому их события работают.
    With d.Controls("btnOk")
        .Left = 180: .Top = y: .Width = 120: .Height = 24
        .Caption = "Создать фасад"
    End With
    With d.Controls("btnCancel")
        .Left = 320: .Top = y: .Width = 80: .Height = 24
        .Caption = "Отмена"
    End With
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

Private Function NumStrPub(v As Double) As String
    NumStrPub = Replace(CStr(v), ",", ".")
End Function


Private Sub SetTip(ByVal d As Object, ByVal name As String, ByVal text As String)
    d.Controls(name).ControlTipText = text
End Sub
