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

    ' Списки редактируют одну секцию за раз; скрытые поля хранят совместимый формат пресета.
    AddLabel d, "lblSecTypes", "Секция:", 10, y + 3, 55
    AddCombo d, "cmbSectionNo", 70, y, 65
    AddLabel d, "lblTypeChoice", "Тип:", 145, y + 3, 30
    AddCombo d, "cmbSectionType", 180, y, 192: y = y + 26
    AddLabel d, "lblEntList", "Подъезд секции:", 10, y + 3, 100
    AddCombo d, "cmbEntrance", 115, y, 257: y = y + 26
    AddText d, "txtSecTypes", 0, 0, 1, SectionTypes
    AddText d, "txtEntList", 0, 0, 1, EntranceSections
    d.Controls("txtSecTypes").Visible = False
    d.Controls("txtEntList").Visible = False

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
    AddLabel d, "lblTA", "A:", 10, y + 3, 25
    AddText d, "txtTexA", 35, y, 95, NumStrPub(TexA)
    AddLabel d, "lblTB", "B:", 145, y + 3, 25
    AddText d, "txtTexB", 170, y, 95, NumStrPub(TexB)
    AddLabel d, "lblTG", "Шов:", 280, y + 3, 33
    AddText d, "txtTexGap", 315, y, 95, NumStrPub(TexGap): y = y + 30

    Header d, "Детализация и стиль линий", 10, y: y = y + 20
    AddCheck d, "chkJitter", "Живые (неидеальные) линии", 10, y + 2, 180
    AddCheck d, "chkBalconies", "Балконы", 205, y + 2, 100: y = y + 24
    AddCheck d, "chkAC", "Кондиционеры", 10, y + 2, 120
    AddCheck d, "chkPipes", "Трубы", 135, y + 2, 70
    AddCheck d, "chkCracks", "Трещины", 210, y + 2, 90: y = y + 24
    AddCheck d, "chkCanopy", "Козырёк подъезда", 10, y + 2, 150
    AddLabel d, "lblSeed", "Вариант рисунка:", 205, y + 3, 110
    AddCombo d, "cmbSeed", 320, y, 72
    d.Controls("cmbSeed").Style = 0
    AddText d, "txtSeed", 0, 0, 1, CStr(Seed)
    d.Controls("txtSeed").Visible = False
    y = y + 32

    ' Подсказки доступны при наведении мыши на поля и списки.
    SetTip d, "txtW", "Ширина панели в мм; стандарт светильника — 165."
    SetTip d, "txtH", "Высота панели в мм; стандарт — 320."
    SetTip d, "txtMargin", "Отступ фасада от края панели в миллиметрах."
    SetTip d, "txtGap", "Минимальный технологический зазор для луча; проверьте пробой на материале."
    InstallSlider d, "txtFloors", 1, 20
    InstallSlider d, "txtPlinth", 0, 20
    InstallSlider d, "txtSections", 1, 20
    InstallSlider d, "txtEntrances", 0, 20
    InstallSlider d, "txtFoundH", 1, 20
    InstallSlider d, "txtTexA", 1, 200
    InstallSlider d, "txtTexB", 1, 200
    InstallSlider d, "txtTexGap", 1, 200
    AddCheck d, "chkFourSides", "Четыре стороны на странице", 10, y, 210
    d.Controls("chkFourSides").Value = FourSides
    SetTip d, "chkFourSides", "Четыре варианта с соседними Seed, разложенные с зазором 5 мм."
    y = y + 22
    SetTip d, "txtFloors", "Число жилых этажей, не включая цоколь."
    SetTip d, "txtPlinth", "Число цокольных этажей под жилыми; подъезды располагаются на самом нижнем этаже."
    SetTip d, "txtSections", "Число вертикальных секций фасада (1–20)."
    SetTip d, "txtEntrances", "Количество подъездов: справа налево, не более числа секций."
    SetTip d, "chkFoundation", "Включить фундамент под этажами."
    SetTip d, "txtFoundH", "Высота фундамента в миллиметрах."
    SetTip d, "cmbRoof", "Вид крыши; схема справа показывает её условно."
    SetTip d, "cmbWindow", "Тип окон; при смешивании стилей тип может различаться по этажам."
    SetTip d, "cmbTexture", "Тип отделки; размеры элементов задаются полями A и B."
    SetTip d, "txtTexA", "Размер A: длина кирпича, ширина доски/панели/камня; у волн — шаг линий."
    SetTip d, "txtTexB", "Размер B: высота кирпича/панели/камня или длина доски; на волны не влияет."
    SetTip d, "txtTexGap", "Ширина шва для панелей и камня; для остальных фактур не используется."
    SetTip d, "chkRandom", "Смешивать стили окон; на схеме это показано условно."
    SetTip d, "chkJitter", "Небольшая неровность векторных линий."
    SetTip d, "cmbSectionNo", "Выберите номер секции слева направо."
    SetTip d, "cmbSectionType", "Выберите тип для выбранной секции."
    SetTip d, "cmbEntrance", "Выберите секцию по номеру и укажите: есть ли в ней подъезд на нижнем этаже. Изменение количества восстанавливает автоматическое размещение справа."
    SetTip d, "chkAC", "Блоки только под окнами и при достаточном свободном месте."
    SetTip d, "chkPipes", "Включить водосточные трубы."
    SetTip d, "chkCracks", "Включить трещины."
    SetTip d, "chkBalconies", "Балконные секции; при отключении заменяются окнами."
    SetTip d, "chkCanopy", "Добавить козырёк подъезда."
    SetTip d, "cmbSeed", "Выберите число или введите своё: одинаковые числа дают одинаковый рисунок."
    Dim sliderName As Variant
    For Each sliderName In Array("Floors", "Plinth", "Sections", "Entrances", "FoundH", "TexA", "TexB", "TexGap")
        d.Controls("scr" & CStr(sliderName)).ControlTipText = d.Controls("txt" & CStr(sliderName)).ControlTipText
    Next sliderName
    Dim preview As Object
    Set preview = d.Controls.Add("Forms.Frame.1", "fraPreview")
    preview.Caption = "Предпросмотр (схематично)"
    preview.Left = 440: preview.Top = 10: preview.Width = 290: preview.Height = 395
    Dim note As Object
    Set note = d.Controls.Add("Forms.Label.1", "lblPreviewInfo")
    note.Left = 440: note.Top = 408: note.Width = 290: note.Height = 36
    note.Caption = "Схема не заменяет проверку линий перед гравировкой."

    AddLabel d, "lblTemplates", "Шаблон фасада:", 440, 449, 125
    AddCombo d, "cmbTemplate", 440, 468, 198
    Dim saveButton As Object
    Set saveButton = d.Controls.Add("Forms.CommandButton.1", "btnSaveTemplate")
    saveButton.Left = 645: saveButton.Top = 468: saveButton.Width = 85: saveButton.Height = 18
    saveButton.Caption = "Сохранить"
    AddLabel d, "lblCladding", "Пресет отделки:", 440, 492, 125
    AddCombo d, "cmbCladding", 440, 510, 290
    SetTip d, "cmbTemplate", "Шесть шаблонов здания и ваши сохранённые пресеты."
    SetTip d, "btnSaveTemplate", "Сохранить настройки под именем для текущего пользователя Windows."
    SetTip d, "cmbCladding", "Стартовые варианты кирпича, досок и панелей с размерами в миллиметрах."

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

Private Sub InstallSlider(ByVal d As Object, ByVal textName As String, ByVal minimum As Long, ByVal maximum As Long)
    Dim txt As Object, bar As Object, label As Object, suffix As String
    Set txt = d.Controls(textName)
    suffix = Mid$(textName, 4)
    Set bar = d.Controls.Add("Forms.ScrollBar.1", "scr" & suffix)
    bar.Left = txt.Left: bar.Top = txt.Top: bar.Width = MaxD(20, txt.Width - 25): bar.Height = 16
    bar.Min = minimum: bar.Max = maximum: bar.SmallChange = 1: bar.LargeChange = 1
    bar.Orientation = 1
    bar.ControlTipText = "Перемещайте ползунок; текущее значение показано справа."
    Set label = d.Controls.Add("Forms.Label.1", "val" & suffix)
    label.Left = bar.Left + bar.Width + 2: label.Top = bar.Top + 1
    label.Width = 26: label.Height = 16
    txt.Visible = False
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
