Attribute VB_Name = "modMain"
Option Explicit

' ============================================================
' modMain - точка входа, глобальные настройки, запасной мастер
' ============================================================

' --- размеры панели (мм) ---
Public PanelW As Double
Public PanelH As Double
Public MarginMm As Double
Public GapMin As Double

' --- структура здания ---
Public NFloors As Long        ' жилые этажи
Public NPlinth As Long        ' цокольные этажи
Public NSections As Long      ' вертикальные секции
Public NEntrances As Long     ' подъезды (в секции лоджий)
Public FoundationOn As Boolean
Public FoundationH As Double

' --- типы ---
Public RoofType As Long       ' 0 плоская, 1 двускатная, 2 односкатная, 3 нет
Public WindowType As Long     ' 0..4 (4 = балконный блок)
Public RandomStyles As Boolean
Public TextureType As Long    ' 0 кирпич, 1 доски, 2 панели
Public TexA As Double         ' кирпич: длина / доски: ширина / панели: ширина
Public TexB As Double         ' кирпич: высота ряда / доски: длина / панели: высота
Public TexGap As Double       ' зазор между панелями

' --- флаги детализации ---
Public JitterOn As Boolean
Public DetailsOn As Boolean
Public CanopyOn As Boolean
Public Seed As Long

Public FormCancelled As Boolean

' главная точка входа (видна в диспетчере макросов CorelDRAW)
Public Sub FasadPanel()
    ApplyDefaultPreset
    FormCancelled = False

    Dim okForm As Boolean
    okForm = False
    On Error Resume Next
    okForm = ShowFacadeForm()
    On Error GoTo 0

    If Not okForm Then
        If MsgBox("Не удалось создать диалоговое окно (доступ к VBA-проекту ограничен)." & vbCrLf & _
                  "Запустить пошаговый мастер на InputBox?", vbYesNo + vbQuestion, _
                  "Генератор фасада") = vbNo Then Exit Sub
        WizardInputBoxes
        If FormCancelled Then Exit Sub
    End If
    If FormCancelled Then Exit Sub

    If Not ValidateSettings Then Exit Sub
    BuildFacade
End Sub

' пресет по умолчанию: светильник 165x320
Public Sub ApplyDefaultPreset()
    PanelW = 165: PanelH = 320
    MarginMm = 4: GapMin = 0.25
    NFloors = 9: NPlinth = 1: NSections = 5: NEntrances = 1
    FoundationOn = True: FoundationH = 12
    RoofType = 0
    WindowType = 1: RandomStyles = True
    TextureType = 0
    TexA = 4#: TexB = 1.7: TexGap = 0.4
    JitterOn = True: DetailsOn = True: CanopyOn = True
    Seed = 101
End Sub

Public Function ValidateSettings() As Boolean
    ValidateSettings = False
    If PanelW < 40 Or PanelH < 60 Then
        MsgBox "Слишком маленькая панель. Минимум 40x60 мм.", vbExclamation
        Exit Function
    End If
    If NFloors < 1 Then NFloors = 1
    If NPlinth < 0 Then NPlinth = 0
    If NSections < 1 Then NSections = 1
    If NSections > 12 Then NSections = 12
    If NEntrances < 0 Then NEntrances = 0
    If MarginMm < 1 Then MarginMm = 1
    If GapMin < 0.1 Then GapMin = 0.1
    If TexA < GapMin + 0.2 Then TexA = GapMin + 0.2
    If TexB < GapMin + 0.2 Then TexB = GapMin + 0.2
    If Seed <= 0 Then Seed = 101
    ValidateSettings = True
End Function

' запасной пошаговый мастер (если форма недоступна)
Private Sub WizardInputBoxes()
    Dim s As String, parts() As String

    s = InputBox("Размер панели: ширина;высота (мм)", "Мастер фасада", _
                 NumStr(PanelW) & ";" & NumStr(PanelH))
    If Len(s) = 0 Then FormCancelled = True: Exit Sub
    parts = Split(s, ";")
    If UBound(parts) >= 1 Then
        PanelW = Num(parts(0), PanelW): PanelH = Num(parts(1), PanelH)
    End If

    s = InputBox("Этажи: жилые;цокольные;секции;подъезды", "Мастер фасада", _
                 NFloors & ";" & NPlinth & ";" & NSections & ";" & NEntrances)
    If Len(s) = 0 Then FormCancelled = True: Exit Sub
    parts = Split(s, ";")
    If UBound(parts) >= 3 Then
        NFloors = NumL(parts(0), NFloors): NPlinth = NumL(parts(1), NPlinth)
        NSections = NumL(parts(2), NSections): NEntrances = NumL(parts(3), NEntrances)
    End If

    s = InputBox("Фундамент: 1/0; высота (мм)", "Мастер фасада", _
                 IIf(FoundationOn, 1, 0) & ";" & NumStr(FoundationH))
    If Len(s) = 0 Then FormCancelled = True: Exit Sub
    parts = Split(s, ";")
    If UBound(parts) >= 1 Then
        FoundationOn = (NumL(parts(0), 1) = 1)
        FoundationH = Num(parts(1), FoundationH)
    End If

    s = InputBox("Крыша: 0 плоская, 1 двускатная, 2 односкатная, 3 нет", "Мастер фасада", RoofType)
    If Len(s) = 0 Then FormCancelled = True: Exit Sub
    RoofType = NumL(s, RoofType)

    s = InputBox("Окна: 0 одностворчатое, 1 двустворчатое, 2 трёхстворчатое," & vbCrLf & _
                 "3 с форточкой, 4 балконный блок; случайный выбор 1/0", "Мастер фасада", _
                 WindowType & ";" & IIf(RandomStyles, 1, 0))
    If Len(s) = 0 Then FormCancelled = True: Exit Sub
    parts = Split(s, ";")
    If UBound(parts) >= 1 Then
        WindowType = NumL(parts(0), WindowType)
        RandomStyles = (NumL(parts(1), 1) = 1)
    End If

    s = InputBox("Фактура: 0 кирпич, 1 доски, 2 панели;" & vbCrLf & _
                 "размер A;размер B;зазор (мм)", "Мастер фасада", _
                 TextureType & ";" & NumStr(TexA) & ";" & NumStr(TexB) & ";" & NumStr(TexGap))
    If Len(s) = 0 Then FormCancelled = True: Exit Sub
    parts = Split(s, ";")
    If UBound(parts) >= 3 Then
        TextureType = NumL(parts(0), TextureType)
        TexA = Num(parts(1), TexA): TexB = Num(parts(2), TexB): TexGap = Num(parts(3), TexGap)
    End If

    s = InputBox("Флаги: живые линии 1/0; детализация 1/0; козырёк 1/0; seed", _
                 "Мастер фасада", IIf(JitterOn, 1, 0) & ";" & IIf(DetailsOn, 1, 0) & ";" & _
                 IIf(CanopyOn, 1, 0) & ";" & Seed)
    If Len(s) = 0 Then FormCancelled = True: Exit Sub
    parts = Split(s, ";")
    If UBound(parts) >= 3 Then
        JitterOn = (NumL(parts(0), 1) = 1)
        DetailsOn = (NumL(parts(1), 1) = 1)
        CanopyOn = (NumL(parts(2), 1) = 1)
        Seed = NumL(parts(3), Seed)
    End If
End Sub

Private Function NumStr(v As Double) As String
    NumStr = Replace(CStr(v), ",", ".")
End Function
