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
Public TextureType As Long    ' 0 кирпич, 1 доски, 2 панели, 3 каменные блоки
Public TexA As Double         ' кирпич: длина / доски: ширина / панели: ширина
Public TexB As Double         ' кирпич: высота ряда / доски: длина / панели: высота
Public TexGap As Double       ' зазор между панелями

' --- флаги детализации ---
Public JitterOn As Boolean
Public DetailsOn As Boolean
Public CanopyOn As Boolean
Public ACOn As Boolean, PipesOn As Boolean, CracksOn As Boolean, BalconiesOn As Boolean
' Пустая строка означает автоматическую схему; иначе номера через точку с запятой.
Public SectionTypes As String, EntranceSections As String
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
        If MsgBox("Не удалось открыть форму frmFasad." & vbCrLf & _
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
    TexA = 4#: TexB = 1.7: TexGap = 0.5
    JitterOn = True: DetailsOn = True: CanopyOn = True
    Seed = 101
    ACOn = True: PipesOn = True: CracksOn = True: BalconiesOn = True
    SectionTypes = "": EntranceSections = ""
End Sub

Public Function ValidateSettings() As Boolean
    ValidateSettings = False
    If PanelW < 40 Or PanelH < 60 Then
        MsgBox "Слишком маленькая панель. Минимум 40x60 мм.", vbExclamation
        Exit Function
    End If
    If NFloors < 1 Then NFloors = 1
    If NPlinth < 0 Then NPlinth = 0
    If NFloors + NPlinth > 100 Then
        MsgBox "Суммарное число этажей не должно превышать 100.", vbExclamation
        Exit Function
    End If
    If NSections < 1 Then NSections = 1
    If NSections > 12 Then NSections = 12
    If NEntrances < 0 Then NEntrances = 0
    If NEntrances > NSections Then NEntrances = NSections
    If Not ValidateSectionLists() Then Exit Function
    If MarginMm < 1 Then MarginMm = 1
    If GapMin < 0.1 Then GapMin = 0.1
    If TexA < GapMin + 0.2 Then TexA = GapMin + 0.2
    If TexB < GapMin + 0.2 Then TexB = GapMin + 0.2
    If TextureType = 2 Or TextureType = 3 Then
        If TexGap < GapMin + 0.2 Or TexGap >= TexA Or TexGap >= TexB Then
            MsgBox "Для панелей и камня зазор должен быть не менее GapMin + 0,2 мм и меньше размеров A и B.", vbExclamation
            Exit Function
        End If
    End If
    If Seed <= 0 Then Seed = 101
    ValidateSettings = True
End Function

Public Function SectionTypeAt(ByVal position As Long, ByVal spec As String) As Long
    Dim values As Variant
    SectionTypeAt = (position - 1) Mod 5
    If Len(Trim$(spec)) = 0 Then Exit Function
    values = Split(spec, ";")
    If position - 1 <= UBound(values) Then
        If IsNumeric(Trim$(values(position - 1))) Then SectionTypeAt = CLng(Trim$(values(position - 1)))
    End If
End Function

Public Function EntranceAt(ByVal position As Long, ByVal count As Long, _
                           ByVal automaticCount As Long, ByVal spec As String) As Boolean
    Dim values As Variant, v As Variant
    If Len(Trim$(spec)) = 0 Then
        EntranceAt = (position > count - automaticCount)
        Exit Function
    End If
    values = Split(spec, ";")
    For Each v In values
        If IsNumeric(Trim$(CStr(v))) Then
            If CDbl(Trim$(CStr(v))) = position Then EntranceAt = True: Exit Function
        End If
    Next v
End Function

Private Function ValidateSectionLists() As Boolean
    Dim values As Variant, v As Variant, n As Double, i As Long
    Dim seen(1 To 12) As Boolean
    ValidateSectionLists = False
    If Len(Trim$(SectionTypes)) > 0 Then
        values = Split(SectionTypes, ";")
        If UBound(values) + 1 <> NSections Then
            MsgBox "Укажите тип для каждой секции, разделяя номера точкой с запятой.", vbExclamation
            Exit Function
        End If
        For Each v In values
            If Not IsNumeric(Trim$(CStr(v))) Then GoTo badTypes
            n = CDbl(Trim$(CStr(v)))
            If n <> Int(n) Or n < 0 Or n > 4 Then GoTo badTypes
        Next v
    End If
    If Len(Trim$(EntranceSections)) > 0 Then
        values = Split(EntranceSections, ";")
        For Each v In values
            If Not IsNumeric(Trim$(CStr(v))) Then GoTo badEntrances
            n = CDbl(Trim$(CStr(v)))
            If n <> Int(n) Or n < 1 Or n > NSections Then GoTo badEntrances
            i = CLng(n)
            If seen(i) Then GoTo badEntrances
            seen(i) = True
        Next v
        NEntrances = UBound(values) + 1
    End If
    ValidateSectionLists = True
    Exit Function
badTypes:
    MsgBox "Тип секции должен быть целым числом от 0 до 4.", vbExclamation
    Exit Function
badEntrances:
    MsgBox "Укажите разные номера секций подъездов от 1 до числа секций.", vbExclamation
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
        ACOn = DetailsOn: PipesOn = DetailsOn: CracksOn = DetailsOn: BalconiesOn = DetailsOn
        CanopyOn = (NumL(parts(2), 1) = 1)
        Seed = NumL(parts(3), Seed)
    End If
End Sub

Private Function NumStr(v As Double) As String
    NumStr = Replace(CStr(v), ",", ".")
End Function
