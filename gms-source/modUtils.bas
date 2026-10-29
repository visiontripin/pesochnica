Attribute VB_Name = "modUtils"
Option Explicit

' ============================================================
' modUtils - базовые утилиты генератора фасада
' RNG, "живые" линии, клиппинг текстур, слои и цвета
' ============================================================

Private mSeed As Long
Private mRndState As Long

' 8 классов объектов (имена слоёв и RGB)
Public Const LAYER_COUNT As Long = 8
Public LayerNames(1 To 8) As String
Public LayerR(1 To 8) As Long
Public LayerG(1 To 8) As Long
Public LayerB(1 To 8) As Long

' индексы классов
Public Const LC_KARKAS As Long = 1
Public Const LC_PROYOM As Long = 2
Public Const LC_BALKON As Long = 3
Public Const LC_KLADKA As Long = 4
Public Const LC_KONDICIONER As Long = 5
Public Const LC_VODOSTOK As Long = 6
Public Const LC_TSEKOL As Long = 7
Public Const LC_KRYSHA As Long = 8

' рабочий документ и слой
Public gDoc As Document
Public gLayers(1 To 8) As Layer

' препятствия для текстур (прямоугольники)
Private mObsX1() As Double, mObsY1() As Double
Private mObsX2() As Double, mObsY2() As Double
Private mObsCount As Long

Public Sub InitLayersData()
    LayerNames(1) = "01_KARKAS":       LayerR(1) = 214: LayerG(1) = 39:  LayerB(1) = 40
    LayerNames(2) = "02_PROYOMY":      LayerR(2) = 31:  LayerG(2) = 95:  LayerB(2) = 214
    LayerNames(3) = "03_BALKONY":      LayerR(3) = 44:  LayerG(3) = 160: LayerB(3) = 44
    LayerNames(4) = "04_FAKTURA":      LayerR(4) = 140: LayerG(4) = 90:  LayerB(4) = 43
    LayerNames(5) = "05_KONDICIONERY": LayerR(5) = 197: LayerG(5) = 18:  LayerB(5) = 197
    LayerNames(6) = "06_VODOSTOK":     LayerR(6) = 15:  LayerG(6) = 154: LayerB(6) = 160
    LayerNames(7) = "07_TSEKOL":       LayerR(7) = 230: LayerG(7) = 138: LayerB(7) = 0
    LayerNames(8) = "08_KRYSHA":       LayerR(8) = 68:  LayerG(8) = 68:  LayerB(8) = 68
End Sub

Public Sub SeedRNG(seed As Long)
    Dim v As Double
    mSeed = seed
    v = Abs(CDbl(seed))
    v = v - Int(v / 2147483646#) * 2147483646#
    mRndState = CLng(v) + 1
End Sub

' Park-Miller / Schrage: без переполнения VBA Long при каждом шаге.
Public Function RndD() As Double
    Dim hi As Long, lo As Long, test As Long
    hi = mRndState \ 127773
    lo = mRndState Mod 127773
    test = 16807 * lo - 2836 * hi
    If test <= 0 Then test = test + 2147483647
    mRndState = test
    RndD = CDbl(mRndState) / 2147483647#
End Function

Public Function RndRange(lo As Double, hi As Double) As Double
    RndRange = lo + (hi - lo) * RndD()
End Function

Public Function MinD(a As Double, b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Public Function MaxD(a As Double, b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

' разбор числа с запятой/точкой (как ЗаменаТчкНаЗапт в шипоящике)
Public Function Num(ByVal txt As String, ByVal defaultValue As Double) As Double
    Dim s As String
    s = Trim$(txt)
    s = Replace(s, ".", ",")
    If IsNumeric(s) Then
        Num = CDbl(s)
    Else
        Num = defaultValue
    End If
End Function

Public Function NumL(ByVal txt As String, ByVal defaultValue As Long) As Long
    Dim v As Double
    v = Num(txt, CDbl(defaultValue))
    NumL = CLng(v)
End Function

Public Sub ResetObstacles()
    mObsCount = 0
    ReDim mObsX1(0 To 255)
    ReDim mObsY1(0 To 255)
    ReDim mObsX2(0 To 255)
    ReDim mObsY2(0 To 255)
End Sub

Public Sub AddObstacle(x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    If mObsCount >= UBound(mObsX1) Then
        ReDim Preserve mObsX1(0 To UBound(mObsX1) + 128)
        ReDim Preserve mObsY1(0 To UBound(mObsY1) + 128)
        ReDim Preserve mObsX2(0 To UBound(mObsX2) + 128)
        ReDim Preserve mObsY2(0 To UBound(mObsY2) + 128)
    End If
    mObsX1(mObsCount) = x1: mObsY1(mObsCount) = y1
    mObsX2(mObsCount) = x2: mObsY2(mObsCount) = y2
    mObsCount = mObsCount + 1
End Sub

' Прямоугольник по левому нижнему углу и размерам (в мм).
Public Sub AddObstacleRect(x As Double, y As Double, w As Double, h As Double)
    If w <= 0 Or h <= 0 Then Exit Sub
    AddObstacle x, y, x + w, y + h
End Sub

' вычитание интервалов-препятствий из отрезка [lo,hi] по одной оси
Private Function SubtractIntervals(lo As Double, hi As Double, blocks() As Double, n As Long) As Variant
    Dim res() As Double
    ReDim res(0 To 63)
    Dim cnt As Long: cnt = 0
    Dim curLo As Double, curHi As Double
    curLo = lo: curHi = hi
    ' Препятствия могут поступать в произвольном порядке: сортируем по началу.
    Dim i As Long, j As Long, a As Double, b As Double
    Dim tmpLo As Double, tmpHi As Double
    For i = 1 To n - 1
        tmpLo = blocks(i, 0): tmpHi = blocks(i, 1)
        j = i - 1
        Do While j >= 0
            If blocks(j, 0) <= tmpLo Then Exit Do
            blocks(j + 1, 0) = blocks(j, 0)
            blocks(j + 1, 1) = blocks(j, 1)
            j = j - 1
        Loop
        blocks(j + 1, 0) = tmpLo: blocks(j + 1, 1) = tmpHi
    Next i
    For i = 0 To n - 1
        a = blocks(i, 0): b = blocks(i, 1)
        If b > curLo And a < curHi Then
            If a <= curLo Then
                If b >= curHi Then
                    ' Остаток перекрыт, но ранее найденные интервалы сохраняем.
                    curLo = curHi
                    Exit For
                Else
                    curLo = b
                End If
            Else
                ' сохранить левый кусок
                If cnt + 1 >= UBound(res) Then ReDim Preserve res(0 To UBound(res) + 32)
                res(cnt) = curLo: res(cnt + 1) = a: cnt = cnt + 2
                curLo = MaxD(curLo, b)
                If curLo >= curHi Then Exit For
            End If
        End If
    Next i
    If curLo < curHi Then
        If cnt + 1 >= UBound(res) Then ReDim Preserve res(0 To UBound(res) + 32)
        res(cnt) = curLo: res(cnt + 1) = curHi: cnt = cnt + 2
    End If
    If cnt = 0 Then
        SubtractIntervals = Empty
    Else
        ReDim Preserve res(0 To cnt - 1)
        SubtractIntervals = res
    End If
End Function

' горизонтальный отрезок с клиппингом по препятствиям
Public Sub DrawHClip(y As Double, xa As Double, xb As Double, layerIdx As Long, _
                     margin As Double, amp As Double, jitterOn As Boolean)
    If mObsCount = 0 Then
        If xb - xa >= 0.7 Then DrawWobbleLine xa, y, xb, y, layerIdx, amp, jitterOn
        Exit Sub
    End If
    Dim blocks() As Double: ReDim blocks(0 To mObsCount * 2 - 1, 0 To 1)
    Dim n As Long: n = 0
    Dim i As Long
    For i = 0 To mObsCount - 1
        If mObsY1(i) - margin < y And y < mObsY2(i) + margin Then
            blocks(n, 0) = MaxD(xa, mObsX1(i) - margin)
            blocks(n, 1) = MinD(xb, mObsX2(i) + margin)
            n = n + 1
        End If
    Next i
    Dim parts As Variant
    parts = SubtractIntervals(xa, xb, blocks, n)
    If IsEmpty(parts) Then Exit Sub
    Dim k As Long
    For k = LBound(parts) To UBound(parts) Step 2
        If parts(k + 1) - parts(k) >= 0.7 Then
            DrawWobbleLine CDbl(parts(k)), y, CDbl(parts(k + 1)), y, layerIdx, amp, jitterOn
        End If
    Next k
End Sub

' вертикальный отрезок с клиппингом
Public Sub DrawVClip(x As Double, ya As Double, yb As Double, layerIdx As Long, _
                     margin As Double, amp As Double, jitterOn As Boolean)
    If mObsCount = 0 Then
        If yb - ya >= 0.7 Then DrawWobbleLine x, ya, x, yb, layerIdx, amp, jitterOn
        Exit Sub
    End If
    Dim blocks() As Double: ReDim blocks(0 To mObsCount * 2 - 1, 0 To 1)
    Dim n As Long: n = 0
    Dim i As Long
    For i = 0 To mObsCount - 1
        If mObsX1(i) - margin < x And x < mObsX2(i) + margin Then
            blocks(n, 0) = MaxD(ya, mObsY1(i) - margin)
            blocks(n, 1) = MinD(yb, mObsY2(i) + margin)
            n = n + 1
        End If
    Next i
    Dim parts As Variant
    parts = SubtractIntervals(ya, yb, blocks, n)
    If IsEmpty(parts) Then Exit Sub
    Dim k As Long
    For k = LBound(parts) To UBound(parts) Step 2
        If parts(k + 1) - parts(k) >= 0.7 Then
            DrawWobbleLine x, CDbl(parts(k)), x, CDbl(parts(k + 1)), layerIdx, amp, jitterOn
        End If
    Next k
End Sub

Public Function PointFree(x As Double, y As Double, margin As Double) As Boolean
    Dim i As Long
    For i = 0 To mObsCount - 1
        If mObsX1(i) - margin < x And x < mObsX2(i) + margin Then
            If mObsY1(i) - margin < y And y < mObsY2(i) + margin Then
                PointFree = False
                Exit Function
            End If
        End If
    Next i
    PointFree = True
End Function

' "живая" ломаная: синусоидальная волна + случайная компонента, концы фиксированы
Public Sub DrawWobbleLine(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                          layerIdx As Long, amp As Double, jitterOn As Boolean)
    Dim ln As Double
    ln = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    If ln < 0.001 Then Exit Sub

    Dim crv As Curve, sp As SubPath
    Set crv = Application.CreateCurve(gDoc)
    Set sp = crv.CreateSubPath(x1, y1)

    If Not jitterOn Or amp <= 0.0001 Then
        sp.AppendLineSegment x2, y2
    Else
        Dim n As Long: n = Int(ln / 3#)
        If n < 1 Then n = 1
        Dim ux As Double, uy As Double, nx As Double, ny As Double
        ux = (x2 - x1) / ln: uy = (y2 - y1) / ln
        nx = -uy: ny = ux
        Dim phase As Double: phase = RndRange(0, 6.2831853)
        Dim i As Long, t As Double, bx As Double, by As Double, off As Double
        For i = 1 To n - 1
            t = i / n
            bx = x1 + ux * ln * t
            by = y1 + uy * ln * t
            off = amp * (0.65 * Sin(t * 3.1415926 * 3# + phase) + 0.35 * RndRange(-1, 1))
            sp.AppendLineSegment bx + nx * off, by + ny * off
        Next i
        sp.AppendLineSegment x2, y2
    End If

    Dim sh As Shape
    Set sh = gLayers(layerIdx).CreateCurve(crv)
    ApplyStyle sh, layerIdx
End Sub

Public Sub DrawRectWobble(x As Double, y As Double, w As Double, h As Double, _
                          layerIdx As Long, amp As Double, jitterOn As Boolean)
    DrawWobbleLine x, y, x + w, y, layerIdx, amp, jitterOn
    DrawWobbleLine x + w, y, x + w, y + h, layerIdx, amp, jitterOn
    DrawWobbleLine x + w, y + h, x, y + h, layerIdx, amp, jitterOn
    DrawWobbleLine x, y + h, x, y, layerIdx, amp, jitterOn
End Sub

Public Sub DrawPoly(pts() As Double, layerIdx As Long, closed As Boolean)
    Dim crv As Curve, sp As SubPath
    Set crv = Application.CreateCurve(gDoc)
    Set sp = crv.CreateSubPath(pts(0, 0), pts(0, 1))
    Dim i As Long
    For i = 1 To UBound(pts, 1)
        sp.AppendLineSegment pts(i, 0), pts(i, 1)
    Next i
    If closed Then sp.AppendLineSegment pts(0, 0), pts(0, 1)
    Dim sh As Shape
    Set sh = gLayers(layerIdx).CreateCurve(crv)
    ApplyStyle sh, layerIdx
End Sub

Public Sub DrawCircleWobble(cx As Double, cy As Double, r As Double, _
                            layerIdx As Long, amp As Double, jitterOn As Boolean)
    Dim pts() As Double
    ReDim pts(0 To 27, 0 To 1)
    Dim k As Long, a As Double, rr As Double
    Dim ph1 As Double, ph2 As Double
    ph1 = RndRange(0, 6.2831853): ph2 = RndRange(0, 6.2831853)
    For k = 0 To 27
        a = 6.2831853 * k / 28
        rr = r
        If jitterOn Then rr = rr + amp * Sin(3 * a + ph1) + amp * 0.5 * Sin(5 * a + ph2)
        pts(k, 0) = cx + rr * Cos(a)
        pts(k, 1) = cy + rr * Sin(a)
    Next k
    DrawPoly pts, layerIdx, True
End Sub

Private Sub ApplyStyle(sh As Shape, layerIdx As Long)
    sh.Fill.ApplyNoFill
    sh.Outline.Width = 0.07
    sh.Outline.Color.RGBAssign LayerR(layerIdx), LayerG(layerIdx), LayerB(layerIdx)
End Sub

' создать 8 именованных слоёв с цветами
Public Sub PrepareLayers(doc As Document)
    Set gDoc = doc
    InitLayersData
    Dim i As Long, lyr As Layer
    For i = 1 To LAYER_COUNT
        Set lyr = Nothing
        On Error Resume Next
        Set lyr = doc.ActivePage.Layers(LayerNames(i))
        On Error GoTo 0
        If lyr Is Nothing Then
            Set lyr = doc.ActivePage.CreateLayer(LayerNames(i))
        End If
        lyr.Color.RGBAssign LayerR(i), LayerG(i), LayerB(i)
        Set gLayers(i) = lyr
    Next i
End Sub
