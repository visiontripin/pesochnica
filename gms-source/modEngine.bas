Attribute VB_Name = "modEngine"
Option Explicit

' ============================================================
' modEngine - параметрический генератор фасада панельного дома
' Все настройки берутся из Public-переменных модуля modMain.
' ============================================================

Private mX0 As Double, mX1 As Double
Private mFloorH As Double
Private mFoundH As Double, mRoofH As Double, mRoofBot As Double
Private mNSec As Long
Private mSecType() As Long
Private mSecX1() As Double, mSecX2() As Double
Private mNF As Long
Private mFY() As Double          ' низ каждого этажа
Private mFK() As Long            ' 0 = цокольный, 1 = жилой
Private mJoint As Double
Private mACmap() As String       ' ключ "sec|floor" -> позиция L/R
Private mShapeCount As Long

Public Sub BuildFacade()
    Dim doc As Document
    Set doc = Application.CreateDocument
    doc.Unit = cdrMillimeter
    doc.ActivePage.SetSize PanelW, PanelH
    doc.ClearSelection
    PrepareLayers doc
    ResetObstacles
    SeedRNG Seed
    mShapeCount = 0
    mJoint = 2#

    ' защита от слишком мелких шагов фактуры под лазерный луч
    If TexA < GapMin + 0.2 Then TexA = GapMin + 0.2
    If TexB < GapMin + 0.2 Then TexB = GapMin + 0.2

    ComputeZones
    ComputeSections
    RegisterDownpipeObstacles
    RegisterAllOpenings
    DrawFrameAndJoints
    DrawFoundation
    DrawFloors
    DrawTextures
    If DetailsOn Then
        DrawDownpipes
        ' Кондиционеры уже рисуются в DrawFloors по карте mACmap.
        DrawCracks
    End If
    DrawRoof

    doc.ClearSelection
    MsgBox "Фасад готов: " & mShapeCount & " объектов, " & NSectionsAll & " секций, " & _
           (NFloors + NPlinth) & " этажей." & vbCrLf & _
           "Слои 01..08 соответствуют классам объектов.", vbInformation, "Генератор фасада"
End Sub

Private Function NSectionsAll() As Long
    NSectionsAll = mNSec
End Function

' ------------------------------------------------------------
' зоны: фундамент, цокольные этажи, жилые этажи, крыша
' ------------------------------------------------------------
Private Sub ComputeZones()
    mX0 = MarginMm: mX1 = PanelW - MarginMm

    Select Case RoofType
        Case 3: mRoofH = 0                                  ' нет крыши
        Case 0: mRoofH = PanelH * 0.115                     ' плоская
        Case 1: mRoofH = PanelH * 0.1                      ' двускатная
        Case Else: mRoofH = PanelH * 0.08                  ' односкатная
    End Select
    If FoundationOn Then
        mFoundH = FoundationH
        If mFoundH <= 0 Then mFoundH = PanelH * 0.05
    Else
        mFoundH = 0
    End If

    Dim avail As Double
    avail = PanelH - 2 * MarginMm - mRoofH - mFoundH
    Dim denom As Double
    denom = CDbl(NPlinth) * 0.9 + CDbl(NFloors)
    mFloorH = avail / denom

    ReDim mFY(0 To NFloors + NPlinth - 1)
    ReDim mFK(0 To NFloors + NPlinth - 1)
    Dim i As Long, y As Double
    y = mFoundH
    For i = 0 To NFloors + NPlinth - 1
        mFY(i) = y
        If i < NPlinth Then
            mFK(i) = 0
            y = y + mFloorH * 0.9
        Else
            mFK(i) = 1
            y = y + mFloorH
        End If
    Next i
    mNF = NFloors + NPlinth
    mRoofBot = y
End Sub

Private Sub ComputeSections()
    mNSec = NSections
    ReDim mSecType(1 To mNSec)
    ReDim mSecX1(1 To mNSec)
    ReDim mSecX2(1 To mNSec)
    Dim base(0 To 4) As Long
    base(0) = 0: base(1) = 1: base(2) = 2: base(3) = 3: base(4) = 4
    Dim totalW As Double: totalW = mX1 - mX0
    Dim jointsW As Double: jointsW = (mNSec - 1) * mJoint
    Dim secW As Double: secW = (totalW - jointsW) / mNSec
    Dim i As Long, x As Double
    x = mX0
    For i = 1 To mNSec
        mSecType(i) = base((i - 1) Mod 5)
        mSecX1(i) = x
        If i = mNSec Then
            mSecX2(i) = mX1
        Else
            mSecX2(i) = x + secW
        End If
        x = x + secW + mJoint
    Next i
End Sub

' ------------------------------------------------------------
' каркас
' ------------------------------------------------------------
Private Sub DrawFrameAndJoints()
    ' физическая кромка панели - строго прямая
    DrawWobbleLine 0, 0, PanelW, 0, LC_KARKAS, 0, False
    DrawWobbleLine PanelW, 0, PanelW, PanelH, LC_KARKAS, 0, False
    DrawWobbleLine PanelW, PanelH, 0, PanelH, LC_KARKAS, 0, False
    DrawWobbleLine 0, PanelH, 0, 0, LC_KARKAS, 0, False
    mShapeCount = mShapeCount + 4

    Dim i As Long
    For i = 0 To mNF - 1
        DrawWobbleLine mX0, mFY(i), mX1, mFY(i), LC_KARKAS, 0.13, JitterOn
        mShapeCount = mShapeCount + 1
    Next i
    ' верхняя граница жилой зоны
    DrawWobbleLine mX0, mRoofBot, mX1, mRoofBot, LC_KARKAS, 0.13, JitterOn
    mShapeCount = mShapeCount + 1

    For i = 1 To mNSec - 1
        Dim jx As Double
        jx = (mSecX2(i) + mSecX1(i + 1)) / 2
        DrawWobbleLine jx, MarginMm, jx, mRoofBot, LC_KARKAS, 0.13, JitterOn
        mShapeCount = mShapeCount + 1
    Next i
End Sub

' ------------------------------------------------------------
' фундамент и цоколь
' ------------------------------------------------------------
Private Sub DrawFoundation()
    If FoundationOn Then
        Dim yy As Double
        yy = mFoundH * 0.4
        DrawWobbleLine mX0 + 0.8, yy, mX1 - 0.8, yy, LC_TSEKOL, 0.1, JitterOn
        yy = mFoundH * 0.75
        DrawWobbleLine mX0 + 0.8, yy, mX1 - 0.8, yy, LC_TSEKOL, 0.1, JitterOn
        DrawVentGrille mX0 + (mX1 - mX0) * 0.18, mFoundH * 0.32
        DrawVentGrille mX0 + (mX1 - mX0) * 0.68, mFoundH * 0.32
        mShapeCount = mShapeCount + 2
    End If

    ' русты и решётки цокольных этажей
    Dim i As Long
    For i = 0 To mNF - 1
        If mFK(i) = 0 Then
            DrawWobbleLine mX0 + 0.8, mFY(i) + mFloorH * 0.18, mX1 - 0.8, _
                           mFY(i) + mFloorH * 0.18, LC_TSEKOL, 0.09, JitterOn
            DrawVentGrille mX0 + (mX1 - mX0) * 0.3, mFY(i) + mFloorH * 0.3
            mShapeCount = mShapeCount + 1
        End If
    Next i
End Sub

Private Sub DrawVentGrille(x As Double, y As Double)
    Dim w As Double, h As Double
    w = MaxD(6, PanelW * 0.05): h = MaxD(2.6, mFloorH * 0.12)
    DrawRectWobble x, y, w, h, LC_TSEKOL, 0.06, JitterOn
    Dim k As Long, gx As Double
    For k = 1 To 4
        gx = x + w * k / 5
        DrawWobbleLine gx, y + 0.3, gx, y + h - 0.3, LC_TSEKOL, 0.04, JitterOn
    Next k
    AddObstacle x - 0.3, y - 0.3, x + w + 0.3, y + h + 0.3
    mShapeCount = mShapeCount + 5
End Sub

' ------------------------------------------------------------
' регистрация проёмов (до фактуры)
' ------------------------------------------------------------
Private Sub RegisterDownpipeObstacles()
    Dim joints(0 To 1) As Long
    joints(0) = MinL(2, mNSec - 1)
    joints(1) = MinL(4, mNSec - 1)
    Dim k As Long, jx As Double
    For k = 0 To 1
        If joints(k) >= 1 And joints(k) < mNSec Then
            jx = (mSecX2(joints(k)) + mSecX1(joints(k) + 1)) / 2 + 1.3
            AddObstacle jx - 1.15, mFoundH - 0.5, jx + 1.15, mRoofBot + 1.5
        End If
    Next k
End Sub

Private Sub RegisterAllOpenings()
    ReDim mACmap(0 To 1023)
    Dim i As Long, f As Long
    For i = 1 To mNSec
        For f = 0 To mNF - 1
            RegisterSectionOpening i, f
        Next f
    Next i
    ' кондиционеры - только на секциях парных окон
    Dim cands() As Long, cnt As Long
    ReDim cands(0 To mNSec * mNF)
    cnt = 0
    For i = 1 To mNSec
        If mSecType(i) = 2 Then
            For f = NPlinth To mNF - 1
                cands(cnt) = i * 1000 + f: cnt = cnt + 1
            Next f
        End If
    Next i
    If cnt > 0 Then
        Dim want As Long: want = MinL(6, cnt)
        Dim got As Long, pick As Long, key As Long, si As Long, fl As Long
        Do While got < want
            pick = Int(RndD() * cnt)
            key = cands(pick)
            If key > 0 Then
                cands(pick) = 0
                si = key \ 1000: fl = key Mod 1000
                mACmap(si * 32 + fl) = IIf(RndD() < 0.5, "L", "R")
                RegisterACObstacle si, fl
                got = got + 1
            End If
        Loop
    End If
End Sub

Private Sub RegisterSectionOpening(si As Long, f As Long)
    Dim x0 As Double, x1 As Double, yb As Double, cx As Double
    x0 = mSecX1(si): x1 = mSecX2(si): yb = mFY(f): cx = (x0 + x1) / 2
    Dim sw As Double, wh As Double, wy As Double
    sw = (x1 - x0) * 0.36
    wh = mFloorH * 0.53
    wy = yb + mFloorH * 0.3

    Select Case mSecType(si)
        Case 0 ' одинарные окна
            AddObstacle cx - sw / 2 - 1.4, wy - 2.2, sw + 2.8, wh + 2.6
        Case 2 ' парные окна
            Dim gapW As Double: gapW = (x1 - x0) * 0.06
            AddObstacle x0 + (x1 - x0) * 0.08 - 1.4, wy - 2.2, sw + 2.8, wh + 2.6
            AddObstacle x1 - (x1 - x0) * 0.08 - sw - 1.4, wy - 2.2, sw + 2.8, wh + 2.6
        Case 1 ' балкон
            AddObstacle cx - (x1 - x0) * 0.42, yb + mFloorH * 0.16, (x1 - x0) * 0.84, mFloorH * 0.7
            AddObstacle cx - (x1 - x0) * 0.47, yb + 0.4, (x1 - x0) * 0.94, mFloorH * 0.15
        Case 3 ' сплошное остекление
            AddObstacle x0 + 1.4, yb + mFloorH * 0.15, (x1 - x0) - 2.8, mFloorH * 0.68
        Case 4 ' лоджия / подъезд
            AddObstacle x0 + 0.3, yb + 0.3, (x1 - x0) - 0.6, mFloorH - 0.6
    End Select
End Sub

Private Sub RegisterACObstacle(si As Long, fl As Long)
    Dim x0 As Double: x0 = mSecX1(si)
    Dim yb As Double: yb = mFY(fl)
    Dim acx As Double, aw As Double
    aw = MaxD(4.2, (mSecX2(si) - x0) * 0.14)
    If mACmap(si * 32 + fl) = "L" Then
        acx = x0 + (mSecX2(si) - x0) * 0.24
    Else
        acx = x0 + (mSecX2(si) - x0) * 0.76
    End If
    AddObstacle acx - aw / 2 - 0.4, yb + 0.9, aw + 0.8, MaxD(3.2, mFloorH * 0.12) + 0.8
End Sub

' ------------------------------------------------------------
' поэтажная отрисовка
' ------------------------------------------------------------
Private Sub DrawFloors()
    Dim f As Long, si As Long
    For f = 0 To mNF - 1
        For si = 1 To mNSec
            DrawSection si, f
        Next si
    Next f
End Sub

Private Sub DrawSection(si As Long, f As Long)
    Dim x0 As Double, x1 As Double, yb As Double, cx As Double
    x0 = mSecX1(si): x1 = mSecX2(si): yb = mFY(f): cx = (x0 + x1) / 2
    Dim sw As Double, wh As Double, wy As Double
    sw = (x1 - x0) * 0.36
    wh = mFloorH * 0.53
    wy = yb + mFloorH * 0.3
    Dim st As Long: st = WindowStyleFor(si, f)

    Select Case mSecType(si)
        Case 0
            If WindowType = 4 Then
                DrawBalconyDoor cx - sw / 2, wy, sw, wh
            Else
                DrawWindow cx - sw / 2, wy, sw, wh, st
            End If
        Case 2
            DrawWindow x0 + (x1 - x0) * 0.08, wy, sw, wh, st
            DrawWindow x1 - (x1 - x0) * 0.08 - sw, wy, sw, wh, (st + 1) Mod 4
            If DetailsOn And Len(mACmap(si * 32 + f)) > 0 Then DrawAC si, f
        Case 1
            DrawBalcony cx, yb, x1 - x0
        Case 3
            DrawGlazed x0, x1, yb
        Case 4
            If f = NPlinth And NEntrances > 0 Then
                DrawEntrance x0, x1, yb
            Else
                DrawLoggia x0, x1, yb
            End If
    End Select
End Sub

Private Function WindowStyleFor(si As Long, f As Long) As Long
    If RandomStyles Then
        WindowStyleFor = (f * 7 + si * 3 + Seed) Mod 4
    Else
        WindowStyleFor = WindowType
        ' форма 4 (балканный блок) применима только к одинарным секциям
        If WindowStyleFor > 3 Then WindowStyleFor = 3
    End If
End Function

' ------------------------------------------------------------
' окно: 4 стандартные формы
' 0 одностворчатое, 1 двустворчатое, 2 трёхстворчатое, 3 с форточкой
' ------------------------------------------------------------
Private Sub DrawWindow(x As Double, y As Double, w As Double, h As Double, style As Long)
    Dim fm As Double: fm = MaxD(0.8, w * 0.07)
    DrawRectWobble x, y, w, h, LC_PROYOM, 0.1, JitterOn
    DrawRectWobble x + fm, y + fm, w - 2 * fm, h - 2 * fm, LC_PROYOM, 0.07, JitterOn
    Dim ix As Double, iy As Double, iw As Double, ih As Double
    ix = x + fm: iy = y + fm: iw = w - 2 * fm: ih = h - 2 * fm
    mShapeCount = mShapeCount + 2

    Dim m As Double, m1 As Double, m2 As Double
    If style = 2 Then
        m1 = ix + iw / 3: m2 = ix + 2 * iw / 3
        DrawWobbleLine m1, iy, m1, iy + ih, LC_PROYOM, 0.06, JitterOn
        DrawWobbleLine m2, iy, m2, iy + ih, LC_PROYOM, 0.06, JitterOn
        DrawSashFrame ix, m1, iy, ih
        DrawSashFrame m1, m2, iy, ih
        DrawSashFrame m2, ix + iw, iy, ih
        mShapeCount = mShapeCount + 5
    Else
        If style = 3 Then m = ix + iw * 0.42 Else m = ix + iw / 2
        DrawWobbleLine m, iy, m, iy + ih, LC_PROYOM, 0.06, JitterOn
        DrawSashFrame ix, m, iy, ih
        DrawSashFrame m, ix + iw, iy, ih
        mShapeCount = mShapeCount + 3
        If style = 3 Then
            Dim fy As Double: fy = iy + ih * 0.66
            DrawWobbleLine m + 0.5, fy, ix + iw - 0.5, fy, LC_PROYOM, 0.05, JitterOn
            DrawWobbleLine (m + ix + iw) / 2, fy + 0.4, (m + ix + iw) / 2, iy + ih - 0.5, _
                           LC_PROYOM, 0.04, JitterOn
            ' ручка форточки
            DrawWobbleLine ix + iw - 0.9, fy - 1#, ix + iw - 0.9, fy - 0.35, _
                           LC_PROYOM, 0.03, JitterOn
            mShapeCount = mShapeCount + 3
        End If
    End If

    ' ручки створок
    Dim hx As Double
    If style = 2 Then
        DrawWobbleLine m1 - 0.8, y + h * 0.42, m1 - 0.8, y + h * 0.56, LC_PROYOM, 0.03, JitterOn
        DrawWobbleLine m2 - 0.8, y + h * 0.42, m2 - 0.8, y + h * 0.56, LC_PROYOM, 0.03, JitterOn
        mShapeCount = mShapeCount + 2
    Else
        hx = IIf(style = 3, m - 0.8, m - 0.8)
        DrawWobbleLine hx, y + h * 0.42, hx, y + h * 0.56, LC_PROYOM, 0.03, JitterOn
        DrawWobbleLine m + 0.8, y + h * 0.42, m + 0.8, y + h * 0.56, LC_PROYOM, 0.03, JitterOn
        mShapeCount = mShapeCount + 2
    End If

    ' отлив с капельником
    DrawRectWobble x - 1#, y - 1.1, w + 2#, 0.65, LC_PROYOM, 0.06, JitterOn
    DrawWobbleLine x - 0.5, y - 1.95, x + w + 0.5, y - 1.95, LC_PROYOM, 0.05, JitterOn
    mShapeCount = mShapeCount + 2
    AddObstacle x - 1.4, y - 2.3, w + 2.8, h + 2.7
End Sub

Private Sub DrawSashFrame(ax As Double, bx As Double, iy As Double, ih As Double)
    Dim inset As Double: inset = 0.5
    If bx - ax < 2.2 * inset + 0.4 Then Exit Sub
    DrawRectWobble ax + inset, iy + inset, (bx - ax) - 2 * inset, ih - 2 * inset, _
                   LC_PROYOM, 0.05, JitterOn
End Sub

' балконный дверной блок (форма 4): дверь + окно
Private Sub DrawBalconyDoor(x As Double, y As Double, w As Double, h As Double)
    Dim fm As Double: fm = MaxD(0.8, w * 0.07)
    DrawRectWobble x, y, w, h, LC_PROYOM, 0.1, JitterOn
    DrawRectWobble x + fm, y + fm, w - 2 * fm, h - 2 * fm, LC_PROYOM, 0.07, JitterOn
    Dim split As Double: split = x + w * 0.52
    DrawWobbleLine split, y + fm, split, y + h - fm, LC_PROYOM, 0.06, JitterOn
    Dim ty As Double: ty = y + h * 0.78
    DrawWobbleLine x + fm, ty, split - 0.4, ty, LC_PROYOM, 0.05, JitterOn
    DrawWobbleLine split + 0.4, ty, x + w - fm, ty, LC_PROYOM, 0.05, JitterOn
    DrawWobbleLine split - 0.8, y + h * 0.42, split - 0.8, y + h * 0.56, LC_PROYOM, 0.03, JitterOn
    DrawRectWobble x - 1#, y - 1.1, w + 2#, 0.65, LC_PROYOM, 0.06, JitterOn
    mShapeCount = mShapeCount + 7
    AddObstacle x - 1.4, y - 2.3, w + 2.8, h + 2.7
End Sub

' ------------------------------------------------------------
' балкон, сплошное остекление, лоджия, подъезд
' ------------------------------------------------------------
Private Sub DrawBalcony(cx As Double, yb As Double, secW As Double)
    Dim gw As Double, gh As Double
    gw = secW * 0.78: gh = mFloorH * 0.62
    Dim gx As Double, gy As Double
    gx = cx - gw / 2: gy = yb + mFloorH * 0.2
    DrawRectWobble gx, gy, gw, gh, LC_BALKON, 0.1, JitterOn
    DrawRectWobble gx + 0.8, gy + 0.8, gw - 1.6, gh - 1.6, LC_BALKON, 0.07, JitterOn
    Dim ix0 As Double, iy0 As Double, iw As Double, ih As Double
    ix0 = gx + 0.8: iy0 = gy + 0.8: iw = gw - 1.6: ih = gh - 1.6
    Dim transH As Double: transH = gh * 0.24
    Dim ty As Double: ty = iy0 + ih - transH
    DrawWobbleLine ix0, ty, ix0 + iw, ty, LC_BALKON, 0.06, JitterOn
    Dim k As Double, tx As Double
    For k = 1 To 4
        tx = ix0 + iw * k / 5
        DrawWobbleLine tx, ty, tx, iy0 + ih, LC_BALKON, 0.05, JitterOn
    Next k
    For k = 1 To 2
        tx = ix0 + iw * k / 3
        DrawWobbleLine tx, iy0, tx, ty, LC_BALKON, 0.06, JitterOn
    Next k
    ' плита и парапет
    DrawWobbleLine gx - 1#, gy - 0.7, gx + gw + 1#, gy - 0.7, LC_BALKON, 0.07, JitterOn
    Dim bw As Double, bh As Double, bx As Double, byy As Double
    bw = secW * 0.9: bh = MaxD(2.8, mFloorH * 0.12)
    bx = cx - bw / 2: byy = yb + 0.8
    DrawRectWobble bx, byy, bw, bh, LC_BALKON, 0.09, JitterOn
    DrawWobbleLine bx + 0.6, byy + bh - 0.9, bx + bw - 0.6, byy + bh - 0.9, LC_BALKON, 0.05, JitterOn
    Dim rx As Double
    rx = bx + 1.2
    Do While rx < bx + bw - 0.8
        DrawVClip rx, byy + 0.35, byy + bh - 0.9, LC_BALKON, GapMin + 0.3, 0.04, JitterOn
        rx = rx + MaxD(1.5, bw * 0.07)
    Loop
    mShapeCount = mShapeCount + 12
End Sub

Private Sub DrawGlazed(x0 As Double, x1 As Double, yb As Double)
    Dim fx0 As Double, fx1 As Double, fy0 As Double, fy1 As Double
    fx0 = x0 + 1.6: fx1 = x1 - 1.6
    fy0 = yb + mFloorH * 0.18: fy1 = yb + mFloorH * 0.8
    DrawRectWobble fx0, fy0, fx1 - fx0, fy1 - fy0, LC_BALKON, 0.09, JitterOn
    DrawRectWobble fx0 + 0.8, fy0 + 0.8, fx1 - fx0 - 1.6, fy1 - fy0 - 1.6, LC_BALKON, 0.06, JitterOn
    Dim iw As Double: iw = fx1 - fx0 - 1.6
    Dim ty As Double: ty = fy1 - 0.8 - (fy1 - fy0) * 0.22
    DrawWobbleLine fx0 + 0.8, ty, fx1 - 0.8, ty, LC_BALKON, 0.05, JitterOn
    Dim k As Double, tx As Double
    For k = 1 To 3
        tx = fx0 + 0.8 + iw * k / 4
        DrawWobbleLine tx, ty, tx, fy1 - 0.8, LC_BALKON, 0.04, JitterOn
    Next k
    For k = 1 To 2
        tx = fx0 + 0.8 + iw * k / 3
        DrawWobbleLine tx, fy0 + 0.8, tx, ty, LC_BALKON, 0.05, JitterOn
    Next k
    ' бетонный фартук
    Dim aw As Double, ah As Double, ay As Double
    aw = (x1 - x0) - 3.2: ah = MaxD(2.4, mFloorH * 0.1): ay = yb + 1#
    DrawRectWobble x0 + 1.6, ay, aw, ah, LC_BALKON, 0.08, JitterOn
    DrawWobbleLine x0 + 2.1, ay + ah / 2, x1 - 2.1, ay + ah / 2, LC_BALKON, 0.05, JitterOn
    mShapeCount = mShapeCount + 10
End Sub

Private Sub DrawLoggia(x0 As Double, x1 As Double, yb As Double)
    ' боковая стенка двойной линией
    DrawRectWobble x1 - 2.2, yb + 0.6, 0.8, mFloorH - 1.2, LC_BALKON, 0.07, JitterOn
    Dim dw As Double, dh As Double, dx As Double, dy As Double
    dw = (x1 - x0) * 0.3: dh = mFloorH * 0.56
    dx = x0 + 2#: dy = yb + mFloorH * 0.2
    DrawRectWobble dx, dy, dw, dh, LC_BALKON, 0.09, JitterOn
    DrawRectWobble dx + 0.7, dy + 0.7, dw - 1.4, dh - 1.4, LC_BALKON, 0.06, JitterOn
    DrawWobbleLine dx + dw / 2, dy + 0.7, dx + dw / 2, dy + dh - 0.7, LC_BALKON, 0.05, JitterOn
    DrawWobbleLine dx + 0.7, dy + dh - dh * 0.22, dx + dw - 0.7, dy + dh - dh * 0.22, _
                   LC_BALKON, 0.05, JitterOn
    ' ограждение
    DrawWobbleLine x0 + 1.2, yb + mFloorH * 0.13, x1 - 2.4, yb + mFloorH * 0.13, LC_BALKON, 0.06, JitterOn
    DrawWobbleLine x0 + 1.2, yb + mFloorH * 0.07, x1 - 2.4, yb + mFloorH * 0.07, LC_BALKON, 0.06, JitterOn
    mShapeCount = mShapeCount + 7
End Sub

Private Sub DrawEntrance(x0 As Double, x1 As Double, yb As Double)
    Dim dw As Double, dh As Double, dx As Double, dy As Double
    dw = (x1 - x0) * 0.42: dh = mFloorH * 0.66
    dx = x0 + (x1 - x0) * 0.12: dy = yb + 0.8
    DrawRectWobble dx, dy, dw, dh, LC_PROYOM, 0.1, JitterOn
    DrawRectWobble dx + 0.8, dy + 0.8, dw - 1.6, dh - 1.6, LC_PROYOM, 0.07, JitterOn
    ' фрамуга и створки
    DrawWobbleLine dx + 0.8, dy + dh - 3.6, dx + dw - 0.8, dy + dh - 3.6, LC_PROYOM, 0.05, JitterOn
    DrawWobbleLine dx + dw / 2, dy + 0.8, dx + dw / 2, dy + dh - 3.6, LC_PROYOM, 0.05, JitterOn
    DrawWobbleLine dx + dw / 2 - 0.8, dy + dh * 0.42, dx + dw / 2 - 0.8, dy + dh * 0.56, LC_PROYOM, 0.03, JitterOn
    DrawWobbleLine dx + dw / 2 + 0.8, dy + dh * 0.42, dx + dw / 2 + 0.8, dy + dh * 0.56, LC_PROYOM, 0.03, JitterOn
    If CanopyOn Then
        DrawWobbleLine dx - 2#, dy + dh + 0.9, dx + dw + 2#, dy + dh + 0.9, LC_PROYOM, 0.07, JitterOn
        DrawWobbleLine dx - 1.4, dy + dh + 1.9, dx + dw + 1.4, dy + dh + 1.9, LC_PROYOM, 0.07, JitterOn
        DrawWobbleLine dx - 1.4, dy + dh + 0.9, dx - 1.4, dy + dh + 1.9, LC_PROYOM, 0.04, JitterOn
        DrawWobbleLine dx + dw + 1.4, dy + dh + 0.9, dx + dw + 1.4, dy + dh + 1.9, LC_PROYOM, 0.04, JitterOn
    End If
    ' ступени
    DrawWobbleLine dx - 1.8, dy - 0.5, dx + dw + 1.8, dy - 0.5, LC_PROYOM, 0.06, JitterOn
    DrawWobbleLine dx - 3#, dy - 1.05, dx + dw + 3#, dy - 1.05, LC_PROYOM, 0.06, JitterOn
    mShapeCount = mShapeCount + 11
End Sub

' ------------------------------------------------------------
' фактуры отделки: 0 кирпич, 1 доски, 2 облицовочные панели
' ------------------------------------------------------------
Private Sub DrawTextures()
    Dim f As Long, si As Long
    For f = 0 To mNF - 1
        For si = 1 To mNSec
            If mSecType(si) <> 4 Then DrawTextureSection si, f
        Next si
    Next f
End Sub

Private Sub DrawTextureSection(si As Long, f As Long)
    Dim x0 As Double, x1 As Double, yb As Double
    x0 = mSecX1(si) + 0.5: x1 = mSecX2(si) - 0.5: yb = mFY(f)
    Dim ya As Double, ybnd As Double
    ya = yb + 0.45: ybnd = yb + mFloorH - 0.45
    Dim amp As Double: amp = 0.045
    Dim margin As Double: margin = GapMin + 0.3

    If TextureType = 0 Then
        ' кирпич ложковой перевязкой
        Dim rowH As Double: rowH = TexB
        Dim brickL As Double: brickL = TexA
        Dim row As Long: row = 0
        Dim y As Double: y = CeilD(ya / rowH) * rowH
        Do While y <= ybnd
            DrawHClip y, x0, x1, LC_KLADKA, margin, amp, JitterOn
            Dim stagger As Double
            If row Mod 2 = 1 Then stagger = brickL / 2 Else stagger = 0
            Dim x As Double
            x = CeilD((x0 - stagger) / brickL) * brickL + stagger
            Do While x <= x1
                DrawVClip x, y + 0.1, MinD(y + rowH - 0.1, ybnd), LC_KLADKA, margin, amp * 0.9, JitterOn
                x = x + brickL
            Loop
            y = y + rowH: row = row + 1
        Loop
    ElseIf TextureType = 1 Then
        ' горизонтальные доски: длинные линии + разбежка стыков
        Dim boardW As Double: boardW = TexA
        Dim boardL As Double: boardL = TexB * 4
        y = CeilD(ya / boardW) * boardW
        row = 0
        Do While y <= ybnd
            DrawHClip y, x0, x1, LC_KLADKA, margin, amp, JitterOn
            stagger = IIf(row Mod 2 = 1, boardL / 2, 0)
            x = CeilD((x0 - stagger) / boardL) * boardL + stagger
            Do While x <= x1
                DrawVClip x, y + 0.1, MinD(y + boardW - 0.1, ybnd), LC_KLADKA, margin, amp * 0.9, JitterOn
                x = x + boardL
            Loop
            y = y + boardW: row = row + 1
        Loop
    Else
        ' облицовочные панели: сетка прямоугольников с зазором TexGap
        Dim pw As Double, ph As Double
        pw = TexA: ph = TexB
        Dim gy0 As Double: gy0 = CeilD(ya / ph) * ph
        y = gy0
        Do While y + ph <= ybnd + 0.01
            x = CeilD(x0 / pw) * pw
            Do While x + pw <= x1 + 0.01
                If PointFree(x + TexGap, y + TexGap, margin) And _
                   PointFree(x + pw - TexGap, y + ph - TexGap, margin) Then
                    DrawRectWobble x + TexGap / 2, y + TexGap / 2, _
                                   pw - TexGap, ph - TexGap, LC_KLADKA, 0.04, JitterOn
                    mShapeCount = mShapeCount + 1
                End If
                x = x + pw
            Loop
            y = y + ph
        Loop
    End If
End Sub

' ------------------------------------------------------------
' водосток, кондиционер, трещины
' ------------------------------------------------------------
Private Sub DrawDownpipes()
    Dim joints(0 To 1) As Long
    joints(0) = MinL(2, mNSec - 1)
    joints(1) = MinL(4, mNSec - 1)
    Dim k As Long, jx As Double, gap As Double
    gap = 0.85
    For k = 0 To 1
        If joints(k) < 1 Then GoTo nextpipe
        jx = (mSecX2(joints(k)) + mSecX1(joints(k) + 1)) / 2 + 1.3
        DrawWobbleLine jx - gap / 2, mFoundH + 0.5, jx - gap / 2, mRoofBot - 0.5, LC_VODOSTOK, 0.08, JitterOn
        DrawWobbleLine jx + gap / 2, mFoundH + 0.5, jx + gap / 2, mRoofBot - 0.5, LC_VODOSTOK, 0.08, JitterOn
        Dim y As Double
        y = mFoundH + 2.5
        Do While y < mRoofBot - 1
            DrawWobbleLine jx - 0.8, y, jx + 0.8, y, LC_VODOSTOK, 0.04, JitterOn
            y = y + mFloorH
        Loop
        ' воронка
        Dim pts1(0 To 1, 0 To 1) As Double, pts2(0 To 1, 0 To 1) As Double
        pts1(0, 0) = jx - 1.4: pts1(0, 1) = mRoofBot + 1.4
        pts1(1, 0) = jx - gap / 2: pts1(1, 1) = mRoofBot
        pts2(0, 0) = jx + 1.4: pts2(0, 1) = mRoofBot + 1.4
        pts2(1, 0) = jx + gap / 2: pts2(1, 1) = mRoofBot
        DrawPoly pts1, LC_VODOSTOK, False
        DrawPoly pts2, LC_VODOSTOK, False
        DrawWobbleLine jx - 1.4, mRoofBot + 1.4, jx + 1.4, mRoofBot + 1.4, LC_VODOSTOK, 0.05, JitterOn
        ' выпуск
        DrawWobbleLine jx - 0.65, mFoundH + 0.5, jx - 0.65, mFoundH - 0.7, LC_VODOSTOK, 0.04, JitterOn
        DrawWobbleLine jx + 0.65, mFoundH + 0.5, jx + 0.65, mFoundH - 0.7, LC_VODOSTOK, 0.04, JitterOn
        mShapeCount = mShapeCount + 8
nextpipe:
    Next k
End Sub

Private Sub DrawAC(si As Long, fl As Long)
    Dim x0 As Double: x0 = mSecX1(si)
    Dim secW As Double: secW = mSecX2(si) - x0
    Dim yb As Double: yb = mFY(fl)
    Dim aw As Double, ah As Double
    aw = MaxD(4.2, secW * 0.14): ah = MaxD(2.6, mFloorH * 0.1)
    Dim acx As Double
    If mACmap(si * 32 + fl) = "L" Then acx = x0 + secW * 0.24 Else acx = x0 + secW * 0.76
    Dim topY As Double: topY = yb + mFloorH * 0.16
    Dim x As Double, y As Double
    x = acx - aw / 2: y = topY - ah
    DrawRectWobble x, y, aw, ah, LC_KONDICIONER, 0.07, JitterOn
    Dim fcx As Double, fcy As Double
    fcx = x + aw * 0.3: fcy = y + ah / 2
    DrawCircleWobble fcx, fcy, MinD(0.75, ah * 0.32), LC_KONDICIONER, 0.04, JitterOn
    DrawCircleWobble fcx, fcy, 0.28, LC_KONDICIONER, 0.03, JitterOn
    Dim k As Long, a As Double
    For k = 0 To 2
        a = k * 2.0944 + 0.4
        DrawWobbleLine fcx + 0.3 * Cos(a), fcy + 0.3 * Sin(a), _
                       fcx + 0.55 * Cos(a), fcy + 0.55 * Sin(a), LC_KONDICIONER, 0.02, JitterOn
    Next k
    Dim gx As Double
    For k = 0 To 2
        gx = x + aw - aw * 0.42 + k * aw * 0.13
        DrawWobbleLine gx, y + 0.45, gx, y + ah - 0.45, LC_KONDICIONER, 0.03, JitterOn
    Next k
    ' кронштейны
    DrawWobbleLine x + 0.6, y, x + 0.6, y - 0.75, LC_KONDICIONER, 0.03, JitterOn
    DrawWobbleLine x + aw - 0.6, y, x + aw - 0.6, y - 0.75, LC_KONDICIONER, 0.03, JitterOn
    ' гибкая подводка
    Dim px As Double: px = x + 0.9
    Dim pts(0 To 8, 0 To 1) As Double
    pts(0, 0) = px: pts(0, 1) = topY
    pts(1, 0) = px: pts(1, 1) = topY + 0.9
    Dim r As Double: r = 0.8
    For k = 1 To 6
        a = -1.5708 * k / 6
        pts(k + 1, 0) = px + r + r * Cos(a + 1.5708)
        pts(k + 1, 1) = topY + 0.9 + r * Sin(a + 1.5708)
    Next k
    pts(8, 0) = px + r: pts(8, 1) = topY + 2.1
    DrawPoly pts, LC_KONDICIONER, False
    mShapeCount = mShapeCount + 10
End Sub

Private Sub DrawCracks()
    Dim made As Long, tries As Long
    made = 0
    Do While made < 3 And tries < 60
        tries = tries + 1
        Dim si As Long: si = 1 + Int(RndD() * mNSec)
        If mSecType(si) = 4 Then GoTo retry
        Dim f As Long: f = Int(RndD() * mNF)
        Dim sx As Double, sy As Double
        sx = mSecX1(si) + 3 + RndD() * ((mSecX2(si) - mSecX1(si)) - 6)
        sy = mFY(f) + mFloorH * (0.2 + RndD() * 0.6)
        If Not PointFree(sx, sy, 0.7) Then GoTo retry
        Dim dirn As Double: dirn = IIf(RndD() < 0.5, -1, 1)
        Dim pts(0 To 4, 0 To 1) As Double
        Dim cx As Double, cy As Double, k As Long, ok As Boolean
        cx = sx: cy = sy: ok = True
        pts(0, 0) = cx: pts(0, 1) = cy
        For k = 1 To 4
            cx = cx + dirn * RndRange(1.3, 2.3)
            cy = cy + IIf(RndD() < 0.5, -1, 1) * RndRange(0.6, 1.2)
            If cx < mSecX1(si) + 1.5 Or cx > mSecX2(si) - 1.5 Then ok = False: Exit For
            If Not PointFree(cx, cy, 0.6) Then ok = False: Exit For
            pts(k, 0) = cx: pts(k, 1) = cy
        Next k
        If ok Then
            Dim trimmed() As Double
            ReDim trimmed(0 To k, 0 To 1)
            Dim j As Long
            For j = 0 To k
                trimmed(j, 0) = pts(j, 0): trimmed(j, 1) = pts(j, 1)
            Next j
            DrawPoly trimmed, LC_KLADKA, False
            made = made + 1
            mShapeCount = mShapeCount + 1
        End If
retry:
    Loop
End Sub

' ------------------------------------------------------------
' крыши: 0 плоская, 1 двускатная, 2 односкатная, 3 нет
' ------------------------------------------------------------
Private Sub DrawRoof()
    If RoofType = 3 Then Exit Sub
    If RoofType = 0 Then
        ' малые окна техэтажа
        Dim si As Long
        For si = 1 To mNSec
            If (si Mod 2 = 1) And mSecType(si) <> 4 Then
                Dim cx As Double: cx = (mSecX1(si) + mSecX2(si)) / 2
                Dim tw As Double: tw = MaxD(5#, (mSecX2(si) - mSecX1(si)) * 0.25)
                Dim tx As Double, ty As Double
                tx = cx - tw / 2: ty = mRoofBot + mRoofH * 0.18
                DrawRectWobble tx, ty, tw, mRoofH * 0.2, LC_KRYSHA, 0.07, JitterOn
                DrawRectWobble tx + 0.6, ty + 0.6, tw - 1.2, mRoofH * 0.2 - 1.2, LC_KRYSHA, 0.05, JitterOn
                AddObstacle tx - 0.3, ty - 0.3, tw + 0.6, mRoofH * 0.2 + 0.6
                mShapeCount = mShapeCount + 2
            End If
        Next si
        Dim yy As Double
        yy = mRoofBot + mRoofH * 0.55
        DrawWobbleLine mX0 + 0.5, yy, mX1 - 0.5, yy, LC_KRYSHA, 0.09, JitterOn
        yy = mRoofBot + mRoofH * 0.82
        DrawWobbleLine mX0 + 0.5, yy, mX1 - 0.5, yy, LC_KRYSHA, 0.09, JitterOn
        yy = mRoofBot + mRoofH * 0.95
        DrawWobbleLine mX0 + 0.5, yy, mX1 - 0.5, yy, LC_KRYSHA, 0.09, JitterOn
        ' вентшахты
        Dim vx As Variant
        For Each vx In Array(mX0 + (mX1 - mX0) * 0.2, mX0 + (mX1 - mX0) * 0.62)
            DrawRectWobble vx, mRoofBot + mRoofH * 0.68, mRoofH * 0.4, mRoofH * 0.17, LC_KRYSHA, 0.07, JitterOn
            DrawWobbleLine vx - 0.7, mRoofBot + mRoofH * 0.9, vx + mRoofH * 0.4 + 0.7, mRoofBot + mRoofH * 0.9, LC_KRYSHA, 0.05, JitterOn
            mShapeCount = mShapeCount + 2
        Next vx
        If Seed Mod 2 = 1 Then
            Dim ax As Double: ax = mX1 - 12
            DrawWobbleLine ax, mRoofBot + mRoofH * 0.95, ax, mRoofBot + mRoofH * 1.18, LC_KRYSHA, 0.04, JitterOn
            Dim hh As Variant
            For Each hh In Array(0.18, 0.1, 0.04)
                DrawWobbleLine ax - 1.4, mRoofBot + mRoofH * (0.95 + hh), ax + 1.4, mRoofBot + mRoofH * (0.95 + hh), LC_KRYSHA, 0.03, JitterOn
            Next hh
            mShapeCount = mShapeCount + 4
        End If
    ElseIf RoofType = 1 Then
        ' двускатная: карниз + два ската + конёк
        DrawWobbleLine mX0, mRoofBot, mX1, mRoofBot, LC_KRYSHA, 0.09, JitterOn
        Dim midX As Double: midX = (mX0 + mX1) / 2
        Dim ridgeY As Double: ridgeY = mRoofBot + mRoofH
        DrawWobbleLine mX0, mRoofBot, midX, ridgeY, LC_KRYSHA, 0.09, JitterOn
        DrawWobbleLine midX, ridgeY, mX1, mRoofBot, LC_KRYSHA, 0.09, JitterOn
        ' слуховое окно
        DrawRectWobble midX - 3, mRoofBot + mRoofH * 0.25, 6, mRoofH * 0.28, LC_KRYSHA, 0.06, JitterOn
        mShapeCount = mShapeCount + 4
    Else
        ' односкатная
        DrawWobbleLine mX0, mRoofBot, mX1, mRoofBot, LC_KRYSHA, 0.09, JitterOn
        DrawWobbleLine mX0, mRoofBot + mRoofH * 0.25, mX1, mRoofBot + mRoofH, LC_KRYSHA, 0.09, JitterOn
        mShapeCount = mShapeCount + 2
    End If
End Sub

' ------------------------------------------------------------
' маленькие хелперы
' ------------------------------------------------------------
Private Function MinL(a As Long, b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function

Private Function CeilD(v As Double) As Double
    CeilD = -Int(-v)
End Function
