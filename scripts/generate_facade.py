# -*- coding: utf-8 -*-
"""
Параметрический генератор векторного макета фактуры панельного дома
для лазерной гравировки на светильнике 320 x 200 мм (ЛМДФ).

Выходные файлы (output/):
  facade_panel_320x200.dxf  — DXF R2010, единицы мм, 2 слоя/цвета
  facade_panel_320x200.eps  — EPS, координаты в мм, 2 цвета контуров
  facade_panel_320x200.svg  — SVG-превью
  preview.png               — растровое превью для контроля

Слои:
  KONTUR  (красный, ACI 1) — несущие контуры: рамы окон, швы панелей,
                             парапеты, плиты, корпуса кондиционеров, рамка
  DETALI  (синий,  ACI 5)  — мелкая фактура: плитка стен, раскладка створок,
                             отливы, решётки ограждений, трубы, вентиляторы

Гарантия технологии: лазерный луч 0.2 мм — никакие два разноимённых
вектора не подходят параллельно ближе 0.30 мм (проверяется аудитом в конце).
"""

import math
import os

# ----------------------------------------------------------------------
# 0. КОНСТАНТЫ РАЗМЕРОВ (мм)
# ----------------------------------------------------------------------
PLATE_W, PLATE_H = 320.0, 200.0          # габарит светильника
MARGIN = 6.0                             # поле до гравировки
FX0, FY0 = MARGIN, MARGIN
FX1, FY1 = PLATE_W - MARGIN, PLATE_H - MARGIN

FLOORS = 7
PH = (FY1 - FY0) / FLOORS                # шаг этажа ~26.86 мм (≈2.7 м, M 1:100)

BEAM = 0.20                              # толщина луча
GAP_MIN = 0.30                           # минимальный зазор между векторами
TILE = 3.0                               # шаг плитки стен (≈0.3 м в натуре)
FRAME = 0.9                              # ширина оконной рамы (двойной контур)

L1, L2 = "KONTUR", "DETALI"

# вертикальные секции фасада (тип, ширина мм) — как на примере:
# узкий одинарный простенок, 2 балкона, пары окон, сплошное остекление,
# ещё пары окон, утопленные лоджии
STRIPS = [
    ("single",  24.0),
    ("balcony", 30.0),
    ("balcony", 30.0),
    ("double",  30.0),
    ("double",  30.0),
    ("glazed",  26.0),
    ("double",  30.0),
    ("double",  30.0),
    ("double",  30.0),
    ("loggia",  30.0),
]
JOINT = 2.0                              # межпанельный деф.шов (визуальный)

# кондиционеры: (индекс секции, этаж, положение)
AC_UNITS = [
    (3, 1, "L"), (4, 3, "R"), (1, 5, "C"),
    (6, 4, "R"), (8, 2, "L"), (9, 0, "L"),
    (3, 6, "R"), (6, 1, "L"),
]

# ----------------------------------------------------------------------
# 1. НАКОПИТЕЛЬ ГЕОМЕТРИИ
# ----------------------------------------------------------------------
ents = []        # (kind, layer, data)
obstacles = []   # прямоугольники-смысловые проёмы для клиппинга сетки


def line(x1, y1, x2, y2, layer):
    ents.append(("line", layer, (x1, y1, x2, y2)))


def rect(x, y, w, h, layer):
    ents.append(("rect", layer, (x, y, w, h)))


def circle(cx, cy, r, layer):
    ents.append(("circle", layer, (cx, cy, r)))


def poly(points, layer, closed=False):
    ents.append(("poly", layer, (tuple(points), closed)))


def obstacle(x, y, w, h):
    obstacles.append((x, y, x + w, y + h))


# ----------------------------------------------------------------------
# 2. КЛИППИНГ СЕТКИ ПРЯМОУГОЛЬНИКАМИ
# ----------------------------------------------------------------------
def _subtract(lo, hi, blocks):
    """вычесть из отрезка [lo,hi] набор перекрывающих интервалов"""
    for a, b in blocks:
        if b <= lo or a >= hi:
            continue
        if a <= lo and b >= hi:
            return []
        if a > lo:
            left = _subtract(lo, a, blocks)
        else:
            left = []
        if b < hi:
            right = _subtract(b, hi, blocks)
        else:
            right = []
        return left + right
    return [(lo, hi)]


def hline_clipped(y, xa, xb, layer, extra=None):
    blocks = [(max(xa, o[0] - GAP_MIN), min(xb, o[2] + GAP_MIN))
              for o in (obstacles + (extra or []))
              if o[1] - GAP_MIN < y < o[3] + GAP_MIN]
    for a, b in _subtract(xa, xb, blocks):
        if b - a >= 0.6:
            line(a, y, b, y, layer)


def vline_clipped(x, ya, yb, layer, extra=None):
    blocks = [(max(ya, o[1] - GAP_MIN), min(yb, o[3] + GAP_MIN))
              for o in (obstacles + (extra or []))
              if o[0] - GAP_MIN < x < o[2] + GAP_MIN]
    for a, b in _subtract(ya, yb, blocks):
        if b - a >= 0.6:
            line(x, a, x, b, layer)


# ----------------------------------------------------------------------
# 3. ЭЛЕМЕНТЫ ФАСАДА
# ----------------------------------------------------------------------
def draw_window(x, y, w, h, style):
    """окно: внешняя рама L1, внутренняя рама и створки L2, отлив L2"""
    rect(x, y, w, h, L1)
    rect(x + FRAME, y + FRAME, w - 2 * FRAME, h - 2 * FRAME, L2)
    ix0, iy0 = x + FRAME, y + FRAME
    iw, ih = w - 2 * FRAME, h - 2 * FRAME
    if style == 0:                                   # центральная стойка
        line(x + w / 2, iy0, x + w / 2, iy0 + ih, L2)
    elif style == 1:                                 # стойка + форточка
        line(x + w / 2, iy0, x + w / 2, iy0 + ih, L2)
        fy = iy0 + ih * 0.62
        line(x + w / 2, fy, ix0 + iw, fy, L2)
    elif style == 2:                                 # трёхстворчатое
        line(ix0 + iw / 3, iy0, ix0 + iw / 3, iy0 + ih, L2)
        line(ix0 + 2 * iw / 3, iy0, ix0 + 2 * iw / 3, iy0 + ih, L2)
    else:                                            # смещённая стойка
        line(x + w * 0.42, iy0, x + w * 0.42, iy0 + ih, L2)
    # отлив (параллельные линии в 0.7 мм)
    rect(x - 1.0, y - 1.1, w + 2.0, 0.7, L2)
    obstacle(x - 1.2, y - 1.3, w + 2.4, h + 1.5)


def draw_balcony(cx, yb, style):
    """балкон: остеклённый блок + сплошной парапет с решёткой"""
    gw, gh = 20.0, 18.0
    gx = cx - gw / 2
    gy = yb + 5.5
    rect(gx, gy, gw, gh, L1)
    rect(gx + FRAME, gy + FRAME, gw - 2 * FRAME, gh - 2 * FRAME, L2)
    line(gx + gw / 3, gy + FRAME, gx + gw / 3, gy + gh - FRAME, L2)
    line(gx + 2 * gw / 3, gy + FRAME, gx + 2 * gw / 3, gy + gh - FRAME, L2)
    line(gx + FRAME, gy + gh - 4.0, gx + gw - FRAME, gy + gh - 4.0, L2)
    obstacle(gx - 0.4, gy - 0.4, gw + 0.8, gh + 0.8)

    bw, bh = 24.0, 3.6
    bx, by = cx - bw / 2, yb + 1.2
    rect(bx, by, bw, bh, L1)
    obstacle(bx - 0.4, by - 0.4, bw + 0.8, bh + 0.8)
    # решётка внутри парапета (с учётом кондиционеров)
    line(bx + 0.6, by + bh - 1.3, bx + bw - 0.6, by + bh - 1.3, L2)
    x = bx + 1.4
    while x <= bx + bw - 1.0:
        vline_clipped(x, by + 0.5, by + bh - 1.3, L2)
        x += 1.8


def draw_glazed_strip(x0, x1, yb):
    """сплошное остекление с плиточным фартуком"""
    fx0, fx1 = x0 + 2.0, x1 - 2.0
    fy0, fy1 = yb + 9.0, yb + 23.6
    rect(fx0, fy0, fx1 - fx0, fy1 - fy0, L1)
    rect(fx0 + FRAME, fy0 + FRAME,
         fx1 - fx0 - 2 * FRAME, fy1 - fy0 - 2 * FRAME, L2)
    w = fx1 - fx0
    line(fx0 + w / 3, fy0 + FRAME, fx0 + w / 3, fy1 - FRAME, L2)
    line(fx0 + 2 * w / 3, fy0 + FRAME, fx0 + 2 * w / 3, fy1 - FRAME, L2)
    line(fx0 + FRAME, fy1 - 3.5, fx1 - FRAME, fy1 - 3.5, L2)
    obstacle(fx0 - 0.4, fy0 - 0.4, w + 0.8, fy1 - fy0 + 0.8)

    ax0, ay0 = x0 + 2.0, yb + 2.0
    aw, ah = (x1 - x0) - 4.0, 6.0
    rect(ax0, ay0, aw, ah, L1)
    obstacle(ax0 - 0.4, ay0 - 0.4, aw + 0.8, ah + 0.8)
    # мелкая плитка фартука, шаг 1.5 мм
    x = ax0 + 1.5
    while x < ax0 + aw - 0.4:
        line(x, ay0 + 0.35, x, ay0 + ah - 0.35, L2)
        x += 1.5
    y = ay0 + 1.5
    while y < ay0 + ah - 0.4:
        line(ax0 + 0.35, y, ax0 + aw - 0.35, y, L2)
        y += 1.5


def draw_loggia(x0, x1, yb):
    """утопленная лоджия: дверь, боковая стенка, линии плит"""
    rect(x1 - 2.4, yb + 0.6, 0.9, PH - 1.2, L1)                      # стенка
    dx, dw, dh = x0 + 2.5, 7.5, 16.5
    dy = yb + 6.0
    rect(dx, dy, dw, dh, L1)
    rect(dx + 0.8, dy + 0.8, dw - 1.6, dh - 1.6, L2)
    line(dx + dw - 1.7, dy + dh * 0.35, dx + dw - 1.7, dy + dh * 0.6, L2)
    line(x0 + 1.5, yb + 4.6, x1 - 2.4, yb + 4.6, L1)                 # плита (в стенку)
    line(x0 + 1.5, yb + 2.0, x1 - 2.4, yb + 2.0, L1)
    obstacle(x0 + 0.4, yb + 0.4, (x1 - x0) - 0.8, PH - 0.8)          # без сетки


def draw_ac(cx, top_y, h=2.8, w=4.4):
    """наружный блок кондиционера с вентилятором и изогнутой трубой"""
    x, y = cx - w / 2, top_y - h
    rect(x, y, w, h, L1)
    fan_cx, fan_cy = x + 1.2, y + h / 2
    circle(fan_cx, fan_cy, 0.85, L2)                    # зазор до корпуса 0.35
    circle(fan_cx, fan_cy, 0.38, L2)
    for i in range(3):
        gx = x + w - 1.9 + i * 0.7
        line(gx, y + 0.55, gx, y + h - 0.55, L2)
    # труба: вертикаль -> плавный отвод -> вертикаль к подоконнику
    px = x + 0.9
    pts = [(px, top_y), (px, top_y + 1.1)]
    r = 0.9
    for k in range(1, 7):
        a = -math.pi / 2 * k / 6
        pts.append((px + r + r * math.cos(a + math.pi / 2),
                    top_y + 1.1 + r * math.sin(a + math.pi / 2)))
    pts.append((px + r, top_y + 2.6))
    poly(pts, L2)
    obstacle(x - 0.45, y - 0.45, w + 0.9, h + 0.9)


# ----------------------------------------------------------------------
# 4. РАСКЛАДКА СЕКЦИЙ
# ----------------------------------------------------------------------
bounds = []
xacc = FX0
for i, (typ, w) in enumerate(STRIPS):
    x0 = xacc
    x1 = x0 + w
    bounds.append((typ, x0, x1))
    xacc = x1 + JOINT
# последнюю секцию растягиваем до правого поля гравировки
last = bounds[-1]
bounds[-1] = (last[0], last[1], FX1)
strip_x = bounds
assert abs(strip_x[-1][2] - FX1) < 1e-9

# ----------------------------------------------------------------------
# 5. ПЕРВЫЙ ПРОХОД — РЕГИСТРАЦИЯ ПРОЁМОВ (включая кондиционеры)
# ----------------------------------------------------------------------
ac_map = {}
for si, fl, pos in AC_UNITS:
    ac_map[(si, fl)] = pos

for f in range(FLOORS):
    yb = FY0 + f * PH
    for si, (typ, x0, x1) in enumerate(strip_x):
        cx = (x0 + x1) / 2
        if typ in ("single", "double"):
            if typ == "single":
                draw_window_obstacle = [(cx - 6.5, yb + 8.5, 13.0, 15.1)]
            else:
                draw_window_obstacle = [
                    (x0 + 2.0, yb + 8.5, 11.6, 15.1),
                    (x0 + 16.4, yb + 8.5, 11.6, 15.1),
                ]
            for wx, wy, ww, wh in draw_window_obstacle:
                obstacle(wx - 1.2, wy - 1.3, ww + 2.4, wh + 1.5)
        elif typ == "balcony":
            obstacle(cx - 10.4, yb + 5.1, 20.8, 18.8)
            obstacle(cx - 12.4, yb + 0.8, 24.8, 4.4)
        elif typ == "glazed":
            obstacle(x0 + 1.6, yb + 8.6, (x1 - x0) - 3.2, 15.4)
            obstacle(x0 + 1.6, yb + 1.6, (x1 - x0) - 3.2, 6.8)
        elif typ == "loggia":
            obstacle(x0 + 0.4, yb + 0.4, (x1 - x0) - 0.8, PH - 0.8)
        if (si, f) in ac_map:
            if typ == "balcony":
                obstacle(cx - 2.65, yb + 1.4, 5.3, 3.4)
            else:
                pos = ac_map[(si, f)]
                if typ == "loggia":
                    acx = x0 + 6.0
                elif pos == "L":
                    acx = x0 + 7.8
                elif pos == "R":
                    acx = x0 + 22.2
                else:
                    acx = cx
                obstacle(acx - 2.65, yb + 1.25, 5.3, 3.7)

# ----------------------------------------------------------------------
# 6. ВТОРОЙ ПРОХОД — ОТРИСОВКА
# ----------------------------------------------------------------------
# 6.1 рамка поля и межпанельные швы
rect(FX0, FY0, FX1 - FX0, FY1 - FY0, L1)
for f in range(1, FLOORS):
    y = FY0 + f * PH
    line(FX0, y, FX1, y, L1)
for i in range(len(strip_x) - 1):
    x = (strip_x[i][2] + strip_x[i + 1][1]) / 2
    line(x, FY0, x, FY1, L1)

# 6.2 заполнение секций
for f in range(FLOORS):
    yb = FY0 + f * PH
    for si, (typ, x0, x1) in enumerate(strip_x):
        cx = (x0 + x1) / 2
        style = (f * 7 + si * 3) % 4
        if typ == "single":
            draw_window(cx - 6.5, yb + 8.5, 13.0, 15.1, style)
        elif typ == "double":
            draw_window(x0 + 2.0, yb + 8.5, 11.6, 15.1, style)
            draw_window(x0 + 16.4, yb + 8.5, 11.6, 15.1, (style + 1) % 4)
        elif typ == "balcony":
            draw_balcony(cx, yb, style)
        elif typ == "glazed":
            draw_glazed_strip(x0, x1, yb)
        elif typ == "loggia":
            draw_loggia(x0, x1, yb)
        if (si, f) in ac_map:
            if typ == "balcony":
                draw_ac(cx, yb + 4.4, h=2.6)
            elif typ == "loggia":
                draw_ac(x0 + 6.0, yb + 4.0)
            else:
                pos = ac_map[(si, f)]
                acx = x0 + 7.8 if pos == "L" else (x0 + 22.2 if pos == "R" else cx)
                draw_ac(acx, yb + 4.5)

# 6.3 плиточная фактура стен (клиппинг по проёмам и швам)
for f in range(FLOORS):
    yb = FY0 + f * PH
    for si, (typ, x0, x1) in enumerate(strip_x):
        if typ == "loggia":
            continue
        xa, xb = x0 + 0.6, x1 - 0.6
        ya, ybnd = yb + 0.6, yb + PH - 0.6
        x = math.ceil(xa / TILE) * TILE
        while x <= xb:
            vline_clipped(x, ya, ybnd, L2)
            x += TILE
        y = math.ceil(ya / TILE) * TILE
        while y <= ybnd:
            hline_clipped(y, xa, xb, L2)
            y += TILE

# ----------------------------------------------------------------------
# 7. АУДИТ ЗАЗОРОВ (луч 0.2 мм → дистанция между векторами >= 0.30 мм)
# ----------------------------------------------------------------------
def _segments():
    out = []
    for kind, layer, d in ents:
        if kind == "line":
            out.append((layer, [(d[0], d[1]), (d[2], d[3])]))
        elif kind == "rect":
            x, y, w, h = d
            out.append((layer, [(x, y), (x + w, y), (x + w, y + h),
                                (x, y + h), (x, y)]))
        elif kind == "poly":
            pts, closed = d
            p = list(pts) + ([pts[0]] if closed else [])
            out.append((layer, p))
        elif kind == "circle":
            cx0, cy0, r = d
            p = [(cx0 + r * math.cos(2 * math.pi * k / 32),
                  cy0 + r * math.sin(2 * math.pi * k / 32))
                 for k in range(33)]
            out.append((layer, p))
    return out


def _clamp(v, lo=0.0, hi=1.0):
    return max(lo, min(hi, v))


def _seg_closest(a, b, c, d):
    """минимальное расстояние между двумя 2D-отрезками (Эриксон)"""
    ux, uy = b[0] - a[0], b[1] - a[1]
    vx, vy = d[0] - c[0], d[1] - c[1]
    wx, wy = a[0] - c[0], a[1] - c[1]
    A = ux * ux + uy * uy
    B = ux * vx + uy * vy
    C = vx * vx + vy * vy
    D = ux * wx + uy * wy
    E = vx * wx + vy * wy
    denom = A * C - B * B
    s = _clamp((B * E - C * D) / denom) if abs(denom) > 1e-12 else 0.0
    t = (B * s + E) / C
    if t < 0.0:
        t = 0.0
        s = _clamp(-D / A)
    elif t > 1.0:
        t = 1.0
        s = _clamp((B - D) / A)
    if s < 0.0:
        s = 0.0
        t = _clamp(E / C)
    elif s > 1.0:
        s = 1.0
        t = _clamp((B + E) / C)
    px, py = a[0] + s * ux, a[1] + s * uy
    qx, qy = c[0] + t * vx, c[1] + t * vy
    return math.hypot(px - qx, py - qy)


def _seg_angle(a, b, c, d):
    ux, uy = b[0] - a[0], b[1] - a[1]
    vx, vy = d[0] - c[0], d[1] - c[1]
    nu, nv = math.hypot(ux, uy), math.hypot(vx, vy)
    if nu < 1e-12 or nv < 1e-12:
        return 90.0
    cos_a = abs(ux * vx + uy * vy) / (nu * nv)
    return math.degrees(math.acos(_clamp(cos_a)))


def _parallel_overlap(a, b, c, d):
    """длина перекрытия проекций отрезков на направление первого"""
    ux, uy = b[0] - a[0], b[1] - a[1]
    nu = math.hypot(ux, uy)
    if nu < 1e-12:
        return 0.0
    ux, uy = ux / nu, uy / nu
    s1 = (0.0, nu)
    s2 = sorted(((c[0] - a[0]) * ux + (c[1] - a[1]) * uy,
                 (d[0] - a[0]) * ux + (d[1] - a[1]) * uy))
    return max(0.0, min(s1[1], s2[1]) - max(s1[0], s2[0]))


def audit_gaps():
    """
    Контроль технологии:
      * пересечения под углом >=12° (стыки, кресты, T) — допустимы;
      * параллельные векторы с перекрытием >0.5 мм и зазором <0.30 мм — брак;
      * выход за габарит плиты — брак.
    """
    chains = _segments()
    seen, uniq = set(), []
    for layer, p in chains:
        key = (layer, tuple(tuple(round(v, 3) for v in pt) for pt in p))
        if key not in seen:
            seen.add(key)
            uniq.append((layer, p))

    # кандидаты на сближение — через пространственные ячейки
    cell = GAP_MIN
    buckets = {}
    for eid, (_, p) in enumerate(uniq):
        for a, b in zip(p[:-1], p[1:]):
            n = max(2, int(math.hypot(b[0] - a[0], b[1] - a[1]) / 0.2) + 1)
            for k in range(n):
                t = k / (n - 1)
                px = a[0] + (b[0] - a[0]) * t
                py = a[1] + (b[1] - a[1]) * t
                key = (int(px // cell), int(py // cell))
                buckets.setdefault(key, set()).add(eid)
    candidates = set()
    for (i, j), eids in buckets.items():
        for di in (-1, 0, 1):
            for dj in (-1, 0, 1):
                for q in buckets.get((i + di, j + dj), ()):
                    for eid in eids:
                        if eid != q:
                            candidates.add((min(eid, q), max(eid, q)))

    violations = []
    for e1, e2 in sorted(candidates):
        p1, p2 = uniq[e1][1], uniq[e2][1]
        worst = None
        for a, b in zip(p1[:-1], p1[1:]):
            for c, d in zip(p2[:-1], p2[1:]):
                dist = _seg_closest(a, b, c, d)
                if dist >= GAP_MIN - 1e-9:
                    continue
                ang = _seg_angle(a, b, c, d)
                if ang >= 12.0:
                    continue                      # угловой стык — норма
                if _parallel_overlap(a, b, c, d) <= 0.5:
                    continue                      # торцевой контакт — норма
                if worst is None or dist < worst[0]:
                    worst = (dist, ang, (a, b), (c, d))
        if worst:
            violations.append((e1, e2, worst))

    out_of_bounds = [e for e in uniq
                     if any(px < -0.01 or px > PLATE_W + 0.01 or
                            py < -0.01 or py > PLATE_H + 0.01
                            for pt in e[1] for px, py in [pt])]
    return uniq, violations, out_of_bounds


uniq, violations, oob = audit_gaps()
print(f"примитивов: {len(ents)} | уникальных сегментных цепочек: {len(uniq)}")
print(f"параллельных наложений с зазором < {GAP_MIN} мм: {len(violations)}")
for e1, e2, (dist, ang, s1, s2) in violations[:15]:
    print(f"  d={dist:.3f} мм, угол={ang:.1f}°: {s1} <-> {s2}")
print(f"выходов за габарит {PLATE_W}x{PLATE_H}: {len(oob)}")
assert not violations and not oob, "аудит зазоров не пройден"

# ----------------------------------------------------------------------
# 8. ЭКСПОРТ DXF
# ----------------------------------------------------------------------
import ezdxf

doc = ezdxf.new("R2010", setup=True)
doc.units = ezdxf.units.MM
doc.header["$INSUNITS"] = 4          # миллиметры
doc.header["$MEASUREMENT"] = 1
doc.layers.add(name=L1, color=1)      # красный
doc.layers.add(name=L2, color=5)      # синий
msp = doc.modelspace()

for kind, layer, d in ents:
    if kind == "line":
        msp.add_line((d[0], d[1]), (d[2], d[3]), dxfattribs={"layer": layer})
    elif kind == "rect":
        x, y, w, h = d
        msp.add_lwpolyline(
            [(x, y), (x + w, y), (x + w, y + h), (x, y + h)],
            close=True, dxfattribs={"layer": layer})
    elif kind == "circle":
        msp.add_circle((d[0], d[1]), d[2], dxfattribs={"layer": layer})
    elif kind == "poly":
        pts, closed = d
        msp.add_lwpolyline(pts, close=closed, dxfattribs={"layer": layer})

OUT = os.path.join(os.path.dirname(__file__), "..", "output")
os.makedirs(OUT, exist_ok=True)
dxf_path = os.path.join(OUT, "facade_panel_320x200.dxf")
doc.saveas(dxf_path)
print("DXF ->", dxf_path)

# ----------------------------------------------------------------------
# 9. ЭКСПОРТ EPS (координаты в мм через scale, 2 цвета)
# ----------------------------------------------------------------------
PT = 72.0 / 25.4
eps_path = os.path.join(OUT, "facade_panel_320x200.eps")
lines_out = [
    "%!PS-Adobe-3.0 EPSF-3.0",
    "%%Title: Panel facade texture 320x200mm",
    "%%Creator: generate_facade.py",
    f"%%BoundingBox: 0 0 {PLATE_W * PT:.2f} {PLATE_H * PT:.2f}",
    "%%EndComments",
    f"{PT:.6f} {PT:.6f} scale",
    "0 setlinecap 1 setlinejoin",
    "/C1 { 0.90 0.05 0.05 setcolor 0.07 setlinewidth } bind def",
    "/C2 { 0.05 0.20 0.85 setcolor 0.07 setlinewidth } bind def",
]


def ps_color(layer):
    return "C1" if layer == L1 else "C2"


for kind, layer, d in ents:
    c = ps_color(layer)
    if kind == "line":
        lines_out += [c, f"{d[0]:.3f} {d[1]:.3f} moveto",
                      f"{d[2]:.3f} {d[3]:.3f} lineto stroke"]
    elif kind == "rect":
        x, y, w, h = d
        lines_out += [c, f"{x:.3f} {y:.3f} moveto",
                      f"{w:.3f} 0 rlineto 0 {h:.3f} rlineto",
                      f"{-w:.3f} 0 rlineto closepath stroke"]
    elif kind == "circle":
        lines_out += [c, "newpath",
                      f"{d[0]:.3f} {d[1]:.3f} {d[2]:.3f} 0 360 arc stroke"]
    elif kind == "poly":
        pts, _ = d
        lines_out += [c, f"{pts[0][0]:.3f} {pts[0][1]:.3f} moveto"]
        for px, py in pts[1:]:
            lines_out.append(f"{px:.3f} {py:.3f} lineto")
        lines_out.append("stroke")
lines_out += ["%%EOF"]
with open(eps_path, "w", encoding="ascii") as fh:
    fh.write("\n".join(lines_out))
print("EPS ->", eps_path)

# ----------------------------------------------------------------------
# 10. SVG-ПРЕВЬЮ
# ----------------------------------------------------------------------
svg_path = os.path.join(OUT, "facade_panel_320x200.svg")
svg = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{PLATE_W}mm" '
       f'height="{PLATE_H}mm" viewBox="0 0 {PLATE_W} {PLATE_H}">',
       '<rect width="100%" height="100%" fill="white"/>',
       f'<g transform="scale(1,-1) translate(0,-{PLATE_H})" fill="none" '
       'stroke-width="0.12">']
colors = {L1: "#e60d0d", L2: "#0d33d9"}
for kind, layer, d in ents:
    col = colors[layer]
    if kind == "line":
        svg.append(f'<line x1="{d[0]:.3f}" y1="{d[1]:.3f}" x2="{d[2]:.3f}" '
                   f'y2="{d[3]:.3f}" stroke="{col}"/>')
    elif kind == "rect":
        svg.append(f'<rect x="{d[0]:.3f}" y="{d[1]:.3f}" width="{d[2]:.3f}" '
                   f'height="{d[3]:.3f}" stroke="{col}"/>')
    elif kind == "circle":
        svg.append(f'<circle cx="{d[0]:.3f}" cy="{d[1]:.3f}" r="{d[2]:.3f}" '
                   f'stroke="{col}"/>')
    elif kind == "poly":
        pts, closed = d
        p = " ".join(f"{x:.3f},{y:.3f}" for x, y in pts)
        svg.append(f'<polyline points="{p}" stroke="{col}" '
                   f'fill="none"{"" if not closed else ""}/>')
svg.append("</g></svg>")
with open(svg_path, "w", encoding="utf-8") as fh:
    fh.write("\n".join(svg))
print("SVG ->", svg_path)

# ----------------------------------------------------------------------
# 11. PNG-ПРЕВЬЮ
# ----------------------------------------------------------------------
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection

fig, ax = plt.subplots(figsize=(16, 10), dpi=190)
segs_by_layer = {L1: [], L2: []}
for layer, p in uniq:
    for a, b in zip(p[:-1], p[1:]):
        segs_by_layer[layer].append([a, b])
for layer, col in ((L1, "#e60d0d"), (L2, "#0d33d9")):
    ax.add_collection(LineCollection(segs_by_layer[layer], colors=col,
                                     linewidths=0.55))
ax.set_xlim(-2, PLATE_W + 2)
ax.set_ylim(-2, PLATE_H + 2)
ax.set_aspect("equal")
ax.set_facecolor("white")
ax.set_xticks(range(0, 321, 20))
ax.set_yticks(range(0, 201, 20))
ax.grid(True, color="#cccccc", lw=0.3, alpha=0.6)
ax.set_title("Panel facade 320x200 mm — red: KONTUR, blue: DETALI")
png_path = os.path.join(OUT, "preview.png")
fig.savefig(png_path, bbox_inches="tight", facecolor="white")
print("PNG ->", png_path)
