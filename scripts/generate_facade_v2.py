# -*- coding: utf-8 -*-
"""
Генератор v2: фактура панельного дома для светильника-коробки.
4 стороны, каждая панель 320 x 165 мм.

Отличия от v1:
  * «живые» линии — детерминированное микро-дрожание всех векторов;
  * кирпичная кладка ложковой перевязкой вместо квадратной сетки;
  * 8 слоёв/цветов по классам объектов;
  * цоколь на фундаменте под первым этажом + секция плоской крыши с парапетом;
  * пересечения векторов разрешены; аудит ловит только параллельные слияния
    (два вектора с перекрытием >0.5 мм ближе 0.25 мм — борозды луча 0.2 мм
    сольются в одну);
  * больше деталей: ручки створок, отливы, водосточные трубы с хомутами,
    подъезд с козырьком, вентрешётки цоколя, трещины, вентшахты, антенна.
"""

import math
import os
import random

# ----------------------------------------------------------------------
# 0. КОНСТАНТЫ
# ----------------------------------------------------------------------
W, H = 320.0, 165.0
MARGIN = 4.0
X0, X1 = MARGIN, W - MARGIN                 # 4 .. 316

Z_TOP = 15.0                                # верх цоколя
FLOOR_ZONE_TOP = 140.0                      # верх жилой зоны (низ крыши)
FLOORS = 5
PH = (FLOOR_ZONE_TOP - Z_TOP) / FLOORS      # шаг этажа 25.0 мм (~2.5 м, M 1:100)

BEAM = 0.20
PARALLEL_MIN = 0.25                         # мин. зазор параллельных векторов
BRICK_PITCH = 1.8                           # высота кирпичного ряда
BRICK_LEN = 4.2                             # длина кирпича
FRAME = 0.9                                 # ширина рамы (двойной контур)

# 8 классов объектов: (слой, ACI-цвет, RGB)
LAYERS = [
    ("01_KARKAS",        1,  (214, 39, 40)),    # рамка, швы, плиты
    ("02_PROYOMY",       5,  (31, 95, 214)),    # окна, двери, отливы, ручки
    ("03_BALKONY",       3,  (44, 160, 44)),    # балконы и лоджии
    ("04_KLADKA",        32, (140, 90, 43)),    # кирпич, трещины
    ("05_KONDICIONERY",  6,  (197, 18, 197)),   # кондиционеры, кабели
    ("06_VODOSTOK",      4,  (15, 154, 160)),   # водосточные трубы
    ("07_TSEKOL",        30, (230, 138, 0)),    # фундамент/цоколь
    ("08_KRYSHA",        8,  (68, 68, 68)),     # парапет, плоская крыша
]
L = {name: name for name, _, _ in LAYERS}

STRIPS = [
    ("single",  22.0),
    ("balcony", 28.0),
    ("balcony", 28.0),
    ("double",  29.0),
    ("double",  29.0),
    ("glazed",  25.0),
    ("double",  29.0),
    ("double",  29.0),
    ("double",  29.0),
    ("loggia",  29.0),
]
JOINT = 2.0

# кандидаты для кондиционеров: только полосы с парными окнами
# (на балконах/лоджиях для блока нет технологичного места)
AC_CANDIDATES = [
    (3, 1, "L"), (4, 3, "R"), (6, 2, "R"), (8, 1, "L"),
    (3, 3, "R"), (6, 0, "L"), (8, 3, "R"), (4, 0, "L"),
    (7, 2, "R"), (3, 0, "R"),
]

# ----------------------------------------------------------------------
# 1. ГЕОМЕТРИЯ С «ЖИВЫМИ» ЛИНИЯМИ
# ----------------------------------------------------------------------
class Panel:
    def __init__(self, seed):
        self.rng = random.Random(seed)
        self.ents = []          # (kind, layer, data)
        self.obstacles = []     # прямоугольники проёмов для клиппинга кладки

    # --- сырые примитивы -------------------------------------------------
    def _add(self, kind, layer, data):
        self.ents.append((kind, layer, data))

    def line(self, x1, y1, x2, y2, layer, amp=0.15, step=3.2):
        """прямая с микро-дрожанием, концы фиксированы"""
        ln = math.hypot(x2 - x1, y2 - y1)
        n = max(1, int(ln / step))
        ux, uy = (x2 - x1) / ln, (y2 - y1) / ln
        px_, py_ = -uy, ux
        phase = self.rng.uniform(0, 2 * math.pi)
        pts = [(x1, y1)]
        for i in range(1, n):
            t = i / n
            bx, by = x1 + ux * ln * t, y1 + uy * ln * t
            off = amp * (0.65 * math.sin(t * math.pi * 3 + phase)
                         + 0.35 * self.rng.uniform(-1, 1))
            pts.append((bx + px_ * off, by + py_ * off))
        pts.append((x2, y2))
        self._add("poly", layer, (pts, False))

    def rect(self, x, y, w, h, layer, amp=0.14):
        """прямоугольник со слегка «гуляющими» углами и сторонами"""
        c = [(x + self.rng.uniform(-amp, amp),
              y + self.rng.uniform(-amp, amp)),
             (x + w + self.rng.uniform(-amp, amp),
              y + self.rng.uniform(-amp, amp)),
             (x + w + self.rng.uniform(-amp, amp),
              y + h + self.rng.uniform(-amp, amp)),
             (x + self.rng.uniform(-amp, amp),
              y + h + self.rng.uniform(-amp, amp))]
        for i in range(4):
            a, b = c[i], c[(i + 1) % 4]
            self.line(a[0], a[1], b[0], b[1], layer, amp=amp * 0.8)

    def circle(self, cx, cy, r, layer, amp=0.08):
        phase1 = self.rng.uniform(0, 2 * math.pi)
        phase2 = self.rng.uniform(0, 2 * math.pi)
        pts = []
        for k in range(31):
            a = 2 * math.pi * k / 30
            rr = r + amp * math.sin(3 * a + phase1) + amp * 0.5 * math.sin(5 * a + phase2)
            pts.append((cx + rr * math.cos(a), cy + rr * math.sin(a)))
        self._add("poly", layer, (pts, True))

    def polyline(self, pts, layer, amp=0.0, step=3.2):
        if amp <= 0:
            self._add("poly", layer, (pts, False))
        else:
            for a, b in zip(pts[:-1], pts[1:]):
                self.line(a[0], a[1], b[0], b[1], layer, amp=amp, step=step)

    def obstacle(self, x, y, w, h):
        self.obstacles.append((x, y, x + w, y + h))


# ----------------------------------------------------------------------
# 2. КЛИППИНГ КЛАДКИ ПРЯМОУГОЛЬНИКАМИ-ПРОЁМАМИ
# ----------------------------------------------------------------------
def _subtract(lo, hi, blocks):
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


CLIP_MARGIN = 0.55                      # отступ кладки от проёмов с запасом на дрожание


def h_clipped(p: Panel, y, xa, xb, layer, amp=0.08):
    blocks = [(max(xa, o[0] - CLIP_MARGIN), min(xb, o[2] + CLIP_MARGIN))
              for o in p.obstacles
              if o[1] - CLIP_MARGIN < y < o[3] + CLIP_MARGIN]
    for a, b in _subtract(xa, xb, blocks):
        if b - a >= 0.8:
            p.line(a, y, b, y, layer, amp=amp, step=2.6)


def v_clipped(p: Panel, x, ya, yb, layer, amp=0.08):
    blocks = [(max(ya, o[1] - CLIP_MARGIN), min(yb, o[3] + CLIP_MARGIN))
              for o in p.obstacles
              if o[0] - CLIP_MARGIN < x < o[2] + CLIP_MARGIN]
    for a, b in _subtract(ya, yb, blocks):
        if b - a >= 0.8:
            p.line(x, a, x, b, layer, amp=amp, step=2.6)


# ----------------------------------------------------------------------
# 3. ЭЛЕМЕНТЫ ФАСАДА
# ----------------------------------------------------------------------
def window(p: Panel, x, y, w, h, style):
    p.rect(x, y, w, h, L["02_PROYOMY"], amp=0.13)
    p.rect(x + FRAME, y + FRAME, w - 2 * FRAME, h - 2 * FRAME,
           L["02_PROYOMY"], amp=0.10)
    ix, iy = x + FRAME, y + FRAME
    iw, ih = w - 2 * FRAME, h - 2 * FRAME
    if style == 0:
        p.line(x + w / 2, iy, x + w / 2, iy + ih, L["02_PROYOMY"], amp=0.09)
    elif style == 1:
        p.line(x + w / 2, iy, x + w / 2, iy + ih, L["02_PROYOMY"], amp=0.09)
        fy = iy + ih * 0.62
        p.line(x + w / 2, fy, ix + iw, fy, L["02_PROYOMY"], amp=0.09)
    elif style == 2:
        p.line(ix + iw / 3, iy, ix + iw / 3, iy + ih, L["02_PROYOMY"], amp=0.09)
        p.line(ix + 2 * iw / 3, iy, ix + 2 * iw / 3, iy + ih,
               L["02_PROYOMY"], amp=0.09)
    else:
        p.line(x + w * 0.42, iy, x + w * 0.42, iy + ih,
               L["02_PROYOMY"], amp=0.09)
    # ручка створки — короткая зарубка рядом с фактической стойкой
    if style == 2:
        hx = ix + iw / 3 + 0.85                 # справа от левой стойки
    elif style == 3:
        hx = x + w * 0.42 - 0.9
    else:
        hx = x + w / 2 - 0.9                     # стили 0 и 1
    p.line(hx, y + h * 0.45, hx, y + h * 0.6,
           L["02_PROYOMY"], amp=0.05, step=1.0)
    # отлив (зазор до нижней рамы >= 0.7 мм)
    p.rect(x - 0.5, y - 1.3, w + 1.0, 0.6, L["02_PROYOMY"], amp=0.07)
    p.obstacle(x - 0.9, y - 1.5, w + 1.8, h + 1.7)


def balcony(p: Panel, cx, yb):
    gw, gh = 19.0, 16.5
    gx, gy = cx - gw / 2, yb + 5.2
    p.rect(gx, gy, gw, gh, L["03_BALKONY"], amp=0.13)
    p.rect(gx + FRAME, gy + FRAME, gw - 2 * FRAME, gh - 2 * FRAME,
           L["03_BALKONY"], amp=0.10)
    p.line(gx + gw / 3, gy + FRAME, gx + gw / 3, gy + gh - FRAME,
           L["03_BALKONY"], amp=0.09)
    p.line(gx + 2 * gw / 3, gy + FRAME, gx + 2 * gw / 3, gy + gh - FRAME,
           L["03_BALKONY"], amp=0.09)
    p.line(gx + FRAME, gy + gh - 3.6, gx + gw - FRAME, gy + gh - 3.6,
           L["03_BALKONY"], amp=0.09)
    p.obstacle(gx - 0.3, gy - 0.3, gw + 0.6, gh + 0.6)
    # парапет с решёткой
    bw, bh = 23.0, 3.4
    bx, by = cx - bw / 2, yb + 1.1
    p.rect(bx, by, bw, bh, L["03_BALKONY"], amp=0.11)
    p.obstacle(bx - 0.3, by - 0.3, bw + 0.6, bh + 0.6)
    p.line(bx + 0.6, by + bh - 1.1, bx + bw - 0.6, by + bh - 1.1,
           L["03_BALKONY"], amp=0.07, step=2.0)
    x = bx + 1.4
    while x <= bx + bw - 0.8:
        v_clipped(p, x, by + 0.4, by + bh - 1.1, L["03_BALKONY"], amp=0.06)
        x += 1.8


def glazed_strip(p: Panel, x0, x1, yb):
    fx0, fx1 = x0 + 2.0, x1 - 2.0
    fy0, fy1 = yb + 8.5, yb + 22.0
    p.rect(fx0, fy0, fx1 - fx0, fy1 - fy0, L["03_BALKONY"], amp=0.12)
    p.rect(fx0 + FRAME, fy0 + FRAME, fx1 - fx0 - 2 * FRAME,
           fy1 - fy0 - 2 * FRAME, L["03_BALKONY"], amp=0.09)
    w = fx1 - fx0
    p.line(fx0 + w / 3, fy0 + FRAME, fx0 + w / 3, fy1 - FRAME,
           L["03_BALKONY"], amp=0.08)
    p.line(fx0 + 2 * w / 3, fy0 + FRAME, fx0 + 2 * w / 3, fy1 - FRAME,
           L["03_BALKONY"], amp=0.08)
    p.line(fx0 + FRAME, fy1 - 3.2, fx1 - FRAME, fy1 - 3.2,
           L["03_BALKONY"], amp=0.08)
    p.obstacle(fx0 - 0.3, fy0 - 0.3, w + 0.6, fy1 - fy0 + 0.6)
    # бетонный фартук с горизонтальными опалубочными линиями
    ax0, ay0 = x0 + 2.0, yb + 1.8
    aw, ah = (x1 - x0) - 4.0, 5.6
    p.rect(ax0, ay0, aw, ah, L["03_BALKONY"], amp=0.10)
    p.obstacle(ax0 - 0.3, ay0 - 0.3, aw + 0.6, ah + 0.6)
    p.line(ax0 + 0.4, ay0 + ah * 0.5, ax0 + aw - 0.4, ay0 + ah * 0.5,
           L["03_BALKONY"], amp=0.07)


def loggia(p: Panel, x0, x1, yb):
    # боковая стенка, дверь, плиты
    p.rect(x1 - 2.4, yb + 0.6, 0.9, PH - 1.2, L["03_BALKONY"], amp=0.09)
    dx, dw, dh = x0 + 2.4, 7.2, 15.5
    dy = yb + 6.2
    p.rect(dx, dy, dw, dh, L["03_BALKONY"], amp=0.11)
    p.rect(dx + 0.8, dy + 0.8, dw - 1.6, dh - 1.6, L["03_BALKONY"], amp=0.08)
    p.line(dx + dw - 1.6, dy + dh * 0.35, dx + dw - 1.6, dy + dh * 0.6,
           L["03_BALKONY"], amp=0.06, step=1.2)
    p.line(x0 + 1.4, yb + 4.4, x1 - 2.4, yb + 4.4, L["03_BALKONY"], amp=0.09)
    p.line(x0 + 1.4, yb + 1.9, x1 - 2.4, yb + 1.9, L["03_BALKONY"], amp=0.09)
    p.obstacle(x0 + 0.3, yb + 0.3, (x1 - x0) - 0.6, PH - 0.6)


def entrance(p: Panel, x0, x1, yb):
    """подъезд на месте лоджии первого этажа: дверь, козырёк, ступени"""
    dx, dw, dh = x0 + 4.0, 11.0, 16.9
    dy = yb + 1.8
    p.rect(dx, dy, dw, dh, L["02_PROYOMY"], amp=0.12)
    p.rect(dx + 0.8, dy + 0.8, dw - 1.6, dh - 1.6, L["02_PROYOMY"], amp=0.09)
    p.line(dx + dw / 2, dy + 0.8, dx + dw / 2, dy + dh - 0.8,
           L["02_PROYOMY"], amp=0.08)
    p.line(dx + dw / 2 - 1.0, dy + dh * 0.5, dx + dw / 2 - 1.0,
           dy + dh * 0.62, L["02_PROYOMY"], amp=0.05, step=1.0)
    p.line(dx + dw / 2 + 1.0, dy + dh * 0.5, dx + dw / 2 + 1.0,
           dy + dh * 0.62, L["02_PROYOMY"], amp=0.05, step=1.0)
    # козырёк
    p.line(dx - 2.2, dy + dh + 1.0, dx + dw + 2.2, dy + dh + 1.0,
           L["02_PROYOMY"], amp=0.09)
    p.line(dx - 1.4, dy + dh + 2.0, dx + dw + 1.4, dy + dh + 2.0,
           L["02_PROYOMY"], amp=0.09)
    # ступени (зазоры до порога и линии перекрытия >= 0.5 мм)
    p.line(dx - 2.0, dy - 0.55, dx + dw + 2.0, dy - 0.55,
           L["02_PROYOMY"], amp=0.07)
    p.line(dx - 3.2, dy - 1.1, dx + dw + 3.2, dy - 1.1,
           L["02_PROYOMY"], amp=0.07)
    p.obstacle(x0 + 0.3, yb + 0.3, (x1 - x0) - 0.6, PH - 0.6)


def ac_unit(p: Panel, cx, top_y, h=2.7, w=4.4):
    x, y = cx - w / 2, top_y - h
    p.rect(x, y, w, h, L["05_KONDICIONERY"], amp=0.09)
    fan_cx, fan_cy = x + 1.2, y + h / 2
    p.circle(fan_cx, fan_cy, 0.85, L["05_KONDICIONERY"], amp=0.06)
    p.circle(fan_cx, fan_cy, 0.36, L["05_KONDICIONERY"], amp=0.05)
    for i in range(3):
        gx = x + w - 1.9 + i * 0.7
        p.line(gx, y + 0.5, gx, y + h - 0.5, L["05_KONDICIONERY"],
               amp=0.05, step=1.0)
    # гибкая труба с плавным отводом
    px = x + 0.9
    pts = [(px, top_y), (px, top_y + 1.0)]
    r = 0.9
    for k in range(1, 7):
        a = -math.pi / 2 * k / 6
        pts.append((px + r + r * math.cos(a + math.pi / 2),
                    top_y + 1.0 + r * math.sin(a + math.pi / 2)))
    pts.append((px + r, top_y + 2.4))
    p.polyline(pts, L["05_KONDICIONERY"], amp=0.07, step=1.2)
    p.obstacle(x - 0.3, y - 0.3, w + 0.6, h + 0.6)


def downpipe(p: Panel, x, y_lo, y_hi):
    """водосточная труба: две близкие параллельные жилы + хомуты"""
    gap = 0.85
    p.line(x - gap / 2, y_lo, x - gap / 2, y_hi, L["06_VODOSTOK"],
           amp=0.10, step=4.0)
    p.line(x + gap / 2, y_lo, x + gap / 2, y_hi, L["06_VODOSTOK"],
           amp=0.10, step=4.0)
    y = y_lo + 3.0
    while y < y_hi - 1:
        p.line(x - 0.85, y, x + 0.85, y, L["06_VODOSTOK"], amp=0.06, step=1.2)
        y += PH


def _point_free(p: Panel, x, y, margin=0.6):
    for ox0, oy0, ox1, oy1 in p.obstacles:
        if ox0 - margin < x < ox1 + margin and oy0 - margin < y < oy1 + margin:
            return False
    return True


def crack(p: Panel, x, y, x_lo, x_hi, direction=1):
    """трещина из диагональных сегментов только по свободной стене"""
    pts = [(x, y)]
    cx, cy = x, y
    for _ in range(p.rng.randint(3, 5)):
        nx = min(x_hi, max(x_lo, cx + direction * p.rng.uniform(1.5, 2.6)))
        cy += p.rng.choice((-1, 1)) * p.rng.uniform(0.7, 1.4)
        pts.append((nx, cy))
        cx = nx
    if all(_point_free(p, px, py) for px, py in pts):
        p.polyline(pts, L["04_KLADKA"], amp=0.0)


# ----------------------------------------------------------------------
# 4. РАСКЛАДКА СЕКЦИЙ
# ----------------------------------------------------------------------
def make_strips():
    bounds, xacc = [], X0
    for typ, w in STRIPS:
        bounds.append((typ, xacc, xacc + w))
        xacc += w + JOINT
    last = bounds[-1]
    bounds[-1] = (last[0], last[1], X1)
    return bounds


# ----------------------------------------------------------------------
# 5. СБОРКА ОДНОЙ ПАНЕЛИ
# ----------------------------------------------------------------------
def build_panel(seed):
    p = Panel(seed)
    strips = make_strips()
    rng = p.rng

    # ---- 5.1 проёмы (нужны до кладки) ----------------------------------
    ac_map = {}
    chosen = rng.sample(AC_CANDIDATES, 8)
    for si, fl, pos in chosen:
        ac_map[(si, fl)] = pos

    for f in range(FLOORS):
        yb = Z_TOP + f * PH
        for si, (typ, x0, x1) in enumerate(strips):
            cx = (x0 + x1) / 2
            if typ == "single":
                p.obstacle(cx - 6.4 - 1.1, yb + 7.8 - 1.2, 12.8 + 2.2, 14.6 + 1.4)
            elif typ == "double":
                for wx in (x0 + 1.8, x0 + 15.6):
                    p.obstacle(wx - 0.9, yb + 7.8 - 1.2, 10.8 + 1.8, 14.6 + 1.4)
            elif typ == "balcony":
                p.obstacle(cx - 9.8, yb + 4.9, 19.6, 17.1)
                p.obstacle(cx - 11.8, yb + 0.8, 23.6, 4.0)
            elif typ == "glazed":
                p.obstacle(x0 + 1.7, yb + 8.2, (x1 - x0) - 3.4, 14.1)
                p.obstacle(x0 + 1.7, yb + 1.5, (x1 - x0) - 3.4, 6.2)
            elif typ == "loggia":
                if f == 0:
                    p.obstacle(x0 + 0.3, yb + 0.3, (x1 - x0) - 0.6, PH - 0.6)
                else:
                    p.obstacle(x0 + 0.3, yb + 0.3, (x1 - x0) - 0.6, PH - 0.6)
            if (si, f) in ac_map and typ == "double":
                pos = ac_map[(si, f)]
                acx = (x0 + 7.5 if pos == "L" else
                       x0 + 21.5 if pos == "R" else cx)
                p.obstacle(acx - 2.5, yb + 1.1, 5.0, 3.4)

    # ---- 5.2 каркас: рамка панели, швы, плиты ---------------------------
    # рамка строго по кромкам 320x165 (физическая граница панели — без дрожания)
    p.line(0, 0, W, 0, L["01_KARKAS"], amp=0.0)
    p.line(W, 0, W, H, L["01_KARKAS"], amp=0.0)
    p.line(W, H, 0, H, L["01_KARKAS"], amp=0.0)
    p.line(0, H, 0, 0, L["01_KARKAS"], amp=0.0)
    # межэтажные швы (границы зон рисуются отдельно, без дублей)
    for f in range(1, FLOORS):
        y = Z_TOP + f * PH
        p.line(X0, y, X1, y, L["01_KARKAS"], amp=0.16)
    for i in range(len(strips) - 1):
        x = (strips[i][2] + strips[i + 1][1]) / 2
        p.line(x, MARGIN, x, H - MARGIN, L["01_KARKAS"], amp=0.16)

    # ---- 5.3 цоколь на фундаменте ---------------------------------------
    p.line(X0, Z_TOP, X1, Z_TOP, L["01_KARKAS"], amp=0.15)
    for yy in (7.2, 10.8):
        p.line(X0 + 1.0, yy, X1 - 1.0, yy, L["07_TSEKOL"], amp=0.12)
    for vx in (40.0, 158.0, 272.0):
        p.rect(vx, 5.6, 7.0, 3.2, L["07_TSEKOL"], amp=0.08)
        for k in range(1, 4):
            p.line(vx + k * 1.4, 5.9, vx + k * 1.4, 8.5,
                   L["07_TSEKOL"], amp=0.05, step=1.0)
        p.obstacle(vx - 0.2, 5.4, 7.4, 3.6)

    # ---- 5.4 жилая зона --------------------------------------------------
    for f in range(FLOORS):
        yb = Z_TOP + f * PH
        for si, (typ, x0, x1) in enumerate(strips):
            cx = (x0 + x1) / 2
            style = (f * 7 + si * 3 + seed) % 4
            if typ == "single":
                window(p, cx - 6.4, yb + 7.8, 12.8, 14.6, style)
            elif typ == "double":
                window(p, x0 + 1.8, yb + 7.8, 10.8, 14.6, style)
                window(p, x0 + 15.6, yb + 7.8, 10.8, 14.6, (style + 1) % 4)
            elif typ == "balcony":
                balcony(p, cx, yb)
            elif typ == "glazed":
                glazed_strip(p, x0, x1, yb)
            elif typ == "loggia":
                if f == 0:
                    entrance(p, x0, x1, yb)
                else:
                    loggia(p, x0, x1, yb)
            if (si, f) in ac_map and typ == "double":
                pos = ac_map[(si, f)]
                acx = (x0 + 7.5 if pos == "L" else
                       x0 + 21.5 if pos == "R" else cx)
                ac_unit(p, acx, yb + 4.2)

    # ---- 5.5 кирпичная кладка (ложковая перевязка) ----------------------
    for f in range(FLOORS):
        ya = Z_TOP + f * PH + 0.5
        ybnd = Z_TOP + (f + 1) * PH - 0.5
        for si, (typ, x0, x1) in enumerate(strips):
            if typ == "loggia":
                continue
            xa, xb = x0 + 0.55, x1 - 0.55
            row = 0
            y = math.ceil(ya / BRICK_PITCH) * BRICK_PITCH
            while y <= ybnd:
                h_clipped(p, y, xa, xb, L["04_KLADKA"], amp=0.07)
                stagger = (BRICK_LEN / 2) if row % 2 else 0.0
                x = math.ceil((xa - stagger) / BRICK_LEN) * BRICK_LEN + stagger
                while x <= xb:
                    v_clipped(p, x, y + 0.12,
                              min(y + BRICK_PITCH - 0.12, ybnd),
                              L["04_KLADKA"], amp=0.06)
                    x += BRICK_LEN
                y += BRICK_PITCH
                row += 1

    # случайные трещины (2-3 на панель, только в полосах с кладкой)
    brick_strips = [s for s in strips if s[0] != "loggia"]
    made = 0
    for _try in range(40):
        if made >= rng.randint(2, 3):
            break
        typ, x0, x1 = rng.choice(brick_strips)
        f = rng.randrange(FLOORS)
        sx = rng.uniform(x0 + 4, x1 - 8)
        sy = Z_TOP + f * PH + rng.uniform(4, PH - 4)
        if not _point_free(p, sx, sy):
            continue
        before = len(p.ents)
        crack(p, sx, sy, x0 + 1.8, x1 - 1.8, direction=rng.choice((1, -1)))
        if len(p.ents) > before:
            made += 1

    # ---- 5.6 водостоки (со смещением от межпанельного шва) ---------------
    for si in (2, 6):
        joint_x = (strips[si][2] + strips[si + 1][1]) / 2
        downpipe(p, joint_x + 0.9, Z_TOP + 0.5, FLOOR_ZONE_TOP - 0.5)

    # ---- 5.7 плоская крыша: техполоса, парапет, вентшахты, антенна -------
    p.line(X0, FLOOR_ZONE_TOP, X1, FLOOR_ZONE_TOP, L["01_KARKAS"], amp=0.15)
    # маленькие окна техэтажа
    for si, (typ, x0, x1) in enumerate(strips):
        if typ in ("single", "double") and si % 2 == 0:
            cx = (x0 + x1) / 2
            p.rect(cx - 3.2, FLOOR_ZONE_TOP + 2.2, 6.4, 4.6,
                   L["08_KRYSHA"], amp=0.09)
            p.obstacle(cx - 3.4, FLOOR_ZONE_TOP + 2.0, 6.8, 5.0)
    # парапет и крышка
    p.line(X0 + 0.5, 153.5, X1 - 0.5, 153.5, L["08_KRYSHA"], amp=0.12)
    p.line(X0 + 0.5, 158.2, X1 - 0.5, 158.2, L["08_KRYSHA"], amp=0.12)
    p.line(X0 + 0.5, 160.2, X1 - 0.5, 160.2, L["08_KRYSHA"], amp=0.12)
    # вентшахты на крыше
    for vx, vw in ((52, 9), (150, 12), (246, 8)):
        p.rect(vx, 157.4, vw, 3.4, L["08_KRYSHA"], amp=0.09)
        p.line(vx + 1.0, 159.1, vx + vw - 1.0, 159.1,
               L["08_KRYSHA"], amp=0.06, step=1.2)
    # антенна только у части панелей
    if seed % 2 == 1:
        ax_ = 295.0
        p.line(ax_, 160.6, ax_, 164.0, L["08_KRYSHA"], amp=0.06, step=1.5)
        p.line(ax_ - 1.6, 163.0, ax_ + 1.6, 163.0, L["08_KRYSHA"],
               amp=0.05, step=1.0)
        p.line(ax_ - 1.1, 161.9, ax_ + 1.1, 161.9, L["08_KRYSHA"],
               amp=0.05, step=1.0)
        p.line(ax_ - 0.8, 160.6, ax_ + 0.8, 160.6, L["08_KRYSHA"],
               amp=0.05, step=1.0)

    return p


# ----------------------------------------------------------------------
# 6. АУДИТ: пересечения разрешены, контроль только параллельных слияний
# ----------------------------------------------------------------------
def _chains(ents):
    out = []
    for kind, layer, d in ents:
        if kind == "poly":
            pts, closed = d
            out.append((layer, list(pts) + ([pts[0]] if closed else [])))
    return out


def _clamp(v):
    return max(0.0, min(1.0, v))


def _point_seg_dist(p, a, b):
    ux, uy = b[0] - a[0], b[1] - a[1]
    nu2 = ux * ux + uy * uy
    if nu2 < 1e-14:
        return math.hypot(p[0] - a[0], p[1] - a[1])
    t = _clamp(((p[0] - a[0]) * ux + (p[1] - a[1]) * uy) / nu2)
    return math.hypot(p[0] - (a[0] + t * ux), p[1] - (a[1] + t * uy))


def _seg_dist(a, b, c, d):
    # вырожденные отрезки — через расстояние точка-отрезок
    if math.hypot(b[0] - a[0], b[1] - a[1]) < 1e-12:
        return _point_seg_dist(a, c, d)
    if math.hypot(d[0] - c[0], d[1] - c[1]) < 1e-12:
        return _point_seg_dist(c, a, b)
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
    if t < 0:
        t, s = 0, _clamp(-D / A)
    elif t > 1:
        t, s = 1, _clamp((B - D) / A)
    if s < 0:
        s, t = 0, _clamp(E / C)
    elif s > 1:
        s, t = 1, _clamp((B + E) / C)
    return math.hypot(a[0] + s * ux - (c[0] + t * vx),
                      a[1] + s * uy - (c[1] + t * vy))


def _angle(a, b, c, d):
    ux, uy = b[0] - a[0], b[1] - a[1]
    vx, vy = d[0] - c[0], d[1] - c[1]
    nu, nv = math.hypot(ux, uy), math.hypot(vx, vy)
    if nu < 1e-12 or nv < 1e-12:
        return 90.0
    return math.degrees(math.acos(_clamp(abs(ux * vx + uy * vy) / (nu * nv))))


def _overlap(a, b, c, d):
    ux, uy = b[0] - a[0], b[1] - a[1]
    nu = math.hypot(ux, uy)
    if nu < 1e-12:
        return 0.0
    ux, uy = ux / nu, uy / nu
    s2 = sorted(((c[0] - a[0]) * ux + (c[1] - a[1]) * uy,
                 (d[0] - a[0]) * ux + (d[1] - a[1]) * uy))
    return max(0.0, min(nu, s2[1]) - max(0.0, s2[0]))


def audit(p: Panel, tag):
    chains = _chains(p.ents)
    cell = PARALLEL_MIN
    buckets, candidates = {}, set()
    for eid, (_, ch) in enumerate(chains):
        touched = set()
        for a, b in zip(ch[:-1], ch[1:]):
            n = max(2, int(math.hypot(b[0] - a[0], b[1] - a[1]) / 0.2) + 1)
            for k in range(n):
                t = k / (n - 1)
                px, py = a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t
                key = (int(px // cell), int(py // cell))
                touched.add(key)
                buckets.setdefault(key, set()).add(eid)
        for key in touched:
            i, j = key
            for di in (-1, 0, 1):
                for dj in (-1, 0, 1):
                    for q in buckets.get((i + di, j + dj), ()):
                        if q != eid:
                            candidates.add((min(eid, q), max(eid, q)))
    violations = []
    for e1, e2 in sorted(candidates):
        ch1, ch2 = chains[e1][1], chains[e2][1]
        worst = None
        for a, b in zip(ch1[:-1], ch1[1:]):
            for c, d in zip(ch2[:-1], ch2[1:]):
                dist = _seg_dist(a, b, c, d)
                if dist >= PARALLEL_MIN:
                    continue
                if _angle(a, b, c, d) >= 12.0:
                    continue                       # пересечение под углом — ок
                if _overlap(a, b, c, d) <= 0.5:
                    continue                       # торцевой контакт — ок
                if worst is None or dist < worst:
                    worst = dist
        if worst is not None:
            violations.append((e1, e2, worst))
    oob = sum(1 for _, ch in chains for pt in ch
              if pt[0] < -0.02 or pt[0] > W + 0.02 or
              pt[1] < -0.02 or pt[1] > H + 0.02)
    print(f"[{tag}] примитивов: {len(p.ents)} | "
          f"параллельных слияний < {PARALLEL_MIN} мм: {len(violations)} | "
          f"за габаритом: {oob}")
    for e1, e2, dist in violations[:10]:
        print(f"   d={dist:.3f}: {chains[e1][0]} <-> {chains[e2][0]}")
    assert not violations and not oob, f"аудит панели {tag} не пройден"


# ----------------------------------------------------------------------
# 7. ЭКСПОРТ DXF
# ----------------------------------------------------------------------
def save_dxf(panels, path, per_side=False):
    import ezdxf
    doc = ezdxf.new("R2010", setup=True)
    doc.units = ezdxf.units.MM
    doc.header["$INSUNITS"] = 4
    doc.header["$MEASUREMENT"] = 1
    for name, aci, rgb in LAYERS:
        layer = doc.layers.add(name=name, color=aci)
        layer.rgb = rgb
    msp = doc.modelspace()

    def emit(ents, ox, oy):
        for kind, layer, d in ents:
            if kind != "poly":
                continue
            pts, closed = d
            shifted = [(x + ox, y + oy) for x, y in pts]
            msp.add_lwpolyline(shifted, close=closed,
                               dxfattribs={"layer": layer})

    if per_side:
        emit(panels[0].ents, 0, 0)
    else:
        gap = 15.0
        for i, p in enumerate(panels):
            col, row = i % 2, i // 2
            ox, oy = col * (W + gap), (1 - row) * (H + gap)
            emit(p.ents, ox, oy)
    doc.saveas(path)
    print("DXF ->", path)


# ----------------------------------------------------------------------
# 8. ЭКСПОРТ EPS (8 цветов, координаты в мм)
# ----------------------------------------------------------------------
def save_eps(panels, path):
    gap = 15.0
    sheet_w, sheet_h = 2 * W + gap, 2 * H + gap
    PT = 72.0 / 25.4
    out = [
        "%!PS-Adobe-3.0 EPSF-3.0",
        "%%Title: Panel facade 4 sides 320x165mm",
        "%%Creator: generate_facade_v2.py",
        f"%%BoundingBox: 0 0 {sheet_w * PT:.2f} {sheet_h * PT:.2f}",
        "%%EndComments",
        f"{PT:.6f} {PT:.6f} scale",
        "0 setlinecap 1 setlinejoin",
    ]
    for i, (name, aci, (r, g, b)) in enumerate(LAYERS, start=1):
        out.append(f"/C{i} {{ {r/255:.3f} {g/255:.3f} {b/255:.3f} "
                   f"setcolor 0.07 setlinewidth }} bind def")
    layer_proc = {name: f"C{i}" for i, (name, _, _) in enumerate(LAYERS, 1)}

    for i, p in enumerate(panels):
        col, row = i % 2, i // 2
        ox, oy = col * (W + gap), (1 - row) * (H + gap)
        out.append(f"%% side {i + 1}")
        out.append(f"{ox:.3f} {oy:.3f} translate")
        for kind, layer, d in p.ents:
            pts, closed = d
            proc = layer_proc[layer]
            out.append(proc)
            out.append(f"{pts[0][0]:.3f} {pts[0][1]:.3f} moveto")
            for x, y in pts[1:]:
                out.append(f"{x:.3f} {y:.3f} lineto")
            out.append("closepath stroke" if closed else "stroke")
        out.append(f"{-ox:.3f} {-oy:.3f} translate")
    out.append("%%EOF")
    with open(path, "w", encoding="ascii") as fh:
        fh.write("\n".join(out))
    print("EPS ->", path)


# ----------------------------------------------------------------------
# 9. ЗАПУСК
# ----------------------------------------------------------------------
def main():
    out_dir = os.path.join(os.path.dirname(__file__), "..", "output")
    os.makedirs(out_dir, exist_ok=True)
    panels = [build_panel(seed) for seed in (101, 202, 303, 404)]
    for i, p in enumerate(panels, 1):
        audit(p, f"side {i}")

    save_eps(panels, os.path.join(out_dir, "facade_4sides_320x165.eps"))
    save_dxf(panels, os.path.join(out_dir, "facade_4sides_320x165.dxf"))
    for i, p in enumerate(panels, 1):
        save_dxf([p], os.path.join(out_dir, f"facade_side{i}_320x165.dxf"),
                 per_side=True)


if __name__ == "__main__":
    main()
