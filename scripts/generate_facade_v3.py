# -*- coding: utf-8 -*-
"""
Генератор v3: фактура панельного дома для светильника-коробки.
4 стороны, каждая панель: ШИРИНА 165 мм, ВЫСОТА 320 мм (портрет).

Отличия от v2:
  * портретная ориентация 165x320, 9 жилых этажей + цоколь + техэтаж крыши;
  * плотная детализация по эталону:
      - мелкая квадратная плитка стен (шаг 1.5 мм) + кирпичные подоконные
        пояса ложковой перевязкой;
      - окна с двойной рамой, отдельными рамами створок, форточками с
        собственным переплётом, ручками, двойным отливом с капельником;
      - балконы: верхний ряд мелких стекол + крупные створки, парапет с
        панельной филёнкой и решёткой, плитные простенки;
      - подъезд с фрамугой, козырьком и ступенями;
      - водостоки с воронками и хомутами, кондиционеры с кронштейнами и
        лопатками вентилятора, трещины, вентшахты, антенна, русты цоколя;
  * живые линии (детерминированное микро-дрожание);
  * 8 слоёв/цветов по классам объектов;
  * пересечения разрешены; аудит — только параллельные слияния < 0.25 мм.
"""

import math
import os
import random

# ----------------------------------------------------------------------
# 0. КОНСТАНТЫ
# ----------------------------------------------------------------------
W, H = 165.0, 320.0                 # ШИРИНА x ВЫСОТА стороны, мм
MARGIN = 4.0
X0, X1 = MARGIN, W - MARGIN         # 4 .. 161

PLINTH_TOP = 20.0                   # верх цоколя
ROOF_BOT = 284.0                    # низ техэтажа/крыши
FLOORS = 9
PH = (ROOF_BOT - PLINTH_TOP) / FLOORS   # шаг этажа 29.33 мм (~2.9 м, M 1:100)

PARALLEL_MIN = 0.25                 # мин. зазор параллельных векторов
GRID_PITCH = 1.5                    # шаг квадратной плитки стен
BRICK_PITCH = 1.7                   # высота кирпичного ряда
BRICK_LEN = 4.0                     # длина кирпича
CLIP_MARGIN = 0.55                  # отступ текстур от проёмов

LAYERS = [
    ("01_KARKAS",        1,  (214, 39, 40)),
    ("02_PROYOMY",       5,  (31, 95, 214)),
    ("03_BALKONY",       3,  (44, 160, 44)),
    ("04_KLADKA",        32, (140, 90, 43)),
    ("05_KONDICIONERY",  6,  (197, 18, 197)),
    ("06_VODOSTOK",      4,  (15, 154, 160)),
    ("07_TSEKOL",        30, (230, 138, 0)),
    ("08_KRYSHA",        8,  (68, 68, 68)),
]
L = {name: name for name, _, _ in LAYERS}

# 5 вертикальных секций (портрет, узкая панель)
STRIPS = [
    ("single",  24.0),
    ("balcony", 32.0),
    ("double",  34.0),
    ("glazed",  28.0),
    ("loggia",  30.0),
]
JOINT = 2.0

AC_CANDIDATES = [
    (2, 1, "L"), (2, 4, "R"), (2, 7, "L"), (2, 3, "R"),
    (2, 6, "L"), (2, 8, "R"), (2, 2, "R"), (2, 5, "L"),
]
DP_OFFSET = -1.3                     # смещение водостока от панельного шва (в сторону предыдущей секции)

# ----------------------------------------------------------------------
# 1. ПАНЕЛЬ С ЖИВЫМИ ЛИНИЯМИ
# ----------------------------------------------------------------------
class Panel:
    def __init__(self, seed):
        self.rng = random.Random(seed)
        self.ents = []
        self.obstacles = []

    def _add(self, kind, layer, data):
        self.ents.append((kind, layer, data))

    def line(self, x1, y1, x2, y2, layer, amp=0.12, step=3.0):
        ln = math.hypot(x2 - x1, y2 - y1)
        if ln < 1e-9:
            return
        n = max(1, int(ln / step))
        ux, uy = (x2 - x1) / ln, (y2 - y1) / ln
        nx_, ny_ = -uy, ux
        phase = self.rng.uniform(0, 2 * math.pi)
        pts = [(x1, y1)]
        for i in range(1, n):
            t = i / n
            bx, by = x1 + ux * ln * t, y1 + uy * ln * t
            off = amp * (0.65 * math.sin(t * math.pi * 3 + phase)
                         + 0.35 * self.rng.uniform(-1, 1))
            pts.append((bx + nx_ * off, by + ny_ * off))
        pts.append((x2, y2))
        self._add("poly", layer, (pts, False))

    def rect(self, x, y, w, h, layer, amp=0.11):
        c = [(x + self.rng.uniform(-amp, amp), y + self.rng.uniform(-amp, amp)),
             (x + w + self.rng.uniform(-amp, amp), y + self.rng.uniform(-amp, amp)),
             (x + w + self.rng.uniform(-amp, amp), y + h + self.rng.uniform(-amp, amp)),
             (x + self.rng.uniform(-amp, amp), y + h + self.rng.uniform(-amp, amp))]
        for i in range(4):
            a, b = c[i], c[(i + 1) % 4]
            self.line(a[0], a[1], b[0], b[1], layer, amp=amp * 0.75, step=2.6)

    def circle(self, cx, cy, r, layer, amp=0.06):
        ph1 = self.rng.uniform(0, 2 * math.pi)
        ph2 = self.rng.uniform(0, 2 * math.pi)
        pts = []
        for k in range(29):
            a = 2 * math.pi * k / 28
            rr = r + amp * math.sin(3 * a + ph1) + amp * 0.5 * math.sin(5 * a + ph2)
            pts.append((cx + rr * math.cos(a), cy + rr * math.sin(a)))
        self._add("poly", layer, (pts, True))

    def polyline(self, pts, layer, amp=0.0, step=3.0):
        if amp <= 0:
            self._add("poly", layer, (pts, False))
        else:
            for a, b in zip(pts[:-1], pts[1:]):
                self.line(a[0], a[1], b[0], b[1], layer, amp=amp, step=step)

    def obstacle(self, x, y, w, h):
        self.obstacles.append((x, y, x + w, y + h))


# ----------------------------------------------------------------------
# 2. КЛИППИНГ ТЕКСТУР
# ----------------------------------------------------------------------
def _subtract(lo, hi, blocks):
    for a, b in blocks:
        if b <= lo or a >= hi:
            continue
        if a <= lo and b >= hi:
            return []
        left = _subtract(lo, a, blocks) if a > lo else []
        right = _subtract(b, hi, blocks) if b < hi else []
        return left + right
    return [(lo, hi)]


def _h_blocks(p, y, extra):
    src = p.obstacles + (extra or [])
    return [(max(-1e9, o[0] - CLIP_MARGIN), min(1e9, o[2] + CLIP_MARGIN))
            for o in src if o[1] - CLIP_MARGIN < y < o[3] + CLIP_MARGIN]


def _v_blocks(p, x, extra):
    src = p.obstacles + (extra or [])
    return [(max(-1e9, o[1] - CLIP_MARGIN), min(1e9, o[3] + CLIP_MARGIN))
            for o in src if o[0] - CLIP_MARGIN < x < o[2] + CLIP_MARGIN]


def h_clip(p: Panel, y, xa, xb, layer, amp=0.06, extra=None):
    for a, b in _subtract(xa, xb, _h_blocks(p, y, extra)):
        if b - a >= 0.7:
            p.line(a, y, b, y, layer, amp=amp, step=2.2)


def v_clip(p: Panel, x, ya, yb, layer, amp=0.06, extra=None):
    for a, b in _subtract(ya, yb, _v_blocks(p, x, extra)):
        if b - a >= 0.7:
            p.line(x, a, x, b, layer, amp=amp, step=2.2)


def _point_free(p: Panel, x, y, margin=0.6):
    return not any(ox0 - margin < x < ox1 + margin and oy0 - margin < y < oy1 + margin
                   for ox0, oy0, ox1, oy1 in p.obstacles)


# ----------------------------------------------------------------------
# 3. ОКНО ПОВЫШЕННОЙ ДЕТАЛИЗАЦИИ
# ----------------------------------------------------------------------
def window(p: Panel, x, y, w, h, style):
    # внешняя и внутренняя рамы
    p.rect(x, y, w, h, L["02_PROYOMY"], amp=0.11)
    p.rect(x + 1.0, y + 1.0, w - 2.0, h - 2.0, L["02_PROYOMY"], amp=0.08)
    ix, iy, iw, ih = x + 1.0, y + 1.0, w - 2.0, h - 2.0

    if style == 2:                                   # трёхстворчатое
        m1, m2 = ix + iw / 3, ix + 2 * iw / 3
        for m in (m1, m2):
            p.line(m, iy, m, iy + ih, L["02_PROYOMY"], amp=0.07, step=2.0)
        sashes = [(ix, m1), (m1, m2), (m2, ix + iw)]
        fort = 1
    else:
        mr = 0.42 if style == 1 else 0.5
        m = ix + iw * mr
        p.line(m, iy, m, iy + ih, L["02_PROYOMY"], amp=0.07, step=2.0)
        sashes = [(ix, m), (m, ix + iw)]
        fort = 0 if style == 1 else 1
        if style == 3:
            fort = -1

    # собственные тонкие рамы каждой створки
    for ax, bx in sashes:
        p.rect(ax + 0.55, iy + 0.55, (bx - ax) - 1.1, ih - 1.1,
               L["02_PROYOMY"], amp=0.06)

    # форточка с собственным переплётом
    if fort >= 0:
        ax, bx = sashes[fort]
        fy = iy + ih * 0.66
        p.line(ax + 0.6, fy, bx - 0.6, fy, L["02_PROYOMY"], amp=0.06, step=1.6)
        p.line((ax + bx) / 2, fy + 0.5, (ax + bx) / 2, iy + ih - 0.6,
               L["02_PROYOMY"], amp=0.05, step=1.2)
        # ручка форточки
        p.line(bx - 1.0, fy - 1.1, bx - 1.0, fy - 0.35,
               L["02_PROYOMY"], amp=0.04, step=0.8)

    # ручки створок
    for ax, bx in sashes:
        hx = bx - 0.95
        p.line(hx, y + h * 0.42, hx, y + h * 0.56,
               L["02_PROYOMY"], amp=0.04, step=0.9)

    # отлив с капельником
    p.rect(x - 1.2, y - 1.25, w + 2.4, 0.75, L["02_PROYOMY"], amp=0.07)
    p.line(x - 0.6, y - 2.15, x + w + 0.6, y - 2.15,
           L["02_PROYOMY"], amp=0.06, step=2.0)
    p.obstacle(x - 1.6, y - 2.6, w + 3.2, h + 2.9)


# ----------------------------------------------------------------------
# 4. БАЛКОН
# ----------------------------------------------------------------------
def balcony(p: Panel, cx, yb):
    gw, gh = 25.0, 18.8
    gx, gy = cx - gw / 2, yb + 5.5
    # рамы остекления
    p.rect(gx, gy, gw, gh, L["03_BALKONY"], amp=0.11)
    p.rect(gx + 0.9, gy + 0.9, gw - 1.8, gh - 1.8, L["03_BALKONY"], amp=0.08)
    ix0, iy0, iw, ih = gx + 0.9, gy + 0.9, gw - 1.8, gh - 1.8
    # верхний ряд мелких стекол (форточный пояс)
    trans_h = 4.4
    ty = iy0 + ih - trans_h
    p.line(ix0, ty, ix0 + iw, ty, L["03_BALKONY"], amp=0.07, step=2.0)
    for k in range(1, 5):
        tx = ix0 + iw * k / 5
        p.line(tx, ty, tx, iy0 + ih, L["03_BALKONY"], amp=0.06, step=1.4)
    # нижние крупные створки
    for k in range(1, 3):
        mx = ix0 + iw * k / 3
        p.line(mx, iy0, mx, ty, L["03_BALKONY"], amp=0.07, step=2.0)
    # плита под остеклением (зазоры до парапета и рамы >= 0.6 мм)
    p.line(gx - 1.2, gy - 0.7, gx + gw + 1.2, gy - 0.7,
           L["03_BALKONY"], amp=0.08, step=2.4)
    p.obstacle(gx - 1.4, gy - 1.2, gw + 2.8, gh + 1.4)
    # парапет: плита + филёнка + решётка
    bw, bh = 28.5, 3.4
    bx, by = cx - bw / 2, yb + 0.8
    p.rect(bx, by, bw, bh, L["03_BALKONY"], amp=0.10)
    p.line(bx + 0.7, by + bh - 1.0, bx + bw - 0.7, by + bh - 1.0,
           L["03_BALKONY"], amp=0.06, step=1.8)
    x = bx + 1.3
    while x < bx + bw - 0.8:
        v_clip(p, x, by + 0.4, by + bh - 1.0, L["03_BALKONY"], amp=0.05)
        x += 1.7
    p.obstacle(bx - 0.4, by - 0.4, bw + 0.8, bh + 0.8)


# ----------------------------------------------------------------------
# 5. СПЛОШНОЕ ОСТЕКЛЕНИЕ
# ----------------------------------------------------------------------
def glazed_strip(p: Panel, x0, x1, yb):
    fx0, fx1 = x0 + 1.8, x1 - 1.8
    fy0, fy1 = yb + 4.8, yb + 23.4
    p.rect(fx0, fy0, fx1 - fx0, fy1 - fy0, L["03_BALKONY"], amp=0.10)
    p.rect(fx0 + 0.9, fy0 + 0.9, fx1 - fx0 - 1.8, fy1 - fy0 - 1.8,
           L["03_BALKONY"], amp=0.07)
    iw = fx1 - fx0 - 1.8
    ix0, iy0 = fx0 + 0.9, fy0 + 0.9
    ty = iy0 + (fy1 - fy0 - 1.8) - 4.2
    p.line(ix0, ty, ix0 + iw, ty, L["03_BALKONY"], amp=0.06, step=2.0)
    for k in range(1, 4):
        tx = ix0 + iw * k / 4
        p.line(tx, ty, tx, iy0 + (fy1 - fy0 - 1.8),
               L["03_BALKONY"], amp=0.05, step=1.4)
    for k in range(1, 3):
        mx = ix0 + iw * k / 3
        p.line(mx, iy0, mx, ty, L["03_BALKONY"], amp=0.06, step=2.0)
    p.obstacle(fx0 - 0.4, fy0 - 1.0, (fx1 - fx0) + 0.8, (fy1 - fy0) + 1.2)
    # бетонный фартук с опалубочными линиями
    ax0, ay0 = x0 + 1.8, yb + 1.2
    aw, ah = (x1 - x0) - 3.6, 2.9
    p.rect(ax0, ay0, aw, ah, L["03_BALKONY"], amp=0.09)
    p.line(ax0 + 0.5, ay0 + ah * 0.5, ax0 + aw - 0.5, ay0 + ah * 0.5,
           L["03_BALKONY"], amp=0.06, step=1.8)
    p.obstacle(ax0 - 0.4, ay0 - 0.4, aw + 0.8, ah + 0.8)


# ----------------------------------------------------------------------
# 6. ЛОДЖИЯ И ПОДЪЕЗД
# ----------------------------------------------------------------------
def loggia(p: Panel, x0, x1, yb):
    # двойная линия боковой стенки
    p.rect(x1 - 2.3, yb + 0.6, 0.85, PH - 1.4, L["03_BALKONY"], amp=0.08)
    # дверь с фрамугой
    dx, dw, dh = x0 + 2.2, 8.6, 16.8
    dy = yb + 5.6
    p.rect(dx, dy, dw, dh, L["03_BALKONY"], amp=0.10)
    p.rect(dx + 0.75, dy + 0.75, dw - 1.5, dh - 1.5, L["03_BALKONY"], amp=0.07)
    p.line(dx + dw / 2, dy + 0.75, dx + dw / 2, dy + dh - 0.75,
           L["03_BALKONY"], amp=0.06, step=2.0)
    p.line(dx + 0.75, dy + dh - 3.6, dx + dw - 0.75, dy + dh - 3.6,
           L["03_BALKONY"], amp=0.06, step=1.6)
    p.line(dx + dw - 1.5, dy + dh * 0.4, dx + dw - 1.5, dy + dh * 0.55,
           L["03_BALKONY"], amp=0.04, step=1.0)
    # ограждение: две горизонтали + стойки
    for ry in (yb + 3.4, yb + 1.9):
        p.line(x0 + 1.3, ry, x1 - 2.6, ry, L["03_BALKONY"], amp=0.07, step=2.2)
    rx = x0 + 13.0
    while rx < x1 - 3.2:
        v_clip(p, rx, yb + 1.6, yb + 3.7, L["03_BALKONY"], amp=0.05)
        rx += 1.8
    p.obstacle(x0 + 0.3, yb + 0.3, (x1 - x0) - 0.6, PH - 0.6)


def entrance(p: Panel, x0, x1, yb):
    dx, dw, dh = x0 + 3.5, 12.5, 19.5
    dy = yb + 2.0
    p.rect(dx, dy, dw, dh, L["02_PROYOMY"], amp=0.11)
    p.rect(dx + 0.8, dy + 0.8, dw - 1.6, dh - 1.6, L["02_PROYOMY"], amp=0.08)
    # фрамуга
    p.line(dx + 0.8, dy + dh - 4.0, dx + dw - 0.8, dy + dh - 4.0,
           L["02_PROYOMY"], amp=0.06, step=1.8)
    p.line(dx + dw / 2, dy + 0.8, dx + dw / 2, dy + dh - 4.0,
           L["02_PROYOMY"], amp=0.06, step=2.0)
    # ручки двустворчатой двери
    for hx in (dx + dw / 2 - 0.9, dx + dw / 2 + 0.9):
        p.line(hx, dy + dh * 0.42, hx, dy + dh * 0.56,
               L["02_PROYOMY"], amp=0.04, step=0.9)
    # козырёк (две линии с боковыми упорами)
    p.line(dx - 2.4, dy + dh + 1.0, dx + dw + 2.4, dy + dh + 1.0,
           L["02_PROYOMY"], amp=0.08, step=2.2)
    p.line(dx - 1.6, dy + dh + 2.1, dx + dw + 1.6, dy + dh + 2.1,
           L["02_PROYOMY"], amp=0.08, step=2.2)
    p.line(dx - 1.6, dy + dh + 1.0, dx - 1.6, dy + dh + 2.1,
           L["02_PROYOMY"], amp=0.05, step=1.0)
    p.line(dx + dw + 1.6, dy + dh + 1.0, dx + dw + 1.6, dy + dh + 2.1,
           L["02_PROYOMY"], amp=0.05, step=1.0)
    # ступени (зазор до межэтажного шва >= 1 мм)
    p.line(dx - 2.0, dy - 0.5, dx + dw + 2.0, dy - 0.5,
           L["02_PROYOMY"], amp=0.07, step=2.0)
    p.line(dx - 3.3, dy - 1.05, dx + dw + 3.3, dy - 1.05,
           L["02_PROYOMY"], amp=0.07, step=2.0)
    p.obstacle(x0 + 0.3, yb + 0.3, (x1 - x0) - 0.6, PH - 0.6)


# ----------------------------------------------------------------------
# 7. КОНДИЦИОНЕР И ВОДОСТОК
# ----------------------------------------------------------------------
def ac_unit(p: Panel, cx, top_y, h=2.8, w=4.6):
    x, y = cx - w / 2, top_y - h
    p.rect(x, y, w, h, L["05_KONDICIONERY"], amp=0.08)
    fcx, fcy = x + 1.4, y + h / 2
    p.circle(fcx, fcy, 0.7, L["05_KONDICIONERY"], amp=0.045)
    p.circle(fcx, fcy, 0.3, L["05_KONDICIONERY"], amp=0.035)
    # лопатки вентилятора (не из центра — без схождения линий)
    for k in range(3):
        a = k * 2 * math.pi / 3 + 0.4
        p.line(fcx + 0.32 * math.cos(a), fcy + 0.32 * math.sin(a),
               fcx + 0.55 * math.cos(a), fcy + 0.55 * math.sin(a),
               L["05_KONDICIONERY"], amp=0.025, step=0.6)
    # жалюзи (3 планки, зазоры до вентилятора и корпуса >= 0.45 мм)
    for i in range(3):
        gx = x + 2.55 + i * 0.65
        p.line(gx, y + 0.5, gx, y + h - 0.5, L["05_KONDICIONERY"],
               amp=0.04, step=0.9)
    # кронштейны
    p.line(x + 0.6, y, x + 0.6, y - 0.85, L["05_KONDICIONERY"], amp=0.04, step=0.8)
    p.line(x + w - 0.6, y, x + w - 0.6, y - 0.85,
           L["05_KONDICIONERY"], amp=0.04, step=0.8)
    # гибкая подводка
    px = x + 0.9
    pts = [(px, top_y), (px, top_y + 1.0)]
    r = 0.9
    for k in range(1, 7):
        a = -math.pi / 2 * k / 6
        pts.append((px + r + r * math.cos(a + math.pi / 2),
                    top_y + 1.0 + r * math.sin(a + math.pi / 2)))
    pts.append((px + r, top_y + 2.3))
    p.polyline(pts, L["05_KONDICIONERY"], amp=0.06, step=1.1)
    p.obstacle(x - 0.4, y - 1.1, w + 0.8, h + 1.3)


def downpipe(p: Panel, x, y_lo, y_hi):
    gap = 0.85
    p.line(x - gap / 2, y_lo, x - gap / 2, y_hi, L["06_VODOSTOK"],
           amp=0.09, step=4.0)
    p.line(x + gap / 2, y_lo, x + gap / 2, y_hi, L["06_VODOSTOK"],
           amp=0.09, step=4.0)
    # хомуты на каждом этаже
    y = y_lo + 2.5
    while y < y_hi - 1:
        p.line(x - 0.85, y, x + 0.85, y, L["06_VODOSTOK"], amp=0.05, step=1.1)
        y += PH
    # воронка вверху
    p.polyline([(x - 1.5, y_hi + 1.6), (x - gap / 2, y_hi)],
               L["06_VODOSTOK"], amp=0.06, step=1.2)
    p.polyline([(x + 1.5, y_hi + 1.6), (x + gap / 2, y_hi)],
               L["06_VODOSTOK"], amp=0.06, step=1.2)
    p.line(x - 1.5, y_hi + 1.6, x + 1.5, y_hi + 1.6,
           L["06_VODOSTOK"], amp=0.06, step=1.2)
    # отметка внизу (выпуск)
    p.line(x - 0.7, y_lo, x - 0.7, y_lo - 1.1, L["06_VODOSTOK"], amp=0.05, step=1.0)
    p.line(x + 0.7, y_lo, x + 0.7, y_lo - 1.1, L["06_VODOSTOK"], amp=0.05, step=1.0)


def crack(p: Panel, x, y, x_lo, x_hi, direction=1):
    pts = [(x, y)]
    cx, cy = x, y
    for _ in range(p.rng.randint(3, 5)):
        nx = min(x_hi, max(x_lo, cx + direction * p.rng.uniform(1.4, 2.4)))
        cy += p.rng.choice((-1, 1)) * p.rng.uniform(0.7, 1.3)
        pts.append((nx, cy))
        cx = nx
    if all(_point_free(p, px, py) for px, py in pts):
        p.polyline(pts, L["04_KLADKA"], amp=0.0)


# ----------------------------------------------------------------------
# 8. РАСКЛАДКА СЕКЦИЙ
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
# 9. СБОРКА ПАНЕЛИ
# ----------------------------------------------------------------------
def build_panel(seed):
    p = Panel(seed)
    strips = make_strips()
    rng = p.rng

    ac_map = {}
    for si, fl, pos in rng.sample(AC_CANDIDATES, 6):
        ac_map[(si, fl)] = pos

    # ---- 9.1 регистрация проёмов ---------------------------------------
    for f in range(FLOORS):
        yb = PLINTH_TOP + f * PH
        for si, (typ, x0, x1) in enumerate(strips):
            cx = (x0 + x1) / 2
            if typ == "single":
                p.obstacle(cx - 7.0 - 1.6, yb + 8.5 - 2.6, 14.0 + 3.2, 15.5 + 2.9)
            elif typ == "double":
                for wx in (x0 + 1.9, x0 + 18.4):
                    p.obstacle(wx - 1.6, yb + 8.5 - 2.6, 13.2 + 3.2, 15.5 + 2.9)
            elif typ == "balcony":
                p.obstacle(cx - 13.9, yb + 3.6, 27.8, 20.4)
                p.obstacle(cx - 14.7, yb + 0.5, 29.4, 4.4)
            elif typ == "glazed":
                p.obstacle(x0 + 1.4, yb + 3.8, (x1 - x0) - 2.8, 20.2)
            elif typ == "loggia":
                p.obstacle(x0 + 0.3, yb + 0.3, (x1 - x0) - 0.6, PH - 0.6)
            if (si, f) in ac_map and typ == "double":
                pos = ac_map[(si, f)]
                acx = x0 + 8.2 if pos == "L" else x0 + 25.3
                p.obstacle(acx - 2.7, yb + 1.0, 5.4, 3.6)

    # водостоки — сквозные препятствия для текстур и трещин
    for si in (1, 3):
        jx = (strips[si][2] + strips[si + 1][1]) / 2 + DP_OFFSET
        p.obstacle(jx - 1.15, PLINTH_TOP - 0.5, 2.3, ROOF_BOT - PLINTH_TOP + 2.0)

    # ---- 9.2 каркас ------------------------------------------------------
    p.line(0, 0, W, 0, L["01_KARKAS"], amp=0.0)
    p.line(W, 0, W, H, L["01_KARKAS"], amp=0.0)
    p.line(W, H, 0, H, L["01_KARKAS"], amp=0.0)
    p.line(0, H, 0, 0, L["01_KARKAS"], amp=0.0)
    for f in range(FLOORS + 1):
        y = PLINTH_TOP + f * PH
        p.line(X0, y, X1, y, L["01_KARKAS"], amp=0.14)
    for i in range(len(strips) - 1):
        x = (strips[i][2] + strips[i + 1][1]) / 2
        p.line(x, MARGIN, x, H - MARGIN, L["01_KARKAS"], amp=0.14)

    # ---- 9.3 цоколь на фундаменте ---------------------------------------
    p.line(X0 + 0.8, 6.2, X1 - 0.8, 6.2, L["07_TSEKOL"], amp=0.11)
    for yy in (9.8, 13.4, 16.9):
        p.line(X0 + 0.8, yy, X1 - 0.8, yy, L["07_TSEKOL"], amp=0.10)
    for vx in (10.0, 86.0):
        p.rect(vx, 10.6, 8.0, 3.6, L["07_TSEKOL"], amp=0.07)
        for k in range(1, 5):
            p.line(vx + k * 1.35, 10.9, vx + k * 1.35, 13.9,
                   L["07_TSEKOL"], amp=0.04, step=0.9)
        p.obstacle(vx - 0.3, 10.3, 8.6, 4.2)

    # ---- 9.4 жилая зона --------------------------------------------------
    for f in range(FLOORS):
        yb = PLINTH_TOP + f * PH
        for si, (typ, x0, x1) in enumerate(strips):
            cx = (x0 + x1) / 2
            style = (f * 7 + si * 3 + seed) % 4
            if typ == "single":
                window(p, cx - 7.0, yb + 8.5, 14.0, 15.5, style)
            elif typ == "double":
                window(p, x0 + 1.9, yb + 8.5, 13.2, 15.5, style)
                window(p, x0 + 18.4, yb + 8.5, 13.2, 15.5, (style + 1) % 4)
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
                acx = x0 + 8.2 if pos == "L" else x0 + 25.3
                ac_unit(p, acx, yb + 4.0)

    # ---- 9.5 кирпичные подоконные пояса ---------------------------------
    # блокеры поясов — только для квадратной сетки (9.6), не для самой кладки
    band_blockers = []
    for f in range(FLOORS):
        yb = PLINTH_TOP + f * PH
        for si, (typ, x0, x1) in enumerate(strips):
            if typ != "double":
                continue
            xa, xb = x0 + 1.0, x1 - 1.0
            ya, ybnd = yb + 1.1, yb + 5.6
            band_blockers.append((xa - 0.2, ya - 0.2, xb + 0.2, ybnd + 0.2))
            row = 0
            y = math.ceil(ya / BRICK_PITCH) * BRICK_PITCH
            while y <= ybnd:
                h_clip(p, y, xa, xb, L["04_KLADKA"], amp=0.05)
                stagger = BRICK_LEN / 2 if row % 2 else 0.0
                x = math.ceil((xa - stagger) / BRICK_LEN) * BRICK_LEN + stagger
                while x <= xb:
                    v_clip(p, x, y + 0.1, min(y + BRICK_PITCH - 0.1, ybnd),
                           L["04_KLADKA"], amp=0.04)
                    x += BRICK_LEN
                y += BRICK_PITCH
                row += 1

    # ---- 9.6 мелкая квадратная плитка стен ------------------------------
    for f in range(FLOORS):
        ya = PLINTH_TOP + f * PH + 0.45
        ybnd = PLINTH_TOP + (f + 1) * PH - 0.45
        for si, (typ, x0, x1) in enumerate(strips):
            if typ == "loggia":
                continue
            xa, xb = x0 + 0.5, x1 - 0.5
            x = math.ceil(xa / GRID_PITCH) * GRID_PITCH
            while x <= xb:
                v_clip(p, x, ya, ybnd, L["04_KLADKA"], amp=0.045,
                       extra=band_blockers)
                x += GRID_PITCH
            y = math.ceil(ya / GRID_PITCH) * GRID_PITCH
            while y <= ybnd:
                h_clip(p, y, xa, xb, L["04_KLADKA"], amp=0.045,
                       extra=band_blockers)
                y += GRID_PITCH

    # трещины
    brick_strips = [s for s in strips if s[0] != "loggia"]
    made = 0
    for _try in range(50):
        if made >= rng.randint(2, 3):
            break
        typ, x0, x1 = rng.choice(brick_strips)
        f = rng.randrange(FLOORS)
        sx = rng.uniform(x0 + 3.5, x1 - 7)
        sy = PLINTH_TOP + f * PH + rng.uniform(4, PH - 4)
        if not _point_free(p, sx, sy):
            continue
        before = len(p.ents)
        crack(p, sx, sy, x0 + 1.6, x1 - 1.6, direction=rng.choice((1, -1)))
        if len(p.ents) > before:
            made += 1

    # ---- 9.7 водостоки ---------------------------------------------------
    for si in (1, 3):
        jx = (strips[si][2] + strips[si + 1][1]) / 2
        downpipe(p, jx + DP_OFFSET, PLINTH_TOP + 0.6, ROOF_BOT - 0.6)

    # ---- 9.8 техэтаж и плоская крыша ------------------------------------
    # малые окна техэтажа с двойной рамой
    for si in (0, 2, 3):
        typ, x0, x1 = strips[si]
        cx = (x0 + x1) / 2
        tw = 7.5 if typ != "glazed" else 9.0
        tx, ty = cx - tw / 2, ROOF_BOT + 4.0
        p.rect(tx, ty, tw, 5.0, L["08_KRYSHA"], amp=0.08)
        p.rect(tx + 0.7, ty + 0.7, tw - 1.4, 3.6, L["08_KRYSHA"], amp=0.06)
        p.line(tx + tw / 2, ty + 0.7, tx + tw / 2, ty + 4.3,
               L["08_KRYSHA"], amp=0.05, step=1.2)
        p.obstacle(tx - 0.4, ty - 0.4, tw + 0.8, 5.8)
    # парапет: три линии (тело, шов, крышка)
    for yy in (ROOF_BOT + 12.5, ROOF_BOT + 21.0, ROOF_BOT + 27.0):
        p.line(X0 + 0.5, yy, X1 - 0.5, yy, L["08_KRYSHA"], amp=0.10)
    # вентшахты с крышками и прорезями (крышка выше линии парапета)
    for vx, vw in ((34.0, 10.0), (98.0, 13.0)):
        p.rect(vx, ROOF_BOT + 22.2, vw, 4.2, L["08_KRYSHA"], amp=0.08)
        p.line(vx + 1.0, ROOF_BOT + 24.3, vx + vw - 1.0, ROOF_BOT + 24.3,
               L["08_KRYSHA"], amp=0.05, step=1.2)
        p.line(vx - 0.8, ROOF_BOT + 27.7, vx + vw + 0.8, ROOF_BOT + 27.7,
               L["08_KRYSHA"], amp=0.06, step=1.4)
    # антенна
    if seed % 2 == 1:
        ax_ = 142.0
        p.line(ax_, ROOF_BOT + 27.2, ax_, ROOF_BOT + 32.0,
               L["08_KRYSHA"], amp=0.05, step=1.4)
        for yy, half in ((ROOF_BOT + 31.6, 1.7), (ROOF_BOT + 30.2, 1.3),
                         (ROOF_BOT + 28.8, 1.0)):
            p.line(ax_ - half, yy, ax_ + half, yy, L["08_KRYSHA"],
                   amp=0.04, step=0.9)

    return p


# ----------------------------------------------------------------------
# 10. АУДИТ (пересечения разрешены; контроль параллельных слияний)
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


def _point_seg_dist(pt, a, b):
    ux, uy = b[0] - a[0], b[1] - a[1]
    nu2 = ux * ux + uy * uy
    if nu2 < 1e-14:
        return math.hypot(pt[0] - a[0], pt[1] - a[1])
    t = _clamp(((pt[0] - a[0]) * ux + (pt[1] - a[1]) * uy) / nu2)
    return math.hypot(pt[0] - (a[0] + t * ux), pt[1] - (a[1] + t * uy))


def _seg_dist(a, b, c, d):
    if math.hypot(b[0] - a[0], b[1] - a[1]) < 1e-12:
        return _point_seg_dist(a, c, d)
    if math.hypot(d[0] - c[0], d[1] - c[1]) < 1e-12:
        return _point_seg_dist(c, a, b)
    ux, uy = b[0] - a[0], b[1] - a[1]
    vx, vy = d[0] - c[0], d[1] - c[1]
    wx, wy = a[0] - c[0], a[1] - c[1]
    A, B, C = ux * ux + uy * uy, ux * vx + uy * vy, vx * vx + vy * vy
    D, E = ux * wx + uy * wy, vx * wx + vy * wy
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
                    continue
                if _overlap(a, b, c, d) <= 0.5:
                    continue
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
# 11. ЭКСПОРТ DXF
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
            msp.add_lwpolyline([(x + ox, y + oy) for x, y in pts],
                               close=closed, dxfattribs={"layer": layer})

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
# 12. ЭКСПОРТ EPS
# ----------------------------------------------------------------------
def save_eps(panels, path):
    gap = 15.0
    sheet_w, sheet_h = 2 * W + gap, 2 * H + gap
    PT = 72.0 / 25.4
    out = [
        "%!PS-Adobe-3.0 EPSF-3.0",
        "%%Title: Panel facade 4 sides 165x320mm (portrait)",
        "%%Creator: generate_facade_v3.py",
        f"%%BoundingBox: 0 0 {sheet_w * PT:.2f} {sheet_h * PT:.2f}",
        "%%EndComments",
        f"{PT:.6f} {PT:.6f} scale",
        "0 setlinecap 1 setlinejoin",
    ]
    for i, (name, aci, (r, g, b)) in enumerate(LAYERS, start=1):
        out.append(f"/C{i} {{ {r/255:.3f} {g/255:.3f} {b/255:.3f} "
                   f"setcolor 0.07 setlinewidth }} bind def")
    proc = {name: f"C{i}" for i, (name, _, _) in enumerate(LAYERS, 1)}
    for i, p in enumerate(panels):
        col, row = i % 2, i // 2
        ox, oy = col * (W + gap), (1 - row) * (H + gap)
        out.append(f"%% side {i + 1}")
        out.append(f"{ox:.3f} {oy:.3f} translate")
        for kind, layer, d in p.ents:
            pts, closed = d
            out.append(proc[layer])
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
# 13. ЗАПУСК
# ----------------------------------------------------------------------
def main():
    out_dir = os.path.join(os.path.dirname(__file__), "..", "output")
    os.makedirs(out_dir, exist_ok=True)
    panels = [build_panel(seed) for seed in (101, 202, 303, 404)]
    for i, p in enumerate(panels, 1):
        audit(p, f"side {i}")
    save_eps(panels, os.path.join(out_dir, "facade_4sides_165x320.eps"))
    save_dxf(panels, os.path.join(out_dir, "facade_4sides_165x320.dxf"))
    for i, p in enumerate(panels, 1):
        save_dxf([p], os.path.join(out_dir, f"facade_side{i}_165x320.dxf"),
                 per_side=True)


if __name__ == "__main__":
    main()
