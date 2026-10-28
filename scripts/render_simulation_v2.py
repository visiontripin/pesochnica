# -*- coding: utf-8 -*-
"""
Превью v2:
  1) layout_colors.png  — 4 стороны, 8 цветов по классам объектов;
  2) simulation_side.png — сторона 1, имитация луча 0.2 мм на ЛМДФ (1:1);
  3) simulation_zoom.png — увеличенный фрагмент с кладкой и окном.
Читает итоговые DXF из output/.
"""
import math
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection
import ezdxf

HERE = os.path.dirname(__file__)
OUT = os.path.join(HERE, "..", "output")
BEAM_MM = 0.2
WOOD = "#e9ddc6"

RGB = {
    "01_KARKAS": (214 / 255, 39 / 255, 40 / 255),
    "02_PROYOMY": (31 / 255, 95 / 255, 214 / 255),
    "03_BALKONY": (44 / 255, 160 / 255, 44 / 255),
    "04_KLADKA": (140 / 255, 90 / 255, 43 / 255),
    "05_KONDICIONERY": (197 / 255, 18 / 255, 197 / 255),
    "06_VODOSTOK": (15 / 255, 154 / 255, 160 / 255),
    "07_TSEKOL": (230 / 255, 138 / 255, 0 / 255),
    "08_KRYSHA": (68 / 255, 68 / 255, 68 / 255),
}
ENGRAVE = {k: "#2b2118" for k in RGB}      # на ЛМДФ все классы гравируются одинаково


def load(path):
    doc = ezdxf.readfile(path)
    by = {k: [] for k in RGB}
    for e in doc.modelspace():
        layer = e.dxf.layer
        if layer not in by:
            continue
        t = e.dxftype()
        if t == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            if e.closed:
                pts.append(pts[0])
            for a, b in zip(pts[:-1], pts[1:]):
                by[layer].append([a, b])
        elif t == "LINE":
            by[layer].append([(e.dxf.start.x, e.dxf.start.y),
                              (e.dxf.end.x, e.dxf.end.y)])
        elif t == "CIRCLE":
            c, r = e.dxf.center, e.dxf.radius
            pts = [(c.x + r * math.cos(2 * math.pi * k / 40),
                    c.y + r * math.sin(2 * math.pi * k / 40)) for k in range(41)]
            for a, b in zip(pts[:-1], pts[1:]):
                by[layer].append([a, b])
    return by


def render(by, colors, path, xlim, ylim, fig_w_mm, bg="white", beam=True):
    span_x = xlim[1] - xlim[0]
    fig_w_in = fig_w_mm / 25.4
    fig_h_in = fig_w_in * (ylim[1] - ylim[0]) / span_x
    fig = plt.figure(figsize=(fig_w_in, fig_h_in), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    fig.patch.set_facecolor(bg)
    ax.set_facecolor(bg)
    lw_pt = (BEAM_MM / span_x * fig_w_in * 72.0) if beam else 0.9
    for layer, segs in by.items():
        if segs:
            ax.add_collection(LineCollection(
                segs, colors=[colors[layer]] * len(segs),
                linewidths=lw_pt, capstyle="round", joinstyle="round"))
    ax.set_xlim(*xlim); ax.set_ylim(*ylim)
    ax.set_aspect("equal"); ax.set_xticks([]); ax.set_yticks([])
    for sp in ax.spines.values():
        sp.set_visible(False)
    fig.savefig(os.path.join(OUT, path), facecolor=bg)
    plt.close(fig)
    print("saved", path)


# цветной слой-превью всего листа 2x2
by_sheet = load(os.path.join(OUT, "facade_4sides_320x165.dxf"))
render(by_sheet, RGB, "layout_colors.png", (0, 655), (0, 345), 330.0,
       bg="white", beam=False)

# симуляция гравировки на стороне 1
by_side = load(os.path.join(OUT, "facade_side1_320x165.dxf"))
render(by_side, ENGRAVE, "simulation_side.png", (0, 320), (0, 165), 280.0,
       bg=WOOD, beam=True)
render(by_side, ENGRAVE, "simulation_zoom.png", (0, 118), (55, 118), 240.0,
       bg=WOOD, beam=True)
