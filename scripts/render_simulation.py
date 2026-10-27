# -*- coding: utf-8 -*-
"""
Рендер итогового DXF с имитацией лазерной гравировки:
ширина штриха = 0.2 мм (толщина луча), фон под цвет ЛМДФ.
Читает именно output/facade_panel_320x200.dxf — то, что уйдёт на станок.
"""
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
LAYER_STYLE = {
    "KONTUR": "#2b2118",   # глубокий проход — тёмный
    "DETALI": "#7a6a55",   # лёгкий проход — светлее
}


def load_segments():
    doc = ezdxf.readfile(os.path.join(OUT, "facade_panel_320x200.dxf"))
    msp = doc.modelspace()
    by_layer = {"KONTUR": [], "DETALI": []}
    for e in msp:
        t = e.dxftype()
        layer = e.dxf.layer
        if layer not in by_layer:
            continue
        if t == "LINE":
            by_layer[layer].append([(e.dxf.start.x, e.dxf.start.y),
                                    (e.dxf.end.x, e.dxf.end.y)])
        elif t == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            if e.closed:
                pts.append(pts[0])
            for a, b in zip(pts[:-1], pts[1:]):
                by_layer[layer].append([a, b])
        elif t == "CIRCLE":
            import math
            c, r = e.dxf.center, e.dxf.radius
            pts = [(c.x + r * math.cos(2 * math.pi * k / 40),
                    c.y + r * math.sin(2 * math.pi * k / 40)) for k in range(41)]
            for a, b in zip(pts[:-1], pts[1:]):
                by_layer[layer].append([a, b])
    return by_layer


def render(by_layer, path, xlim, ylim, fig_w_mm, dpi=300):
    span_x = xlim[1] - xlim[0]
    span_y = ylim[1] - ylim[0]
    fig_w_in = fig_w_mm / 25.4
    fig_h_in = fig_w_in * span_y / span_x
    fig = plt.figure(figsize=(fig_w_in, fig_h_in), dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])              # оси на всю фигуру — масштаб точный
    fig.patch.set_facecolor(WOOD)
    ax.set_facecolor(WOOD)
    # ширина штриха строго 0.2 мм в масштабе геометрии
    lw_pt = BEAM_MM / span_x * fig_w_in * 72.0
    for layer, col in LAYER_STYLE.items():
        ax.add_collection(LineCollection(by_layer[layer], colors=col,
                                         linewidths=lw_pt, capstyle="round",
                                         joinstyle="round"))
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_aspect("equal")
    ax.set_xticks([]); ax.set_yticks([])
    for sp in ax.spines.values():
        sp.set_visible(False)
    fig.savefig(os.path.join(OUT, path), facecolor=WOOD)
    plt.close(fig)
    print("saved", path, f"(beam {lw_pt:.2f} pt)")


by_layer = load_segments()
render(by_layer, "simulation_full.png", (0, 320), (0, 200), 260.0)
# зум: две полосы балконов и пары окон, 2 этажа
render(by_layer, "simulation_zoom.png", (30, 150), (56, 116), 220.0)
