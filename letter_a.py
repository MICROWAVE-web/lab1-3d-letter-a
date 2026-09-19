"""Каркас объёмной буквы «А» (вершины + рёбра)."""

from __future__ import annotations

from typing import List, Tuple

Vec3 = Tuple[float, float, float]
Edge = Tuple[int, int]


def build_letter_a(
    width: float = 120.0,
    height: float = 160.0,
    depth: float = 100.0,
    leg: float = 28.0,
    bar_bottom: float = 55.0,
    bar_height: float = 22.0,
    hole_top: float = 125.0,
) -> Tuple[List[Vec3], List[Edge]]:
    """Объёмная буква А в стиле эскиза: две ножки, перекладина, отверстие сверху.

    Модель — призма с вырезами. Центр около начала координат.
    """
    w, h, d = width, height, depth
    cx, cz = w / 2.0, h / 2.0
    y0, y1 = -d / 2.0, d / 2.0
    bar_top = bar_bottom + bar_height

    # Передняя грань (локальные координаты, затем центрирование)
    # Нижний контур (ножки + низ перекладины как потолок выреза):
    # 0--1                 4--5
    # |  |                 |  |
    # |  2-----------------3  |
    # |                       |
    # |  11----10   9----8    |   отверстие над перекладиной
    # |  |                |   |
    # 7-----------------------6
    front = [
        (0.0, y0, 0.0),                 # 0  низ левой ножки снаружи
        (leg, y0, 0.0),                 # 1  низ левой ножки внутри
        (leg, y0, bar_bottom),          # 2  низ перекладины слева
        (w - leg, y0, bar_bottom),      # 3  низ перекладины справа
        (w - leg, y0, 0.0),             # 4  низ правой ножки внутри
        (w, y0, 0.0),                   # 5  низ правой ножки снаружи
        (w, y0, h),                     # 6  верх справа
        (0.0, y0, h),                   # 7  верх слева
        (w - leg, y0, bar_top),         # 8  верх перекладины справа
        (w - leg, y0, hole_top),        # 9  верх отверстия справа
        (leg, y0, hole_top),            # 10 верх отверстия слева
        (leg, y0, bar_top),             # 11 верх перекладины слева
    ]

    back = [(x, y1, z) for x, _, z in front]

    vertices: List[Vec3] = []
    for x, y, z in front + back:
        vertices.append((x - cx, y, z - cz))

    n = len(front)
    # Контур передней грани (включая отверстие и перекладину)
    # Внешний контур арки + отверстие над перекладиной + бока перекладины
    front_loop = [
        (0, 1), (1, 2), (2, 3), (3, 4), (4, 5), (5, 6), (6, 7), (7, 0),
        (2, 11), (3, 8),           # боковые рёбра перекладины
        (11, 8),                   # верх перекладины
        (11, 10), (10, 9), (9, 8), # отверстие («дырка» буквы А)
    ]
    edges: List[Edge] = list(front_loop)

    edges.extend((a + n, b + n) for a, b in front_loop)

    for i in range(n):
        edges.append((i, i + n))

    return vertices, edges


def build_axes(length: float = 180.0) -> Tuple[List[Vec3], List[Edge], List[str]]:
    """Оси координат для ориентира (как на эскизе)."""
    verts: List[Vec3] = [
        (0.0, 0.0, 0.0),
        (length, 0.0, 0.0),
        (0.0, length, 0.0),
        (0.0, 0.0, length),
    ]
    edges: List[Edge] = [(0, 1), (0, 2), (0, 3)]
    labels = ["X", "Y", "Z"]
    return verts, edges, labels
