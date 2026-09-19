#!/usr/bin/env python3
"""
Лабораторная работа 1.
Проекции 3D-объектов и аффинные преобразования.
Объект: каркасная буква «А». Матрицы/проекции — вручную (math3d).
Отрисовка: собственный буфер кадра → PPM → PhotoImage"""

from __future__ import annotations

import math
import tkinter as tk
from typing import List, Optional, Tuple

from letter_a import build_axes, build_letter_a
from math3d import (
    Mat4,
    identity,
    mat_mul,
    perspective_project,
    rotate_x,
    rotate_y,
    rotate_z,
    scale,
    transform_points,
    translate,
)

Vec3 = Tuple[float, float, float]
RGB = Tuple[int, int, int]


def hex_rgb(color: str) -> RGB:
    c = color.lstrip("#")
    return int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)


class SoftCanvas:
    """RGB-буфер + выгрузка в tk.PhotoImage через PPM."""

    def __init__(self, width: int, height: int, bg: str) -> None:
        self.bg = hex_rgb(bg)
        self.width = 0
        self.height = 0
        self.buf = bytearray()
        self.image: Optional[tk.PhotoImage] = None
        self.resize(width, height)

    def resize(self, width: int, height: int) -> None:
        width = max(160, int(width))
        height = max(120, int(height))
        if width == self.width and height == self.height and self.buf:
            return
        self.width = width
        self.height = height
        self.buf = bytearray(width * height * 3)
        self.clear()

    def clear(self) -> None:
        r, g, b = self.bg
        # быстрая заливка
        pixel = bytes((r, g, b))
        self.buf[:] = pixel * (self.width * self.height)

    def point(self, x: int, y: int, rgb: RGB) -> None:
        if 0 <= x < self.width and 0 <= y < self.height:
            i = (y * self.width + x) * 3
            self.buf[i] = rgb[0]
            self.buf[i + 1] = rgb[1]
            self.buf[i + 2] = rgb[2]

    def line(self, x0: float, y0: float, x1: float, y1: float, color: str) -> None:
        rgb = hex_rgb(color)
        x0i, y0i = int(round(x0)), int(round(y0))
        x1i, y1i = int(round(x1)), int(round(y1))
        dx = abs(x1i - x0i)
        dy = -abs(y1i - y0i)
        sx = 1 if x0i < x1i else -1
        sy = 1 if y0i < y1i else -1
        err = dx + dy
        x, y = x0i, y0i
        while True:
            self.point(x, y, rgb)
            self.point(x + 1, y, rgb)
            self.point(x, y + 1, rgb)
            if x == x1i and y == y1i:
                break
            e2 = 2 * err
            if e2 >= dy:
                err += dy
                x += sx
            if e2 <= dx:
                err += dx
                y += sy

    def flush(self) -> tk.PhotoImage:
        header = f"P6 {self.width} {self.height} 255\n".encode("ascii")
        # Tk на macOS принимает сырой PPM в data= (без base64)
        self.image = tk.PhotoImage(data=header + self.buf)
        return self.image


class LetterAApp:
    BG = "#f4efe6"
    LINE = "#b91c1c"
    AXIS = ("#dc2626", "#16a34a", "#2563eb")
    TEXT = "#111827"

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ЛР1 — Объёмная буква А | Аффинные преобразования")
        self.root.geometry("960x700")
        self.root.minsize(640, 480)
        self.root.configure(bg=self.BG)

        self.width = 920
        self.height = 560

        self.status = tk.Label(
            root,
            text="Загрузка…",
            anchor="w",
            padx=12,
            pady=6,
            bg="#111827",
            fg="#f9fafb",
            font=("Menlo", 12),
        )
        self.status.pack(side=tk.TOP, fill=tk.X)

        self.help = tk.Label(
            root,
            text=(
                "ЛКМ — вращение | колесо — масштаб | WASD — перенос | "
                "RF/TG/YH — поворот | Пробел — анимация | 1/2 — проекция | 0 — сброс | Esc — выход"
            ),
            anchor="w",
            padx=12,
            pady=8,
            bg="#e7dfd2",
            fg=self.TEXT,
            font=("TkDefaultFont", 12),
        )
        self.help.pack(side=tk.BOTTOM, fill=tk.X)

        self.frame = tk.Frame(root, bg="#94a3b8", padx=2, pady=2)
        self.frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.soft = SoftCanvas(self.width, self.height, self.BG)
        self.soft.flush()
        self.view = tk.Label(self.frame, image=self.soft.image, bg=self.BG, bd=0)
        self.view.pack(fill=tk.BOTH, expand=True)
        self.view.image = self.soft.image

        self.vertices, self.edges = build_letter_a()
        self.axis_verts, self.axis_edges, self.axis_labels = build_axes(160.0)

        self.tx = self.ty = self.tz = 0.0
        self.rx = math.radians(-18.0)
        self.ry = math.radians(28.0)
        self.rz = 0.0
        self.sx = self.sy = self.sz = 1.0

        self.perspective = True
        self.focal = 420.0
        self.animating = True
        self.anim_t = 0.0
        self.frame_id = 0

        self._drag: Optional[Tuple[int, int]] = None
        self._bind()

        self.root.update_idletasks()
        self._sync_size()
        self._draw()

        self.root.after(33, self._tick)
        self.root.lift()
        try:
            self.root.attributes("-topmost", True)
            self.root.after(400, lambda: self.root.attributes("-topmost", False))
        except tk.TclError:
            pass
        self.root.focus_force()

    def _bind(self) -> None:
        self.root.bind("<Key>", self._on_key)
        self.view.bind("<ButtonPress-1>", self._on_press)
        self.view.bind("<B1-Motion>", self._on_drag)
        self.view.bind("<ButtonRelease-1>", self._on_release)
        self.view.bind("<MouseWheel>", self._on_wheel)
        self.view.bind("<Button-4>", lambda e: self._scale_by(1.08))
        self.view.bind("<Button-5>", lambda e: self._scale_by(1 / 1.08))
        self.frame.bind("<Configure>", self._on_configure)
        self.view.focus_set()

    def _sync_size(self) -> None:
        w = self.frame.winfo_width() - 8
        h = self.frame.winfo_height() - 8
        if w > 40 and h > 40:
            self.width, self.height = w, h

    def _on_configure(self, _event: tk.Event) -> None:
        prev = (self.soft.width, self.soft.height)
        self._sync_size()
        if (self.width, self.height) != prev:
            self.soft.resize(self.width, self.height)
            self.view.configure(image=self.soft.image)
            self.view.image = self.soft.image

    def model_matrix(self) -> Mat4:
        m = identity()
        m = mat_mul(scale(self.sx, self.sy, self.sz), m)
        m = mat_mul(rotate_x(self.rx), m)
        m = mat_mul(rotate_y(self.ry), m)
        m = mat_mul(rotate_z(self.rz), m)
        m = mat_mul(translate(self.tx, self.ty, self.tz), m)
        return m

    def animation_matrix(self) -> Mat4:
        if not self.animating:
            return identity()
        spin = rotate_z(self.anim_t * 0.9)
        pulse = 1.0 + 0.08 * math.sin(self.anim_t * 2.2)
        return mat_mul(spin, scale(pulse, pulse, pulse))

    def project(self, p: Vec3) -> Tuple[float, float]:
        x, y, z = p
        if self.perspective:
            return perspective_project(
                (x, y, z),
                focal=self.focal,
                screen_w=self.width,
                screen_h=self.height,
            )
        iso_x = x - 0.45 * y
        iso_y = z + 0.35 * y
        return self.width * 0.5 + iso_x, self.height * 0.5 - iso_y

    def _draw(self) -> None:
        if self.soft.width != self.width or self.soft.height != self.height:
            self.soft.resize(self.width, self.height)

        self.soft.clear()
        world = mat_mul(self.animation_matrix(), self.model_matrix())

        axis_m = mat_mul(
            rotate_z(self.rz),
            mat_mul(rotate_y(self.ry), rotate_x(self.rx)),
        )
        axis_pts = transform_points(axis_m, self.axis_verts)
        for (i, j), color in zip(self.axis_edges, self.AXIS):
            x1, y1 = self.project(axis_pts[i])
            x2, y2 = self.project(axis_pts[j])
            self.soft.line(x1, y1, x2, y2, color)

        pts = transform_points(world, self.vertices)
        projected = [self.project(p) for p in pts]
        for a, b in self.edges:
            x1, y1 = projected[a]
            x2, y2 = projected[b]
            self.soft.line(x1, y1, x2, y2, self.LINE)

        img = self.soft.flush()
        self.view.configure(image=img)
        self.view.image = img

        mode = "perspective" if self.perspective else "orthogonal"
        anim = "on" if self.animating else "off"
        self.frame_id += 1
        self.status.configure(
            text=(
                f"frame {self.frame_id} | {mode} | anim {anim} | "
                f"T=({self.tx:.0f},{self.ty:.0f},{self.tz:.0f}) "
                f"R=({math.degrees(self.rx):.0f},{math.degrees(self.ry):.0f},{math.degrees(self.rz):.0f}) "
                f"S={self.sx:.2f} | {self.width}x{self.height}"
            )
        )

    def _tick(self) -> None:
        try:
            if not self.view.winfo_exists():
                return
            if self.animating:
                self.anim_t += 0.04
            self._draw()
        except tk.TclError:
            return
        except Exception as exc:
            self.status.configure(text=f"Ошибка: {exc}")
            import traceback

            traceback.print_exc()
            return
        self.root.after(33, self._tick)

    def _scale_by(self, factor: float) -> None:
        self.sx = max(0.15, min(5.0, self.sx * factor))
        self.sy = self.sz = self.sx

    def _on_key(self, event: tk.Event) -> None:
        key = event.keysym.lower()
        step = 8.0
        rot = math.radians(5.0)

        if key == "escape":
            self.root.destroy()
            return
        if key == "space":
            self.animating = not self.animating
            return
        if key == "1":
            self.perspective = False
            return
        if key == "2":
            self.perspective = True
            return
        if key == "0":
            self.tx = self.ty = self.tz = 0.0
            self.rx = math.radians(-18.0)
            self.ry = math.radians(28.0)
            self.rz = 0.0
            self.sx = self.sy = self.sz = 1.0
            self.anim_t = 0.0
            return

        if key in ("left", "a"):
            self.tx -= step
        elif key in ("right", "d"):
            self.tx += step
        elif key in ("up", "w"):
            self.tz += step
        elif key in ("down", "s"):
            self.tz -= step
        elif key == "q":
            self.ty -= step
        elif key == "e":
            self.ty += step
        elif key == "r":
            self.rx += rot
        elif key == "f":
            self.rx -= rot
        elif key == "t":
            self.ry += rot
        elif key == "g":
            self.ry -= rot
        elif key == "y":
            self.rz += rot
        elif key == "h":
            self.rz -= rot
        elif key in ("plus", "equal", "kp_add"):
            self._scale_by(1.08)
        elif key in ("minus", "underscore", "kp_subtract"):
            self._scale_by(1 / 1.08)

    def _on_press(self, event: tk.Event) -> None:
        self._drag = (event.x, event.y)
        self.view.focus_set()

    def _on_drag(self, event: tk.Event) -> None:
        if self._drag is None:
            return
        dx = event.x - self._drag[0]
        dy = event.y - self._drag[1]
        self._drag = (event.x, event.y)
        self.ry += dx * 0.01
        self.rx += dy * 0.01

    def _on_release(self, _event: tk.Event) -> None:
        self._drag = None

    def _on_wheel(self, event: tk.Event) -> None:
        delta = event.delta
        if abs(delta) >= 120:
            delta /= 120
        self._scale_by(1.08 if delta > 0 else 1 / 1.08)


def main() -> None:
    root = tk.Tk()
    LetterAApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
