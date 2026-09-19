#!/usr/bin/env python3
"""Отчёт ЛР1 в формате титульного листа АлтГТУ (как в scheduling-labs)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "report_assets"
OUT = ROOT / "reports" / "Кованов_ПИ32_Компьютерная_графика_Лаб1.docx"
GITHUB = "https://github.com/MICROWAVE-web/lab1-3d-letter-a"


def font(run, name="Times New Roman", size=14, bold=False, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def p(
    doc,
    text="",
    *,
    size=14,
    bold=False,
    italic=False,
    center=False,
    indent=False,
    after=6,
    before=0,
    mono=False,
):
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.space_after = Pt(after)
    pf.space_before = Pt(before)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.first_line_indent = Cm(1.25) if indent else Cm(0)
    if center:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run(text)
    font(run, name="Courier New" if mono else "Times New Roman", size=size, bold=bold, italic=italic)
    return para


def formula(doc, text):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(4)
    para.paragraph_format.space_after = Pt(4)
    para.paragraph_format.first_line_indent = Cm(0)
    run = para.add_run(text)
    font(run, name="Cambria Math", size=13, italic=True)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")


def matrix(doc, title, rows):
    if title:
        p(doc, title, center=True, italic=True, size=13, after=2)
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = "Table Grid"
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = ""
            rp = cell.paragraphs[0]
            rp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = rp.add_run(str(val))
            font(run, name="Cambria Math", size=12, italic=True)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def code_block(doc, lines):
    for line in lines:
        para = doc.add_paragraph()
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.line_spacing = 1.0
        para.paragraph_format.first_line_indent = Cm(0)
        run = para.add_run(line)
        font(run, name="Courier New", size=10)


def add_figure(doc, path: Path, caption: str, width_in=5.8):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(8)
    para.paragraph_format.space_after = Pt(2)
    para.paragraph_format.first_line_indent = Cm(0)
    run = para.add_run()
    run.add_picture(str(path), width=Inches(width_in))
    p(doc, caption, center=True, size=12, italic=True, after=10)


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(3)
    sec.right_margin = Cm(1.5)

    # --- Титульный лист (как в образце) ---
    for line in [
        "Министерство науки и высшего образования Российской Федерации",
        "Федеральное государственное бюджетное образовательное учреждение",
        "высшего образования",
        "«Алтайский государственный технический университет им. И. И. Ползунова»",
        "Факультет информационных технологий",
        "Кафедра прикладной математики",
    ]:
        p(doc, line, center=True, size=12, after=0)

    p(doc, "", after=20)
    p(doc, "Отчет защищен с оценкой _____", center=True, size=12, after=0)
    p(doc, "Преподаватель _____________", center=True, size=12, after=0)
    p(doc, "(подпись)", center=True, size=10, after=0)
    p(doc, "«___» ____________ 2026 г.", center=True, size=12, after=24)

    p(doc, "Отчет", center=True, bold=True, size=16, after=8)
    p(doc, "По лабораторной работе №1", center=True, bold=True, size=14, after=8)
    p(
        doc,
        "«Реализация алгоритмов построения проекций трёхмерных объектов.",
        center=True,
        bold=True,
        size=14,
        after=0,
    )
    p(
        doc,
        "Аффинные преобразования в пространстве»",
        center=True,
        bold=True,
        size=14,
        after=8,
    )
    p(doc, "по дисциплине «Компьютерная графика»", center=True, size=14, after=20)
    p(doc, "Студент группы ПИ-32 Кованов Алексей Вадимович", center=True, size=14, after=4)
    p(doc, "Преподаватель Потапов Даниил Петрович", center=True, size=14, after=24)
    p(doc, "Барнаул 2026", center=True, size=14, after=12)

    doc.add_page_break()

    # --- Содержание работы ---
    p(doc, "Цель работы", bold=True, size=14, after=6)
    p(
        doc,
        "Изучение алгоритмов построения проекций трёхмерных объектов и аффинных "
        "преобразований в пространстве; программная реализация каркасной модели "
        "объёмной буквы «А» с управлением переносом, поворотом, масштабированием "
        "и динамическим преобразованием (анимацией) с помощью мыши и клавиатуры.",
        indent=True,
    )

    p(doc, "Ссылка на исходный код программы:", bold=True, after=4)
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.first_line_indent = Cm(0)
    run = para.add_run(GITHUB)
    font(run, size=12, bold=True)
    run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)

    p(
        doc,
        "Репозиторий GitHub: MICROWAVE-web / lab1-3d-letter-a. "
        "Язык Python 3. Матрицы 4×4, аффинные преобразования и проекции реализованы вручную "
        "(без OpenGL, NumPy и готовых 3D-библиотек). Для окна используется стандартный модуль tkinter.",
        indent=True,
        after=10,
    )

    # --- 1 Теория ---
    p(doc, "1. Краткие теоретические сведения", bold=True, size=14, before=10)

    p(doc, "1.1. Каркасная модель", bold=True, size=14)
    p(
        doc,
        "Объект задаётся вершинами — точками (x, y, z) — и рёбрами (парами индексов вершин). "
        "На экране рисуются отрезки между спроецированными вершинами. Буква «А» представлена "
        "как призма: передняя и задняя грани связаны рёбрами глубины; имеются вырез между ножками, "
        "горизонтальная перекладина и отверстие над ней.",
        indent=True,
    )

    p(doc, "1.2. Однородные координаты и матрицы 4×4", bold=True, size=14)
    p(
        doc,
        "Точку (x, y, z) записывают в однородных координатах как столбец p = (x, y, z, 1)ᵀ. "
        "Четвёртая координата позволяет записать перенос, поворот и масштаб одной матрицей M "
        "и комбинировать преобразования умножением: p′ = M · p. В развёрнутом виде:",
        indent=True,
    )
    formula(doc, "x′ = m₁₁x + m₁₂y + m₁₃z + m₁₄")
    formula(doc, "y′ = m₂₁x + m₂₂y + m₂₃z + m₂₄")
    formula(doc, "z′ = m₃₁x + m₃₂y + m₃₃z + m₃₄")
    formula(doc, "w′ = m₄₁x + m₄₂y + m₄₃z + m₄₄")
    p(
        doc,
        "Для аффинных преобразований обычно w′ = 1. Если w′ ≠ 0, декартовы координаты: "
        "X = x′/w′, Y = y′/w′, Z = z′/w′.",
        indent=True,
    )

    p(doc, "1.3. Элементарные аффинные преобразования", bold=True, size=14)
    p(doc, "Перенос T:", indent=True)
    matrix(
        doc,
        None,
        [["1", "0", "0", "dx"], ["0", "1", "0", "dy"], ["0", "0", "1", "dz"], ["0", "0", "0", "1"]],
    )
    formula(doc, "x′ = x + dx,   y′ = y + dy,   z′ = z + dz")

    p(doc, "Масштабирование S:", indent=True)
    matrix(
        doc,
        None,
        [["sx", "0", "0", "0"], ["0", "sy", "0", "0"], ["0", "0", "sz", "0"], ["0", "0", "0", "1"]],
    )
    formula(doc, "x′ = sx·x,   y′ = sy·y,   z′ = sz·z")

    p(doc, "Поворот вокруг осей (c = cos φ, s = sin φ):", indent=True)
    p(doc, "Rx:", center=True, italic=True, size=12)
    matrix(
        doc,
        None,
        [["1", "0", "0", "0"], ["0", "c", "−s", "0"], ["0", "s", "c", "0"], ["0", "0", "0", "1"]],
    )
    p(doc, "Ry:", center=True, italic=True, size=12)
    matrix(
        doc,
        None,
        [["c", "0", "s", "0"], ["0", "1", "0", "0"], ["−s", "0", "c", "0"], ["0", "0", "0", "1"]],
    )
    p(doc, "Rz:", center=True, italic=True, size=12)
    matrix(
        doc,
        None,
        [["c", "−s", "0", "0"], ["s", "c", "0", "0"], ["0", "0", "1", "0"], ["0", "0", "0", "1"]],
    )

    p(doc, "1.4. Композиция преобразований", bold=True, size=14)
    p(
        doc,
        "Порядок умножения важен. В программе модельная матрица: "
        "M = T · Rz · Ry · Rx · S. С учётом анимации A: Mworld = A · M.",
        indent=True,
    )

    p(doc, "1.5. Проекции", bold=True, size=14)
    p(
        doc,
        "Ортогональная (аксонометрия в программе):",
        indent=True,
    )
    formula(doc, "Xэкр = W/2 + x − 0,45·y")
    formula(doc, "Yэкр = H/2 − (z + 0,35·y)")
    p(doc, "Перспективная (центральная) проекция, d = y + f:", indent=True)
    formula(doc, "xp = (x·f)/d,   zp = (z·f)/d")
    formula(doc, "Xэкр = W/2 + xp,   Yэкр = H/2 − zp")

    # --- 2 Структура ---
    p(doc, "2. Структура программы", bold=True, size=14, before=10)
    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    for i, (a, b) in enumerate(
        [
            ("Файл", "Назначение"),
            ("letter_a.py", "Вершины и рёбра буквы «А», оси координат"),
            ("math3d.py", "Матрицы 4×4, перенос, поворот, масштаб, проекция"),
            ("main.py", "Окно, ввод, цикл отрисовки, анимация"),
        ]
    ):
        for j, val in enumerate((a, b)):
            cell = table.cell(i, j)
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            font(run, size=12, bold=(i == 0))

    p(doc, "Запуск:", indent=True, before=8)
    code_block(doc, ["/usr/local/bin/python3 main.py", "# или", "./run.sh"])

    # --- 3 Фрагменты кода ---
    p(doc, "3. Фрагменты кода программы", bold=True, size=14, before=12)
    p(doc, "Матрица переноса и умножение матриц (math3d.py):", indent=True)
    code_block(
        doc,
        [
            "def translate(dx, dy, dz):",
            "    m = identity()",
            "    m[0][3] = dx",
            "    m[1][3] = dy",
            "    m[2][3] = dz",
            "    return m",
            "",
            "def mat_mul(a, b):",
            "    # result = a * b  (сначала b, потом a для столбца справа)",
            "    result = [[0.0] * 4 for _ in range(4)]",
            "    for i in range(4):",
            "        for j in range(4):",
            "            result[i][j] = (a[i][0]*b[0][j] + a[i][1]*b[1][j] +",
            "                            a[i][2]*b[2][j] + a[i][3]*b[3][j])",
            "    return result",
        ],
    )
    p(doc, "Перспективная проекция (math3d.py):", indent=True, before=8)
    code_block(
        doc,
        [
            "def perspective_project(point, *, focal, screen_w, screen_h, ...):",
            "    x, y, z = point",
            "    depth = y + focal",
            "    if depth < 0.1:",
            "        depth = 0.1",
            "    px = (x * focal) / depth",
            "    py = (z * focal) / depth",
            "    sx = screen_w * 0.5 + px",
            "    sy = screen_h * 0.5 - py  # ось Y экрана вниз",
            "    return sx, sy",
        ],
    )
    p(doc, "Сборка модельной матрицы (main.py):", indent=True, before=8)
    code_block(
        doc,
        [
            "def model_matrix(self):",
            "    m = identity()",
            "    m = mat_mul(scale(self.sx, self.sy, self.sz), m)",
            "    m = mat_mul(rotate_x(self.rx), m)",
            "    m = mat_mul(rotate_y(self.ry), m)",
            "    m = mat_mul(rotate_z(self.rz), m)",
            "    m = mat_mul(translate(self.tx, self.ty, self.tz), m)",
            "    return m",
        ],
    )

    # --- 4 Управление ---
    p(doc, "4. Управление программой", bold=True, size=14, before=12)
    ctrl = [
        ("Действие", "Управление"),
        ("Вращение", "ЛКМ + движение; R/F, T/G, Y/H"),
        ("Перенос", "WASD / стрелки; Q/E (глубина)"),
        ("Масштаб", "Колесо мыши; + / −"),
        ("Анимация", "Пробел"),
        ("Проекция", "1 — ортогональная, 2 — перспективная"),
        ("Сброс", "0"),
        ("Выход", "Esc"),
    ]
    t2 = doc.add_table(rows=len(ctrl), cols=2)
    t2.style = "Table Grid"
    for i, (a, b) in enumerate(ctrl):
        for j, val in enumerate((a, b)):
            cell = t2.cell(i, j)
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            font(run, size=12, bold=(i == 0))

    # --- 5 Результаты / скриншоты ---
    p(doc, "5. Результаты работы программы (скриншоты)", bold=True, size=14, before=12)
    p(
        doc,
        "Ниже приведены кадры работы программы: перспективная и ортогональная проекции, "
        "поворот, масштабирование и перенос каркасной буквы «А» с осями координат.",
        indent=True,
    )

    figures = [
        ("fig1_perspective.png", "Рис. 1. Перспективная проекция буквы «А» (исходный вид)"),
        ("fig2_orthogonal.png", "Рис. 2. Ортогональная (аксонометрическая) проекция"),
        ("fig3_rotated.png", "Рис. 3. Результат поворота объекта"),
        ("fig4_scaled.png", "Рис. 4. Результат масштабирования"),
        ("fig5_translated.png", "Рис. 5. Результат переноса объекта"),
    ]
    for fname, cap in figures:
        path = ASSETS / fname
        if path.exists():
            add_figure(doc, path, cap)

    # --- 6 Динамика ---
    p(doc, "6. Динамическое преобразование", bold=True, size=14, before=8)
    p(
        doc,
        "По заданию реализована анимация (клавиша «Пробел»): непрерывный поворот вокруг оси Z "
        "и пульсация масштаба вида 1 + 0,08·sin(t). Это аффинные преобразования с параметрами, "
        "зависящими от времени.",
        indent=True,
    )

    # --- 7 Тестирование ---
    p(doc, "7. Проверка (тестирование) программы", bold=True, size=14, before=8)
    p(doc, "Выполнены следующие проверки:", indent=True)
    for item in [
        "запуск окна и отображение каркасной буквы «А» с осями;",
        "перенос клавишами WASD / стрелками и Q/E;",
        "поворот мышью и клавишами R/F, T/G, Y/H;",
        "масштабирование колесом мыши и клавишами +/−;",
        "переключение проекций клавишами 1 и 2;",
        "включение/выключение анимации пробелом;",
        "сброс преобразований клавишей 0.",
    ]:
        para = doc.add_paragraph(style="List Bullet")
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        para.clear()
        run = para.add_run(item)
        font(run, size=14)

    p(
        doc,
        "Все перечисленные действия отрабатывают корректно: объект изменяет положение, "
        "ориентацию и размер; проекции переключаются; анимация выполняется непрерывно.",
        indent=True,
        before=6,
    )

    # --- Вывод ---
    p(doc, "Вывод", bold=True, size=14, before=12)
    p(
        doc,
        "В ходе лабораторной работы изучены и программно реализованы алгоритмы построения "
        "проекций трёхмерных объектов и аффинные преобразования в пространстве. "
        "Разработана программа отображения объёмной каркасной буквы «А» с переносом, "
        "поворотом, масштабированием, перспективной и ортогональной проекциями, а также "
        "динамическим преобразованием. Математика преобразований выполнена самостоятельно "
        "на матрицах 4×4 без использования готовых 3D-библиотек. Исходный код размещён "
        f"в репозитории {GITHUB}.",
        indent=True,
    )

    doc.save(OUT)
    print(f"Saved: {OUT}")


if __name__ == "__main__":
    build()
