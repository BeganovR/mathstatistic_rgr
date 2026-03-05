"""
Генерация Excel-файла со статистической обработкой выборки X (N=121..240).
Структура строго по ТЗ с доски: 8 листов.
Запуск: python generate_excel.py
"""

import math
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Исходные данные
# ---------------------------------------------------------------------------

N_START = 121
X_VALUES = [
    172, 173, 173, 165, 167, 173, 184, 163, 179, 161, 162, 158,
    171, 177, 164, 166, 171, 174, 170, 174, 169, 174, 169, 175,
    167, 172, 168, 163, 168, 161, 173, 164, 167, 164, 173, 176,
    172, 167, 173, 161, 171, 169, 161, 170, 174, 168, 164, 170,
    164, 162, 166, 172, 169, 169, 163, 178, 166, 168, 168, 180,
    163, 165, 163, 158, 171, 175, 170, 165, 184, 169, 167, 167,
    179, 165, 173, 161, 166, 164, 159, 175, 169, 172, 172, 167,
    160, 156, 161, 174, 167, 174, 167, 168, 168, 167, 167, 171,
    168, 162, 174, 173, 173, 165, 167, 172, 176, 174, 171, 169,
    161, 173, 170, 176, 171, 166, 171, 167, 156, 167, 166, 167,
]

n = len(X_VALUES)  # 120

# ---------------------------------------------------------------------------
# Стилевые константы
# ---------------------------------------------------------------------------

HEADER_FILL     = PatternFill("solid", fgColor="4472C4")
SUBHEADER_FILL  = PatternFill("solid", fgColor="BDD7EE")
SUM_FILL        = PatternFill("solid", fgColor="DDEBF7")
LIGHT_FILL      = PatternFill("solid", fgColor="EBF3FB")
RESULT_FILL     = PatternFill("solid", fgColor="FFF2CC")

HEADER_FONT     = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
SUBHEADER_FONT  = Font(bold=True, name="Calibri", size=11)
NORMAL_FONT     = Font(name="Calibri", size=11)
BOLD_FONT       = Font(bold=True, name="Calibri", size=11)

CENTER  = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT    = Alignment(horizontal="left",   vertical="center")
RIGHT   = Alignment(horizontal="right",  vertical="center")

THIN   = Side(border_style="thin",   color="000000")
MEDIUM = Side(border_style="medium", color="000000")


def thin_border():
    return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def medium_border():
    return Border(left=MEDIUM, right=MEDIUM, top=MEDIUM, bottom=MEDIUM)


def h(cell, text=None, fill=None):
    """Стиль заголовка (тёмный фон, белый жирный шрифт)."""
    if text is not None:
        cell.value = text
    cell.font = HEADER_FONT
    cell.fill = fill or HEADER_FILL
    cell.alignment = CENTER
    cell.border = thin_border()


def sh(cell, text=None):
    """Стиль подзаголовка (светло-синий фон, жирный)."""
    if text is not None:
        cell.value = text
    cell.font = SUBHEADER_FONT
    cell.fill = SUBHEADER_FILL
    cell.alignment = CENTER
    cell.border = thin_border()


def d(cell, value=None, align=CENTER):
    """Стиль данных."""
    if value is not None:
        cell.value = value
    cell.font = NORMAL_FONT
    cell.alignment = align
    cell.border = thin_border()


def s(cell, value=None):
    """Стиль строки суммы."""
    if value is not None:
        cell.value = value
    cell.font = BOLD_FONT
    cell.fill = SUM_FILL
    cell.alignment = CENTER
    cell.border = thin_border()


def lbl(cell, text, bold=False):
    """Метка вне таблицы (без рамки)."""
    cell.value = text
    cell.font = BOLD_FONT if bold else NORMAL_FONT
    cell.alignment = LEFT


def val(cell, value, bold=False):
    """Значение вне таблицы (без рамки)."""
    cell.value = value
    cell.font = BOLD_FONT if bold else NORMAL_FONT
    cell.alignment = LEFT


def set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width


# ---------------------------------------------------------------------------
# Статистические вычисления
# ---------------------------------------------------------------------------

sorted_x = sorted(X_VALUES)
x_min = sorted_x[0]
x_max = sorted_x[-1]
R = x_max - x_min
h_val = R / (1 + 3.32 * math.log10(n))

# Статистический ряд
unique_vals = sorted(set(sorted_x))
freq = {x: sorted_x.count(x) for x in unique_vals}
rel_freq = {x: freq[x] / n for x in unique_vals}

# Интервальный ряд
num_intervals = round(1 + 3.32 * math.log10(n))
h_int = R / num_intervals

intervals = []
for i in range(num_intervals):
    a = x_min + i * h_int
    b = x_min + (i + 1) * h_int
    mid = (a + b) / 2
    # count values in [a, b) or [a, b] for last
    if i < num_intervals - 1:
        cnt = sum(1 for xv in sorted_x if a <= xv < b)
    else:
        cnt = sum(1 for xv in sorted_x if a <= xv <= b)
    w = cnt / n
    intervals.append({"a": a, "b": b, "mid": mid, "k": cnt, "w": w, "wh": w / h_int})

# Мода по интервальному ряду (середина модального класса)
modal = max(intervals, key=lambda iv: iv["k"])
mo_idx = intervals.index(modal)
k_mo = modal["k"]
k_prev = intervals[mo_idx - 1]["k"] if mo_idx > 0 else 0
k_next = intervals[mo_idx + 1]["k"] if mo_idx < len(intervals) - 1 else 0
denom_mo = (k_mo - k_prev) + (k_mo - k_next)
Mo = modal["a"] + h_int * (k_mo - k_prev) / denom_mo if denom_mo != 0 else modal["mid"]

# Медиана по интервальному ряду
half = n / 2
cum = 0
Me = None
for iv in intervals:
    if cum + iv["k"] >= half:
        Me = iv["a"] + h_int * (half - cum) / iv["k"]
        break
    cum += iv["k"]

# Начальные моменты (по точечному ряду xi, wi)
alpha1 = sum(x * rel_freq[x] for x in unique_vals)
alpha2 = sum(x**2 * rel_freq[x] for x in unique_vals)
alpha3 = sum(x**3 * rel_freq[x] for x in unique_vals)
alpha4 = sum(x**4 * rel_freq[x] for x in unique_vals)

# Характеристики выборки
x_mean = alpha1                                           # x̄в = α₁
D_v    = alpha2 - alpha1**2                              # Dв = α₂ - α₁²
sigma_v = math.sqrt(D_v)                                 # σв = √Dв
beta3  = alpha3 - 3*alpha1*alpha2 + 2*alpha1**3          # β₃
beta4  = alpha4 - 4*alpha1*alpha3 + 6*alpha1**2*alpha2 - 3*alpha1**4  # β₄
A_val  = beta3 / sigma_v**3                              # Асимметрия
E_val  = beta4 / sigma_v**4 - 3                          # Эксцесс

# Точечные оценки
S2 = (n / (n - 1)) * D_v       # S² = n/(n-1)·Dв
S  = math.sqrt(S2)              # S  = √S²

# Интервальные оценки (γ = 0.95)
gamma = 0.95
t_val = 1.96
q_val = 0.143
Delta = t_val * S / math.sqrt(n)
x_low  = x_mean - Delta
x_high = x_mean + Delta
s_low  = S * (1 - q_val)
s_high = S * (1 + q_val)

# Накопленные частоты F*(x) для точечного ряда
cum_w = 0.0
F_vals = []
for x in unique_vals:
    cum_w += rel_freq[x]
    F_vals.append(cum_w)

# ---------------------------------------------------------------------------
# Создание книги
# ---------------------------------------------------------------------------

wb = Workbook()
wb.remove(wb.active)  # удалить лист по умолчанию


# ===========================================================================
# Лист 1: Выборка X
# ===========================================================================

ws1 = wb.create_sheet("1. Выборка X")
ws1.freeze_panes = "A2"

set_col_width(ws1, 1, 12)
set_col_width(ws1, 2, 12)

h(ws1["A1"], "N")
h(ws1["B1"], "X")

for i, xv in enumerate(X_VALUES):
    row = i + 2
    d(ws1.cell(row, 1), N_START + i)
    d(ws1.cell(row, 2), xv)


# ===========================================================================
# Лист 2: Вариационный ряд
# ===========================================================================

ws2 = wb.create_sheet("2. Вариационный ряд")
ws2.freeze_panes = "A2"

set_col_width(ws2, 1, 12)
set_col_width(ws2, 2, 12)

h(ws2["A1"], "№")
h(ws2["B1"], "X (вариационный ряд)")

for i, xv in enumerate(sorted_x):
    row = i + 2
    d(ws2.cell(row, 1), i + 1)
    d(ws2.cell(row, 2), xv)


# ===========================================================================
# Лист Табл.1: Статистический ряд
# ===========================================================================

ws_t1 = wb.create_sheet("Табл.1 - Стат. ряд")

# Заголовок
ws_t1.merge_cells("A1:C1")
cell = ws_t1["A1"]
cell.value = "Табл.1. Статистический ряд"
cell.font = Font(bold=True, name="Calibri", size=13)
cell.alignment = CENTER
cell.fill = HEADER_FILL
cell.font = HEADER_FONT

# Шапка таблицы
h(ws_t1["A2"], "xᵢ")
h(ws_t1["B2"], "kᵢ")
h(ws_t1["C2"], "wᵢ = kᵢ/n")

set_col_width(ws_t1, 1, 14)
set_col_width(ws_t1, 2, 10)
set_col_width(ws_t1, 3, 16)

# Данные
row = 3
for x in unique_vals:
    d(ws_t1.cell(row, 1), x)
    d(ws_t1.cell(row, 2), freq[x])
    d(ws_t1.cell(row, 3), round(rel_freq[x], 6))
    row += 1

# Строка суммы
s(ws_t1.cell(row, 1), "Σ")
s(ws_t1.cell(row, 2), n)
s(ws_t1.cell(row, 3), 1)

# Параметры ниже таблицы
info_row = row + 2
lbl(ws_t1.cell(info_row,     1), f"n = {n}", bold=True)
lbl(ws_t1.cell(info_row + 1, 1), f"Xmin = {x_min}", bold=True)
lbl(ws_t1.cell(info_row + 2, 1), f"Xmax = {x_max}", bold=True)
lbl(ws_t1.cell(info_row + 3, 1), f"R = Xmax − Xmin = {x_max} − {x_min} = {R}", bold=True)
lbl(ws_t1.cell(info_row + 4, 1),
    f"h = R / (1 + 3.32·lg n) = {R} / (1 + 3.32·lg {n}) ≈ {h_val:.4f}  →  h (интервал) = {h_int:.4f}", bold=True)
lbl(ws_t1.cell(info_row + 5, 1),
    f"Число интервалов: {num_intervals}", bold=True)


# ===========================================================================
# Лист Табл.2: Интервальный ряд + гистограмма
# ===========================================================================

ws_t2 = wb.create_sheet("Табл.2 - Интервальный ряд")

# Заголовок
ws_t2.merge_cells("A1:E1")
cell = ws_t2["A1"]
cell.value = "Табл.2. Интервальный ряд"
cell.font = HEADER_FONT
cell.alignment = CENTER
cell.fill = HEADER_FILL

# Шапка
h(ws_t2["A2"], "Интервал [a; b)")
h(ws_t2["B2"], "Середина xᵢ*")
h(ws_t2["C2"], "kᵢ")
h(ws_t2["D2"], "wᵢ")
h(ws_t2["E2"], "wᵢ/h")

set_col_width(ws_t2, 1, 20)
set_col_width(ws_t2, 2, 16)
set_col_width(ws_t2, 3, 10)
set_col_width(ws_t2, 4, 10)
set_col_width(ws_t2, 5, 14)

# Данные
row = 3
for iv in intervals:
    d(ws_t2.cell(row, 1), f"[{iv['a']:.2f}; {iv['b']:.2f})")
    d(ws_t2.cell(row, 2), round(iv["mid"], 4))
    d(ws_t2.cell(row, 3), iv["k"])
    d(ws_t2.cell(row, 4), round(iv["w"], 6))
    d(ws_t2.cell(row, 5), round(iv["wh"], 6))
    row += 1

# Строка суммы
s(ws_t2.cell(row, 1), "Σ")
s(ws_t2.cell(row, 2), "—")
s(ws_t2.cell(row, 3), n)
s(ws_t2.cell(row, 4), 1)
s(ws_t2.cell(row, 5), "—")

# Мода и медиана
info_row = row + 2
ws_t2.cell(info_row, 1).value = "Мода:"
ws_t2.cell(info_row, 1).font = BOLD_FONT
ws_t2.cell(info_row, 2).value = round(Mo, 4)
ws_t2.cell(info_row, 2).font = NORMAL_FONT

ws_t2.cell(info_row + 1, 1).value = "Медиана:"
ws_t2.cell(info_row + 1, 1).font = BOLD_FONT
ws_t2.cell(info_row + 1, 2).value = round(Me, 4)
ws_t2.cell(info_row + 1, 2).font = NORMAL_FONT

# --- Гистограмма ---
chart_start_row = info_row + 4

# Ссылки на данные: середины (B3:B{row-1}), wᵢ/h (E3:E{row-1})
data_row_start = 3
data_row_end   = row - 1   # последняя строка данных (без строки суммы)

bar = BarChart()
bar.type   = "col"
bar.style  = 10
bar.title  = "Гистограмма"
bar.y_axis.title = "wᵢ/h (плотность)"
bar.x_axis.title = "xᵢ*"
bar.grouping = "clustered"

# Данные (wᵢ/h)
data_ref = Reference(ws_t2, min_col=5, min_row=2, max_row=data_row_end)
bar.add_data(data_ref, titles_from_data=True)

# Категории (середины)
cats_ref = Reference(ws_t2, min_col=2, min_row=data_row_start, max_row=data_row_end)
bar.set_categories(cats_ref)

# Цвет и нулевые промежутки (настоящая гистограмма)
bar.series[0].graphicalProperties.solidFill = "4472C4"
bar.gapWidth = 0

bar.width  = 20
bar.height = 14

anchor = f"A{chart_start_row}"
ws_t2.add_chart(bar, anchor)


# ===========================================================================
# Лист Табл.3: Полигон и F*(x)
# ===========================================================================

ws_t3 = wb.create_sheet("Табл.3 - Полигон и F(x)")

# --- Таблица полигона ---
ws_t3.merge_cells("A1:B1")
cell = ws_t3["A1"]
cell.value = "Табл.3а. Точечный ряд (полигон)"
cell.font = HEADER_FONT
cell.alignment = CENTER
cell.fill = HEADER_FILL

h(ws_t3["A2"], "xᵢ")
h(ws_t3["B2"], "wᵢ")

set_col_width(ws_t3, 1, 14)
set_col_width(ws_t3, 2, 14)
set_col_width(ws_t3, 4, 14)
set_col_width(ws_t3, 5, 14)

poly_start = 3
for i, x in enumerate(unique_vals):
    row = poly_start + i
    d(ws_t3.cell(row, 1), x)
    d(ws_t3.cell(row, 2), round(rel_freq[x], 6))

poly_end = poly_start + len(unique_vals) - 1

# --- Таблица F*(x) (справа, col 4-5) ---
ws_t3.merge_cells("D1:E1")
cell = ws_t3["D1"]
cell.value = "Табл.3б. F*(x)"
cell.font = HEADER_FONT
cell.alignment = CENTER
cell.fill = HEADER_FILL

h(ws_t3["D2"], "xᵢ")
h(ws_t3["E2"], "F*(xᵢ)")

for i, x in enumerate(unique_vals):
    row = poly_start + i
    d(ws_t3.cell(row, 4), x)
    d(ws_t3.cell(row, 5), round(F_vals[i], 6))

# --- Полигон частот (LineChart) ---
chart_row = poly_end + 4

line_poly = LineChart()
line_poly.style  = 10
line_poly.title  = "Полигон частот"
line_poly.y_axis.title = "wᵢ"
line_poly.x_axis.title = "xᵢ"

y_ref = Reference(ws_t3, min_col=2, min_row=2, max_row=poly_end)
line_poly.add_data(y_ref, titles_from_data=True)

x_ref = Reference(ws_t3, min_col=1, min_row=poly_start, max_row=poly_end)
line_poly.set_categories(x_ref)

ser = line_poly.series[0]
ser.graphicalProperties.line.solidFill = "2E75B6"
ser.graphicalProperties.line.width = 20000  # ~1.5 pt
ser.marker.symbol  = "circle"
ser.marker.size    = 6
ser.marker.graphicalProperties.solidFill   = "2E75B6"
ser.marker.graphicalProperties.line.solidFill = "2E75B6"

line_poly.x_axis.delete = False
line_poly.y_axis.delete = False

line_poly.width  = 20
line_poly.height = 14

ws_t3.add_chart(line_poly, f"A{chart_row}")

# --- График F*(x) ---
fx_chart_row = chart_row + 30

line_fx = LineChart()
line_fx.style  = 10
line_fx.title  = "Эмпирическая функция распределения F*(x)"
line_fx.y_axis.title = "F*(x)"
line_fx.x_axis.title = "xᵢ"

y_ref2 = Reference(ws_t3, min_col=5, min_row=2, max_row=poly_end)
line_fx.add_data(y_ref2, titles_from_data=True)

x_ref2 = Reference(ws_t3, min_col=4, min_row=poly_start, max_row=poly_end)
line_fx.set_categories(x_ref2)

ser2 = line_fx.series[0]
ser2.graphicalProperties.line.solidFill = "ED7D31"
ser2.graphicalProperties.line.width = 20000

line_fx.y_axis.scaling.min = 0
line_fx.y_axis.scaling.max = 1

line_fx.x_axis.delete = False
line_fx.y_axis.delete = False

line_fx.width  = 20
line_fx.height = 14

ws_t3.add_chart(line_fx, f"A{fx_chart_row}")


# ===========================================================================
# Лист Табл.4: Начальные моменты и характеристики выборки
# ===========================================================================

ws_t4 = wb.create_sheet("Табл.4 - Начальные моменты")

# Заголовок
ws_t4.merge_cells("A1:F1")
cell = ws_t4["A1"]
cell.value = "Табл.4. Вычисление начальных моментов"
cell.font = HEADER_FONT
cell.alignment = CENTER
cell.fill = HEADER_FILL

# Шапка
h(ws_t4["A2"], "xᵢ")
h(ws_t4["B2"], "wᵢ = kᵢ/n")
h(ws_t4["C2"], "xᵢ·wᵢ")
h(ws_t4["D2"], "xᵢ²·wᵢ")
h(ws_t4["E2"], "xᵢ³·wᵢ")
h(ws_t4["F2"], "xᵢ⁴·wᵢ")

set_col_width(ws_t4, 1, 12)
set_col_width(ws_t4, 2, 14)
set_col_width(ws_t4, 3, 16)
set_col_width(ws_t4, 4, 18)
set_col_width(ws_t4, 5, 22)
set_col_width(ws_t4, 6, 26)

row = 3
for x in unique_vals:
    w = rel_freq[x]
    d(ws_t4.cell(row, 1), x)
    d(ws_t4.cell(row, 2), round(w, 6))
    d(ws_t4.cell(row, 3), round(x * w, 6))
    d(ws_t4.cell(row, 4), round(x**2 * w, 4))
    d(ws_t4.cell(row, 5), round(x**3 * w, 2))
    d(ws_t4.cell(row, 6), round(x**4 * w, 2))
    row += 1

# Строка Σ
s(ws_t4.cell(row, 1), "Σ")
s(ws_t4.cell(row, 2), 1)
s(ws_t4.cell(row, 3), f"α₁ = {round(alpha1, 6)}")
s(ws_t4.cell(row, 4), f"α₂ = {round(alpha2, 4)}")
s(ws_t4.cell(row, 5), f"α₃ = {round(alpha3, 2)}")
s(ws_t4.cell(row, 6), f"α₄ = {round(alpha4, 2)}")

# Значения моментов
info_row = row + 2
lbl(ws_t4.cell(info_row,     1), "Начальные моменты:", bold=True)
lbl(ws_t4.cell(info_row + 1, 1), f"α₁ = {round(alpha1, 6)}", bold=True)
lbl(ws_t4.cell(info_row + 2, 1), f"α₂ = {round(alpha2, 6)}", bold=True)
lbl(ws_t4.cell(info_row + 3, 1), f"α₃ = {round(alpha3, 4)}", bold=True)
lbl(ws_t4.cell(info_row + 4, 1), f"α₄ = {round(alpha4, 4)}", bold=True)

# Характеристики выборки
chr_row = info_row + 6
lbl(ws_t4.cell(chr_row, 1), "Характеристики выборки:", bold=True)

ws_t4.cell(chr_row + 1, 1).value = "1) x̄в = α₁"
ws_t4.cell(chr_row + 1, 1).font  = NORMAL_FONT
ws_t4.cell(chr_row + 1, 2).value = f"x̄в = {round(x_mean, 6)}"
ws_t4.cell(chr_row + 1, 2).font  = BOLD_FONT

ws_t4.cell(chr_row + 2, 1).value = "2) Dв = α₂ − α₁²"
ws_t4.cell(chr_row + 2, 1).font  = NORMAL_FONT
ws_t4.cell(chr_row + 2, 2).value = f"Dв = {round(D_v, 6)}"
ws_t4.cell(chr_row + 2, 2).font  = BOLD_FONT

ws_t4.cell(chr_row + 3, 1).value = "3) σв = √Dв"
ws_t4.cell(chr_row + 3, 1).font  = NORMAL_FONT
ws_t4.cell(chr_row + 3, 2).value = f"σв = {round(sigma_v, 6)}"
ws_t4.cell(chr_row + 3, 2).font  = BOLD_FONT

ws_t4.cell(chr_row + 4, 1).value = "4) A = β₃/σ³,  β₃ = α₃ − 3α₁α₂ + 2α₁³"
ws_t4.cell(chr_row + 4, 1).font  = NORMAL_FONT
ws_t4.cell(chr_row + 4, 2).value = f"β₃ = {round(beta3, 6)},  A = {round(A_val, 6)}"
ws_t4.cell(chr_row + 4, 2).font  = BOLD_FONT

ws_t4.cell(chr_row + 5, 1).value = "5) E = β₄/σ⁴ − 3,  β₄ = α₄ − 4α₁α₃ + 6α₁²α₂ − 3α₁⁴"
ws_t4.cell(chr_row + 5, 1).font  = NORMAL_FONT
ws_t4.cell(chr_row + 5, 2).value = f"β₄ = {round(beta4, 6)},  E = {round(E_val, 6)}"
ws_t4.cell(chr_row + 5, 2).font  = BOLD_FONT

set_col_width(ws_t4, 2, 32)


# ===========================================================================
# Лист 7: Точечные оценки
# ===========================================================================

ws7 = wb.create_sheet("7. Точечные оценки")

set_col_width(ws7, 1, 46)
set_col_width(ws7, 2, 28)

ws7.merge_cells("A1:B1")
cell = ws7["A1"]
cell.value = "7. Вычисление точечных оценок"
cell.font = HEADER_FONT
cell.alignment = CENTER
cell.fill = HEADER_FILL

row = 3
ws7.cell(row, 1).value = "7.1 Метод статистических оценок"
ws7.cell(row, 1).font  = BOLD_FONT

row += 1
ws7.cell(row, 1).value = "x̄г = x̄в"
ws7.cell(row, 1).font  = NORMAL_FONT
ws7.cell(row, 2).value = f"x̄г = {round(x_mean, 6)}"
ws7.cell(row, 2).font  = BOLD_FONT

row += 1
ws7.cell(row, 1).value = "S² = n/(n−1) · Dв"
ws7.cell(row, 1).font  = NORMAL_FONT
ws7.cell(row, 2).value = f"S² = {round(S2, 6)}"
ws7.cell(row, 2).font  = BOLD_FONT

row += 1
ws7.cell(row, 1).value = "S = √(n/(n−1)) · σв"
ws7.cell(row, 1).font  = NORMAL_FONT
ws7.cell(row, 2).value = f"S = {round(S, 6)}"
ws7.cell(row, 2).font  = BOLD_FONT

row += 2
ws7.cell(row, 1).value = "7.2 Метод моментов"
ws7.cell(row, 1).font  = BOLD_FONT

row += 1
ws7.cell(row, 1).value = "α₁ = α₁' ⇒ x̄г = x̄в"
ws7.cell(row, 1).font  = NORMAL_FONT
ws7.cell(row, 2).value = f"x̄г = {round(x_mean, 6)}"
ws7.cell(row, 2).font  = BOLD_FONT

row += 1
ws7.cell(row, 1).value = "β₂ = β₂ ⇒ Dг = Dв,  σг = σв"
ws7.cell(row, 1).font  = NORMAL_FONT
ws7.cell(row, 2).value = f"Dг = {round(D_v, 6)},  σг = {round(sigma_v, 6)}"
ws7.cell(row, 2).font  = BOLD_FONT


# ===========================================================================
# Лист 8: Интервальные оценки
# ===========================================================================

ws8 = wb.create_sheet("8. Интервальные оценки")

set_col_width(ws8, 1, 50)
set_col_width(ws8, 2, 30)

ws8.merge_cells("A1:B1")
cell = ws8["A1"]
cell.value = "8. Интервальные оценки при γ = 0.95"
cell.font = HEADER_FONT
cell.alignment = CENTER
cell.fill = HEADER_FILL

row = 3
ws8.cell(row, 1).value = f"Уровень надёжности: γ = {gamma}"
ws8.cell(row, 1).font  = NORMAL_FONT

row += 2
ws8.cell(row, 1).value = "8.1 Интервальная оценка среднего x̄г"
ws8.cell(row, 1).font  = BOLD_FONT

row += 1
ws8.cell(row, 1).value = f"t = {t_val}  (при γ = {gamma})"
ws8.cell(row, 1).font  = NORMAL_FONT

row += 1
ws8.cell(row, 1).value = "Δ = t · S / √n"
ws8.cell(row, 1).font  = NORMAL_FONT
ws8.cell(row, 2).value = f"Δ = {round(Delta, 6)}"
ws8.cell(row, 2).font  = BOLD_FONT

row += 1
ws8.cell(row, 1).value = "x̄г − Δ < x̄ < x̄г + Δ"
ws8.cell(row, 1).font  = NORMAL_FONT
ws8.cell(row, 2).value = f"({round(x_low, 4)} ; {round(x_high, 4)})"
ws8.cell(row, 2).font  = BOLD_FONT

row += 2
ws8.cell(row, 1).value = "8.2 Интервальная оценка среднеквадратичного отклонения σ"
ws8.cell(row, 1).font  = BOLD_FONT

row += 1
ws8.cell(row, 1).value = f"q = {q_val}"
ws8.cell(row, 1).font  = NORMAL_FONT

row += 1
ws8.cell(row, 1).value = "S(1 − q) < σ < S(1 + q)"
ws8.cell(row, 1).font  = NORMAL_FONT
ws8.cell(row, 2).value = f"({round(s_low, 4)} ; {round(s_high, 4)})"
ws8.cell(row, 2).font  = BOLD_FONT


# ---------------------------------------------------------------------------
# Сохранение
# ---------------------------------------------------------------------------

wb.save("mathstat_rgr.xlsx")
print("Файл mathstat_rgr.xlsx сохранён.")
print(f"  n={n}, Xmin={x_min}, Xmax={x_max}, R={R}, h={h_val:.4f}")
print(f"  Число интервалов: {num_intervals}")
print(f"  Mo={Mo:.4f}, Me={Me:.4f}")
print(f"  α₁={alpha1:.6f}, α₂={alpha2:.6f}, α₃={alpha3:.4f}, α₄={alpha4:.4f}")
print(f"  x̄в={x_mean:.6f}, Dв={D_v:.6f}, σв={sigma_v:.6f}")
print(f"  A={A_val:.6f}, E={E_val:.6f}")
print(f"  S²={S2:.6f}, S={S:.6f}")
print(f"  Δ={Delta:.6f}, x̄±Δ=({x_low:.4f}; {x_high:.4f})")
print(f"  σ интервал: ({s_low:.4f}; {s_high:.4f})")
