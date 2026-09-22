# -*- coding: utf-8 -*-
"""
Построение графика "глубина-день" (План / Факт / Кол-во дней без НПВ)
средствами самой программы — без LibreOffice.

Идея: сам линейчатый график в файле СР — это обычный Excel XY-график,
который берёт данные с отдельного скрытого листа "График TVD - Данные".
Во всех проверенных файлах СР раскладка этого листа одинаковая
(зафиксирована шаблоном отчёта):

    столбцы C/D — серия "План"                (C — день, D — глубина, м)
    столбцы F/G — серия "Факт"                 (F — день, G — глубина, м)
    столбцы I/J — серия "Кол-во дней без НПВ"  (I — день, J — глубина, м)
    столбец  L  — дата, соответствующая дню из столбца K (K10=0 — день
                  начала работ, K11=1 и т.д.; т.е. дата растёт на 1 сутки
                  на каждую строку начиная с 10-й) — используется только
                  для подписи верхней оси датами, как в самом Excel.

данные начинаются с 10-й строки. Мы читаем эти столбцы напрямую (без
обращения к самому объекту графика — это позволяет работать с книгой,
открытой в экономном режиме read_only, и не требует LibreOffice/UNO
вообще), строим по ним компактную SVG-картинку и вставляем её в отчёт
как обычную разметку (без PNG, без base64 — легче и чётче при любом
масштабе).

Если на листе нет данных или самого листа нет вовсе — просто возвращаем
None, ничего не ломая (отчёт формируется как обычно, без этого блока).
"""

import datetime
import math

SHEET_NAME = "График TVD - Данные"
START_ROW = 10
MAX_ROW = 4000  # с большим запасом; реальные данные почти всегда заканчиваются намного раньше

# ключ, подпись, цвет, пунктир — порядок столбцов задаётся ниже (COL_MIN..COL_MAX)
SERIES_META = {
    "plan": ("План", "#2f8fd1", False),
    "fact": ("Факт", "#c14a26", False),
    "npv_free": ("Кол-во дней без НПВ", "#55606a", True),
}
SERIES = tuple((key, *SERIES_META[key]) for key in ("plan", "fact", "npv_free"))

# столбцы C..L листа "График TVD - Данные" (1-индексация openpyxl): читаем их
# все одним проходом, затем раскладываем по сериям — это на порядки быстрее,
# чем поклеточный доступ ws.cell(row, col) в режиме read_only (там он
# фактически O(n) на вызов и даёт O(n^2) суммарно на диапазонах в сотни строк).
COL_MIN = 3   # C
COL_MAX = 12  # L
COL_OFFSETS = {
    "plan": (0, 1),   # C, D  (индексы внутри кортежа строки, считая от COL_MIN)
    "fact": (3, 4),   # F, G
    "npv_free": (6, 7),  # I, J
}
DATE_COL_OFFSET = 9  # L, считая от COL_MIN=3 (12-3)


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _parse_ru_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        try:
            return datetime.datetime.strptime(v.strip(), "%d.%m.%Y").date()
        except ValueError:
            return None
    return None


def extract_depth_series(wb, sheet_name=SHEET_NAME):
    """Читает серии Плана/Факта/НПВ (и дату начала для подписи верхней оси)
    с листа графика в уже открытой книге (openpyxl, подходит и
    read_only=True — один быстрый последовательный проход по строкам).
    Возвращает словарь {"plan": [(x,y),...], "fact": [...], "npv_free":
    [...], "start_date": date|None} либо None, если подходящего листа нет
    или данных в нём нет вовсе.

    Важно: у некоторых файлов СР тег <dimension> листа-источника в XML
    занижен/устарел, из-за чего ws.max_row может лгать — поэтому явно
    задаём щедрый max_row=MAX_ROW при чтении, а не полагаемся на него."""
    if sheet_name not in wb.sheetnames:
        return None
    ws = wb[sheet_name]
    plan, fact, npv_free = [], [], []
    start_date = None
    rows = ws.iter_rows(min_row=START_ROW, max_row=MAX_ROW, min_col=COL_MIN, max_col=COL_MAX, values_only=True)
    for row_offset, row in enumerate(rows):
        for key, (xi, yi) in COL_OFFSETS.items():
            x, y = row[xi], row[yi]
            if _is_number(x) and _is_number(y):
                (plan if key == "plan" else fact if key == "fact" else npv_free).append((float(x), float(y)))
        if start_date is None:
            parsed = _parse_ru_date(row[DATE_COL_OFFSET])
            if parsed:
                start_date = parsed - datetime.timedelta(days=row_offset)
    series = {"plan": plan, "fact": fact, "npv_free": npv_free, "start_date": start_date}
    if not (plan or fact or npv_free):
        return None
    return series


def _nice_step(span, target_count=8, min_step=1.0):
    if span <= 0:
        return min_step
    raw = span / max(target_count, 1)
    magnitude = 10 ** math.floor(math.log10(raw)) if raw > 0 else 1
    residual = raw / magnitude
    if residual > 5:
        step = 10 * magnitude
    elif residual > 2:
        step = 5 * magnitude
    elif residual > 1:
        step = 2 * magnitude
    else:
        step = magnitude
    return max(step, min_step)


def _axis_ticks(data_max, target_count, min_step):
    """Возвращает (ticks, top) — список подписей сетки от 0 с шагом
    _nice_step(...) и верхнюю границу оси top. top всегда ОКРУГЛЯЕТСЯ ВВЕРХ
    до кратного шагу значения, не меньшего data_max — иначе при data_max,
    не кратном шагу, последняя точка данных оказывалась бы ЗА пределами
    последней подписанной линии сетки и вылезала за рамку графика."""
    step = _nice_step(data_max, target_count=target_count, min_step=min_step)
    n = max(1, math.ceil((data_max / step) - 1e-9))
    ticks = [i * step for i in range(n + 1)]
    return ticks, ticks[-1]


def _fmt_num(v):
    v = int(round(v))
    s = str(abs(v))
    parts = []
    while len(s) > 3:
        parts.insert(0, s[-3:])
        s = s[:-3]
    parts.insert(0, s)
    out = " ".join(parts)
    return f"-{out}" if v < 0 else out


def _esc(s):
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def render_svg(series, width=820, height=460):
    """Строит компактный SVG-график по данным extract_depth_series().
    Возвращает строку с готовой разметкой <svg>...</svg> либо None,
    если рисовать нечего (все серии пустые)."""
    if not series or not any(series.get(k) for k in ("plan", "fact", "npv_free")):
        return None

    all_x = [p[0] for k in ("plan", "fact", "npv_free") for p in (series.get(k) or [])]
    all_y = [p[1] for k in ("plan", "fact", "npv_free") for p in (series.get(k) or [])]
    if not all_x or not all_y:
        return None

    x_max = max(all_x + [1.0])
    y_max = max(all_y + [1.0])

    start_date = series.get("start_date")
    margin_left = 62
    margin_right = 18
    margin_top = 46 if start_date else 14
    margin_bottom = 78
    plot_w = width - margin_left - margin_right
    plot_h = height - margin_top - margin_bottom

    x_ticks, x_top = _axis_ticks(x_max, target_count=10, min_step=1)
    y_ticks, y_top = _axis_ticks(y_max, target_count=9, min_step=50)

    def px(x):
        return margin_left + (x / x_top if x_top else 0) * plot_w

    def py(y):
        return margin_top + (y / y_top if y_top else 0) * plot_h

    clip_id = "srDepthClip"
    parts = []
    parts.append(
        f'<svg viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" '
        f'role="img" aria-label="График глубина-день" style="width:100%;height:auto;font-family:Segoe UI, Arial, sans-serif;">'
    )
    parts.append(f'<rect x="0" y="0" width="{width}" height="{height}" fill="#ffffff"/>')
    parts.append(
        f'<defs><clipPath id="{clip_id}">'
        f'<rect x="{margin_left}" y="{margin_top}" width="{plot_w:.1f}" height="{plot_h:.1f}"/>'
        f'</clipPath></defs>'
    )

    # сетка + подписи оси Y (глубина, растёт вниз — как и в SVG-координатах)
    for gy in y_ticks:
        yy = py(gy)
        parts.append(
            f'<line x1="{margin_left}" y1="{yy:.1f}" x2="{width - margin_right}" y2="{yy:.1f}" '
            f'stroke="#e3e8ea" stroke-width="1"/>'
        )
        parts.append(
            f'<text x="{margin_left - 8}" y="{yy + 3:.1f}" text-anchor="end" font-size="10.5" fill="#6b7680">'
            f'{_fmt_num(gy)}</text>'
        )

    # сетка + подписи оси X (день, снизу)
    for gx in x_ticks:
        xx = px(gx)
        parts.append(
            f'<line x1="{xx:.1f}" y1="{margin_top}" x2="{xx:.1f}" y2="{height - margin_bottom}" '
            f'stroke="#eef1f2" stroke-width="1"/>'
        )
        parts.append(
            f'<text x="{xx:.1f}" y="{height - margin_bottom + 16}" text-anchor="middle" font-size="10.5" fill="#6b7680">'
            f'{_fmt_num(gx)}</text>'
        )

    # верхняя ось с датами (как в самом Excel) — та же сетка по X, только
    # подпись не числом дня, а календарной датой (день 0 = start_date)
    if start_date:
        for gx in x_ticks:
            xx = px(gx)
            label = (start_date + datetime.timedelta(days=gx)).strftime("%d.%m.%Y")
            parts.append(
                f'<line x1="{xx:.1f}" y1="{margin_top - 4:.1f}" x2="{xx:.1f}" y2="{margin_top:.1f}" '
                f'stroke="#c9d1d5" stroke-width="1"/>'
            )
            parts.append(
                f'<text x="0" y="0" text-anchor="start" font-size="9.5" fill="#8a9296" '
                f'transform="translate({xx:.1f},{margin_top - 6:.1f}) rotate(-55)">{label}</text>'
            )

    # рамка области построения
    parts.append(
        f'<rect x="{margin_left}" y="{margin_top}" width="{plot_w:.1f}" height="{plot_h:.1f}" '
        f'fill="none" stroke="#c9d1d5" stroke-width="1"/>'
    )

    # подписи осей
    parts.append(
        f'<text x="{margin_left + plot_w / 2:.1f}" y="{height - 30}" text-anchor="middle" '
        f'font-size="11" fill="#6b7680">День</text>'
    )
    parts.append(
        f'<text x="14" y="{margin_top + plot_h / 2:.1f}" text-anchor="middle" font-size="11" fill="#6b7680" '
        f'transform="rotate(-90 14 {margin_top + plot_h / 2:.1f})">Глубина, м</text>'
    )

    # сами серии — обёрнуты в clip-path на область построения: даже если
    # у какой-то точки координата случайно окажется за пределами осей (не
    # должно происходить благодаря _axis_ticks, но на всякий случай), линия
    # обрежется по рамке графика, а не "вылезет" за неё
    parts.append(f'<g clip-path="url(#{clip_id})">')
    legend_items = []
    for key, label, color, dashed in SERIES:
        pts = series.get(key) or []
        dash_attr = ' stroke-dasharray="5,4"' if dashed else ""
        if len(pts) >= 2:
            pts_sorted = sorted(pts, key=lambda p: p[0])
            path = " ".join(f"{px(x):.1f},{py(y):.1f}" for x, y in pts_sorted)
            parts.append(
                f'<polyline points="{path}" fill="none" stroke="{color}" stroke-width="2.2" '
                f'stroke-linejoin="round" stroke-linecap="round"{dash_attr}/>'
            )
        elif len(pts) == 1:
            x, y = pts[0]
            parts.append(f'<circle cx="{px(x):.1f}" cy="{py(y):.1f}" r="3" fill="{color}"/>')
        legend_items.append((label, color, dashed))
    parts.append('</g>')

    # легенда снизу
    legend_y = height - 12
    lx = margin_left
    for label, color, dashed in legend_items:
        dash_attr = ' stroke-dasharray="4,3"' if dashed else ""
        parts.append(
            f'<line x1="{lx}" y1="{legend_y - 4}" x2="{lx + 20}" y2="{legend_y - 4}" '
            f'stroke="{color}" stroke-width="2.4"{dash_attr}/>'
        )
        label_esc = _esc(label)
        text_x = lx + 26
        parts.append(
            f'<text x="{text_x}" y="{legend_y}" font-size="10.5" fill="#3a4046">{label_esc}</text>'
        )
        lx = text_x + 9 * len(label) + 26

    parts.append("</svg>")
    return "".join(parts)


def build_depth_chart_svg(wb):
    """Удобная обёртка: читает серии и сразу строит SVG. Возвращает
    (svg_str, error) — при неудаче svg_str is None, а error — короткое
    объяснение причины (для журнала программы)."""
    try:
        series = extract_depth_series(wb)
        if not series:
            return None, "на листе «%s» нет данных для графика (или самого листа нет в файле)" % SHEET_NAME
        svg = render_svg(series)
        if not svg:
            return None, "не удалось построить график: пустой набор точек"
        return svg, None
    except Exception as e:
        return None, "ошибка построения графика средствами программы: %s" % e
