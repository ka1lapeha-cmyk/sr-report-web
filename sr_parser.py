\
# -*- coding: utf-8 -*-
"""
sr_parser.py
Разбор файлов "Суточный рапорт" (СР) по бурению (формат ООО "ИНК") и подготовка
данных для HTML-отчёта.

Файлы СР — это .xlsx-книги с фиксированным набором листов, из которых нам нужны:
  - "Суточный отчет"  — основной лист рапорта;
  - "НПВ"             — непроизводительное время по ответственным сторонам.

Расположение большинства ячеек в верхней части листа "Суточный отчет"
(шапка, блок "Конструкция", блок "Забой/проходка") — фиксированное и не
меняется от скважины к скважине. А вот блоки "КНБК и БТ", "Долото" и
"Суточные операции" расположены НИЖЕ таблицы КНБК, длина которой отличается
от рапорта к рапорту — поэтому их положение определяется не номером строки,
а поиском текстовой метки-заголовка ("Долото", "Суточные операции" и т.д.)
и последующим разбором ячеек относительно найденной строки.

Внимание: в этих файлах Excel-атрибут "размерность листа" (dimension) часто
не соответствует реальному количеству строк (типичная проблема реальных
корпоративных книг), поэтому весь построчный поиск меток ведётся с явно
заданной верхней границей (SCAN_MAX_ROW), а не через ws.max_row.
"""

import datetime
import glob
import os
import re
import shutil
import subprocess
import tempfile
import zipfile

import openpyxl
from openpyxl.utils import get_column_letter

import depth_chart
import stage_classifier
import vzd_reference

MAIN_SHEET = "Суточный отчет"
NPV_SHEET = "НПВ"
TEP_SHEET = "ТЭП"
PROGRAM_SHEET = "Программа работ"
NARABOTKA_SHEET = "Наработка"
CHRONOLOGY_SHEET = "Хронология работ"
CHRONOLOGY_MAX_ROW = 20000

SCAN_MAX_ROW = 400          # безопасная верхняя граница построчного поиска меток
BIT_SEARCH_RANGE = 15        # сколько строк после заголовка блока "Долото" просматривать
OPS_SEARCH_RANGE = 250        # сколько строк после "Суточные операции" просматривать в поиске "Итого, часов"
NPV_SEARCH_RANGE = 200
NARABOTKA_SEARCH_RANGE = 600  # сколько строк листа "Наработка" просматривать в поиске блоков КНБК

# Известные производители/модели ВЗД (винтовых забойных двигателей) — используются
# только как ДОПОЛНИТЕЛЬНАЯ подсказка при распознавании (первый и главный признак —
# позиция сразу после долота в таблице "КНБК и БТ", см. extract_narabotka).
VZD_NAME_MARKERS = ("ВЗД", "МВР", "ДРУ", "7LZ", "7lz")
VZD_KNOWN_MAKERS = ("Радиус-Сервис", "Ляньхэ", "Титан")

# Осциллятор (инструмент для снижения сил трения/затяжек при бурении) —
# встречается в таблице КНБК под разными сокращениями/названиями брендов.
# "ОСЦ" — безопасно искать как часть слова (ловит "Осциллятор", "Осц." и
# т.п.), а короткое "ОС" — только как отдельное "слово" в названии, иначе
# слишком велик риск случайного совпадения с не относящимся к делу текстом.
OSC_SUBSTR_MARKERS = ("ОСЦ", "АГТ", "АГАР", "AGT", "AGAR", "AGR", "АГР")
OSC_WORD_MARKERS = ("ОС",)

# СБТ (стальные бурильные трубы) — вкладка "Инструмент": фиксированный ресурс
# (не подбирается по справочнику, как у ВЗД/осциллятора — задан пользователем
# напрямую), актуален только для секции "Хвостовик" (см. extract_narabotka).
SBT_RESOURCE_HOURS = 75

# Параметры промывочной жидкости (блок листа "Суточный отчет" под "Суточными
# операциями") — список показателей, которые ищем в таблице этого блока.
# key — внутренний идентификатор, label — точный текст подписи столбца в
# строке заголовка блока (см. extract_mud_params), unit_fallback — единица
# измерения на случай, если строку единиц не удалось разобрать.
MUD_PARAM_COLUMNS = [
    ("depth", "Глубина замера", "м"),
    ("density", "Плотность", "г/см³"),
    ("viscosity", "Усл. Вязкость", "сек"),
    ("pv", "ПВ", "сП"),
    ("dns", "ДНС", "фунт/100фт²"),
    ("water_loss", "Водоотдача", "мл/30 мин"),
    ("so4", "SO4", "мг/л"),
    ("solids", "Твердая фаза", "%"),
    ("lubricity", "Смазка", "%"),
    ("ph", "рН", ""),
    ("caco3", "CaCO3", "кг/м³"),
    ("v_well", "Vскв.", "м³"),
    ("v_surf", "Vповерх.", "м³"),
    ("v_tech_water", "V тех. воды", "м³"),
]

ARCHIVE_EXTS = (".zip", ".rar")


class ParseError(Exception):
    pass


def find_xlsx_files(folder):
    """Рекурсивно находит все .xlsx файлы в папке, исключая временные файлы Excel (~$...)."""
    pattern = os.path.join(folder, "**", "*.xlsx")
    files = []
    for path in glob.glob(pattern, recursive=True):
        name = os.path.basename(path)
        if name.startswith("~$"):
            continue
        files.append(path)
    return sorted(files)


# ---------------------------------------------------------------------------
# Источник данных: папка, отдельный файл .xlsx, либо архив .zip/.rar
# ---------------------------------------------------------------------------

def _extract_zip(archive_path, dest_dir):
    with zipfile.ZipFile(archive_path) as z:
        z.extractall(dest_dir)


def _find_rar_tool():
    """Ищет во внешний распаковщик RAR, доступный на компьютере пользователя
    (сам Python не умеет распаковывать .rar без внешней утилиты)."""
    for exe in ("unrar", "unrar.exe"):
        if shutil.which(exe):
            return ("unrar", exe)
    for exe in ("7z", "7za", "7z.exe", "7za.exe"):
        if shutil.which(exe):
            return ("7z", exe)
    for exe in ("unar", "unar.exe"):
        if shutil.which(exe):
            return ("unar", exe)
    return None


def _extract_rar(archive_path, dest_dir):
    tool = _find_rar_tool()
    if not tool:
        raise ParseError(
            "Для чтения .rar-архивов нужна внешняя программа-распаковщик "
            "('unrar', '7-Zip' или 'unar'), которая не найдена на этом компьютере.\n"
            "Установите одну из них (например, unrar: 'sudo apt install unrar' на Linux, "
            "'brew install unrar' на macOS, или 7-Zip на Windows) и повторите, "
            "либо переупакуйте архив в формат .zip."
        )
    kind, exe = tool
    if kind == "unrar":
        cmd = [exe, "x", "-y", archive_path, dest_dir + os.sep]
    elif kind == "unar":
        cmd = [exe, "-o", dest_dir, "-f", archive_path]
    else:  # 7z / 7za
        cmd = [exe, "x", f"-o{dest_dir}", "-y", archive_path]
    try:
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except OSError as e:
        raise ParseError(f"Не удалось запустить распаковщик RAR ({exe}): {e}")
    if result.returncode != 0:
        err = result.stderr.decode("utf-8", errors="ignore").strip() or result.stdout.decode("utf-8", errors="ignore").strip()
        raise ParseError(f"Не удалось распаковать RAR-архив (код {result.returncode}): {err[:400]}")


def resolve_xlsx_sources(input_path):
    """Принимает путь, выбранный пользователем — папку, отдельный файл .xlsx
    или архив .zip/.rar с файлами СР — и возвращает кортеж:
      (список путей к .xlsx файлам, список временных папок для последующей очистки).
    Если внутри архива есть вложенные архивы или папки — файлы .xlsx всё равно
    будут найдены рекурсивно."""
    cleanup_dirs = []

    if not os.path.exists(input_path):
        raise ParseError("Указанный путь не найден.")

    if os.path.isdir(input_path):
        return find_xlsx_files(input_path), cleanup_dirs

    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".xlsx":
        if os.path.basename(input_path).startswith("~$"):
            raise ParseError("Это временный файл Excel, а не рапорт СР.")
        return [input_path], cleanup_dirs

    if ext in ARCHIVE_EXTS:
        tmp_dir = tempfile.mkdtemp(prefix="sr_report_")
        cleanup_dirs.append(tmp_dir)
        if ext == ".zip":
            _extract_zip(input_path, tmp_dir)
        else:
            _extract_rar(input_path, tmp_dir)
        files = find_xlsx_files(tmp_dir)
        if not files:
            raise ParseError("В архиве не найдено ни одного файла .xlsx.")
        return files, cleanup_dirs

    raise ParseError(
        f"Неподдерживаемый тип файла: «{ext or 'без расширения'}». "
        "Выберите папку с файлами СР, отдельный файл .xlsx, либо архив .zip/.rar."
    )


class SheetGrid:
    """Материализует значения листа в память за один проход (iter_rows),
    вместо произвольного доступа ws.cell(row=, column=) — в режиме
    read_only у openpyxl такой произвольный доступ на некоторых книгах
    оказывается на порядки медленнее одного последовательного прохода,
    особенно при большом количестве точечных обращений (как в этом модуле:
    десятки/сотни поисков меток на файл)."""

    __slots__ = ("_rows",)

    def __init__(self, ws, max_row=SCAN_MAX_ROW):
        self._rows = list(ws.iter_rows(min_row=1, max_row=max_row, values_only=True))

    def cell(self, row, col):
        r = row - 1
        if r < 0 or r >= len(self._rows):
            return None
        line = self._rows[r]
        c = col - 1
        if c < 0 or c >= len(line):
            return None
        return line[c]

    def row(self, row_idx):
        r = row_idx - 1
        if r < 0 or r >= len(self._rows):
            return ()
        return self._rows[r]

    @property
    def nrows(self):
        return len(self._rows)


def _cell(grid, col, row):
    return grid.cell(row, col)


def _find_row_by_label(grid, col_idx, label, max_row=SCAN_MAX_ROW):
    """Ищет строку, где ячейка в столбце col_idx точно равна label."""
    limit = min(max_row, grid.nrows)
    for r in range(1, limit + 1):
        if grid.cell(r, col_idx) == label:
            return r
    return None


def _find_col_in_row(grid, row_idx, label, max_col=40):
    """Ищет столбец в указанной строке, где значение ячейки точно равно label."""
    row_vals = grid.row(row_idx)
    for c, v in enumerate(row_vals[:max_col], start=1):
        if v == label:
            return c
    return None


def _find_col_startswith(grid, row_idx, prefix, max_col=40):
    """Как _find_col_in_row, но ищет ячейку, ТЕКСТ которой начинается с prefix
    (после strip) — используется там, где в реальных файлах встречаются
    небольшие вариации в подписи (лишние пробелы/точки на конце)."""
    row_vals = grid.row(row_idx)
    for c, v in enumerate(row_vals[:max_col], start=1):
        if isinstance(v, str) and v.strip().startswith(prefix):
            return c
    return None


def _find_all_rows_startswith(grid, col_idx, prefix, max_row=SCAN_MAX_ROW):
    """Возвращает список номеров строк, где ячейка в столбце col_idx —
    строка, начинающаяся с prefix (после strip)."""
    limit = min(max_row, grid.nrows)
    rows = []
    for r in range(1, limit + 1):
        v = grid.cell(r, col_idx)
        if isinstance(v, str) and v.strip().startswith(prefix):
            rows.append(r)
    return rows


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _clean_str(v):
    if v is None:
        return ""
    return str(v).strip()


# ---------------------------------------------------------------------------
# Извлечение блоков листа "Суточный отчет"
# ---------------------------------------------------------------------------

def extract_header(ws):
    """Шапка рапорта — фиксированные адреса ячеек (строки 2-16).

    Поле "Телефон: мобильный / IP" в исходнике на самом деле — ДВЕ ячейки
    в одной строке: мобильный номер (столбец L) и отдельно стационарный
    номер/IP (столбец O), одинаково для куратора, супервайзера и мастера
    (curator_ip/supervisor_ip/driller_ip) — проверено по нескольким
    реальным файлам, столбец O заполнен во всех случаях."""
    return {
        "report_date": _cell(ws, 4, 3),     # D3
        "report_no": _cell(ws, 7, 3),       # G3
        "customer": _cell(ws, 6, 5),        # F5
        "license_area": _cell(ws, 6, 6),    # F6
        "field": _cell(ws, 6, 7),           # F7
        "pad": _cell(ws, 6, 8),             # F8
        "well_no": _cell(ws, 6, 9),         # F9
        "well_sequence": _cell(ws, 6, 10),  # F10
        "well_purpose": _cell(ws, 6, 11),   # F11
        "well_type": _cell(ws, 6, 12),      # F12
        "work_start": _cell(ws, 6, 13),     # F13
        "norm_days": _cell(ws, 6, 14),      # F14
        "forecast_end": _cell(ws, 6, 15),   # F15
        "lag_days": _cell(ws, 6, 16),       # F16
        "curator": _cell(ws, 12, 5),         # L5
        "curator_phone": _cell(ws, 12, 6),   # L6
        "curator_ip": _cell(ws, 15, 6),      # O6 (стационарный телефон/IP, рядом с мобильным)
        "curator_email": _cell(ws, 12, 7),   # L7
        "supervisor": _cell(ws, 12, 8),      # L8
        "supervisor_phone": _cell(ws, 12, 9),   # L9
        "supervisor_ip": _cell(ws, 15, 9),      # O9
        "supervisor_email": _cell(ws, 12, 10),  # L10
        "rig": _cell(ws, 12, 11),            # L11
        "brigade": _cell(ws, 12, 12),        # L12
        "driller": _cell(ws, 12, 13),        # L13
        "driller_phone": _cell(ws, 12, 14),  # L14
        "driller_ip": _cell(ws, 15, 14),     # O14
        "driller_email": _cell(ws, 12, 15),  # L15
        "days_no_incidents": _cell(ws, 12, 16),  # L16
    }


def extract_construction(ws):
    """Таблица 'Конструкция' (план/факт) в правом верхнем блоке рапорта.
    Заголовок 'Конструкция' находится в столбце R (18), строка 5.
    Заголовки колонок — строка 6 (R6..Z6), данные — с 7-й строки, пока
    столбец R (наименование секции) не пуст."""
    anchor = _find_row_by_label(ws, 18, "Конструкция", max_row=20)
    rows = []
    if not anchor:
        return rows
    header_row = anchor + 1
    r = header_row + 1
    while r < header_row + 30:
        name = _cell(ws, 18, r)  # R
        if name is None or _clean_str(name) == "":
            break
        rows.append({
            "name": name,
            "diameter": _cell(ws, 20, r),  # T
            "plan": _cell(ws, 22, r),      # V
            "fact": _cell(ws, 26, r),      # Z
        })
        r += 1
    return rows


def extract_depth_block(ws):
    """Блок 'Забой на 24:00 / Суточная проходка / НПВ за сутки / Забой на 06:00 /
    Планируемые работы на сутки'. Расположение фиксировано (строки 49-61)."""
    return {
        "depth_24": _cell(ws, 5, 49),          # E49
        "work_at_24": _cell(ws, 13, 49),       # M49
        "daily_footage": _cell(ws, 5, 52),     # E52
        "well_footage": _cell(ws, 13, 52),     # M52
        "daily_npv": _cell(ws, 5, 55),         # E55
        "cum_npv": _cell(ws, 13, 55),          # M55
        "depth_06": _cell(ws, 5, 58),          # E58
        "work_at_06": _cell(ws, 13, 58),       # M58
        "planned_work": _cell(ws, 5, 61),      # E61
    }


def extract_bit(ws):
    """Блок 'Долото'. Заголовок ищется по точному совпадению 'Долото' в столбце B.
    Строка заголовков колонок = anchor+1, столбцы фиксированы (проверено на
    всех образцах): B=Ø, C=тип, F=модель, H=изготовитель, R=проходка за сутки,
    U=мех.скорость. Столбцы интервала бурения ('от'/'до') ищутся отдельно,
    т.к. они относятся к объединённым ячейкам."""
    anchor = _find_row_by_label(ws, 2, "Долото")
    if not anchor:
        return None
    header_row = anchor + 1

    from_col = to_col = None
    for rr in range(header_row, header_row + 4):
        for c in range(1, 35):
            v = _cell(ws, c, rr)
            if v == "от":
                from_col = c
            elif v == "до":
                to_col = c

    data_rows = []
    r = header_row + 1
    misses = 0
    while r < header_row + BIT_SEARCH_RANGE:
        v = _cell(ws, 2, r)  # столбец B — Ø долота (число)
        if _is_number(v):
            data_rows.append(r)
            misses = 0
        else:
            misses += 1
            if data_rows and misses > 1:
                break
        r += 1

    if not data_rows:
        return None

    last = data_rows[-1]  # последняя запись = текущее долото
    return {
        "diameter": _cell(ws, 2, last),                       # B
        "type": _cell(ws, 3, last),                           # C
        "model": _cell(ws, 6, last),                          # F
        "maker": _cell(ws, 8, last),                          # H
        "interval_from": _cell(ws, from_col, last) if from_col else None,
        "interval_to": _cell(ws, to_col, last) if to_col else None,
        "daily_footage": _cell(ws, 18, last),                 # R
        "rop": _cell(ws, 21, last),                           # U
    }


def extract_ops(ws):
    """Блок 'Суточные операции'. Заголовок ищется по точному совпадению
    в столбце B. Столбцы 'Категория'/'Продолж.'/'Ответственный' относятся
    к учёту НПВ внутри операций и определяются поиском по тексту в строке
    заголовка+2 (т.к. заголовок таблицы двухуровневый). 'Комментарий' —
    отдельный столбец верхнего уровня (строка заголовка+1) с подробным
    текстовым описанием, что происходило в этом интервале времени."""
    anchor = _find_row_by_label(ws, 2, "Суточные операции")
    if not anchor:
        return []
    sub_header_row = anchor + 2
    cat_col = _find_col_in_row(ws, sub_header_row, "Категория")
    dur_col = _find_col_in_row(ws, sub_header_row, "Продолж.")
    resp_col = _find_col_in_row(ws, sub_header_row, "Ответственный")
    comment_col = _find_col_in_row(ws, anchor + 1, "Комментарий")

    if not cat_col:
        return []

    start_col = _find_col_in_row(ws, anchor + 1, "Время операции") or 2
    end_col = start_col + 1

    events = []
    r = anchor + 3
    limit = anchor + OPS_SEARCH_RANGE
    while r < limit:
        b_val = _cell(ws, 2, r)
        if b_val == "Итого, часов":
            break
        if b_val is None and _cell(ws, 3, r) is None:
            # пустая строка — таблица могла закончиться раньше "Итого"
            # (перестрахуемся и остановимся после пары подряд пустых строк)
            r += 1
            if _cell(ws, 2, r) is None:
                break
            continue
        cat = _cell(ws, cat_col, r) if cat_col else None
        if cat is not None and _clean_str(cat) != "":
            events.append({
                "time_start": _cell(ws, start_col, r),
                "time_end": _cell(ws, end_col, r),
                "category": cat,
                "duration": _cell(ws, dur_col, r) if dur_col else None,
                "responsible": _cell(ws, resp_col, r) if resp_col else None,
                "comment": _cell(ws, comment_col, r) if comment_col else None,
            })
        r += 1
    return events


_MUD_RANGE_RE = re.compile(r"^([\d]+(?:[.,]\d+)?)\s*-\s*([\d]+(?:[.,]\d+)?)$")
_MUD_BOUND_RE = re.compile(r"^(<=|>=|≤|≥|<|>)\s*([\d]+(?:[.,]\d+)?)$")


def _parse_ru_float(s):
    try:
        return float(_clean_str(s).replace(",", "."))
    except ValueError:
        return None


def _parse_mud_plan_bounds(text):
    """Разбирает текст плановой границы параметра промывочной жидкости
    (строка 'План' блока 'Параметры промывочной жидкости' листа 'Суточный
    отчет') в границы (мин, макс) — любая из них может быть None (значит,
    с этой стороны ограничения нет). В реальных файлах встречаются только
    форматы 'a - b' (диапазон, напр. '1,32 - 1,38'), '≤ b'/'<= b' (только
    верхняя граница), '≥ a'/'>= a' (только нижняя), а также просто '< b'/'> a'.
    Если текст не в одном из этих форматов (например, план не задан или это
    не число) — возвращает None, и тогда план/факт для этого параметра просто
    показывается без подсветки."""
    s = _clean_str(text)
    if not s:
        return None
    m = _MUD_RANGE_RE.match(s)
    if m:
        lo, hi = _parse_ru_float(m.group(1)), _parse_ru_float(m.group(2))
        if lo is None or hi is None:
            return None
        return (lo, hi)
    m = _MUD_BOUND_RE.match(s)
    if m:
        op, num = m.group(1), _parse_ru_float(m.group(2))
        if num is None:
            return None
        if op in ("<", "<=", "≤"):
            return (None, num)
        if op in (">", ">=", "≥"):
            return (num, None)
    return None


def _mud_value_status(value, bounds):
    """'bad', если факт выходит за плановую границу, 'ok' — если в пределах,
    None — если сравнивать не с чем (нет плана или факт не число)."""
    if not _is_number(value) or not bounds:
        return None
    lo, hi = bounds
    if lo is not None and value < lo:
        return "bad"
    if hi is not None and value > hi:
        return "bad"
    return "ok"


def extract_mud_params(ws):
    """Блок 'Параметры промывочной жидкости' листа 'Суточный отчет' (сразу
    под блоком 'Суточные операции') — по указанию пользователя источник для
    сравнения план/факт параметров бурового раствора (плотность, вязкость,
    водоотдача, СНС, pH и др.).

    Структура блока (шаблон, фиксирован по вертикали относительно заголовка
    'Параметры промывочной жидкости', но само расположение блока по листу
    "плавающее", т.к. зависит от длины таблицы 'Суточные операции' над ним):
      заголовок+1 — строка с подписями столбцов ('Плотность', 'Усл. Вязкость'
                    и т.д.; у 'СНС' подпись одна на 2 столбца);
      заголовок+2 — подзаголовок для двухколоночных параметров (сейчас
                    только 'СНС': '10сек'/'10мин');
      заголовок+3 — единицы измерения;
      далее — 0 или 1 строка 'План' (столбец B) с плановым диапазоном/
                    границей по каждому параметру, затем 1+ строк 'Факт' —
                    по одной на каждый замер за сутки (обычно 3: конец
                    предыдущей смены/день/ночь, время — столбец 'Время
                    замера'). Строки 'План' в файле может не быть вовсе
                    (тогда сравнивать не с чем — показываем только факт).

    Возвращает None, если блок не найден. Иначе — словарь:
      {'mud_type', 'system' — тип/система бурового раствора (текст, из
         первой строки 'Факт', т.к. в 'План' эти поля дублируются так же);
       'params': [{'key','label','unit','plan_text','bounds'}, ...] —
         только параметры, для которых нашёлся столбец в файле;
       'facts': [{'time', 'values': {key: значение}, 'statuses': {key: 'ok'/
         'bad'/None}}, ...] — по одной записи на строку 'Факт'}."""
    anchor = _find_row_by_label(ws, 2, "Параметры промывочной жидкости")
    if not anchor:
        return None
    header_row = anchor + 1
    sub_header_row = header_row + 1
    units_row = header_row + 2

    type_col = _find_col_in_row(ws, header_row, "Тип раствора")
    system_col = _find_col_in_row(ws, header_row, "Система")
    time_col = _find_col_in_row(ws, header_row, "Время замера")

    cols = {}
    for key, label, unit_fallback in MUD_PARAM_COLUMNS:
        col = _find_col_in_row(ws, header_row, label)
        if col:
            cols[key] = col

    sns10s_col = _find_col_in_row(ws, sub_header_row, "10сек")
    sns10m_col = _find_col_in_row(ws, sub_header_row, "10мин")
    if sns10s_col:
        cols["sns10s"] = sns10s_col
    if sns10m_col:
        cols["sns10m"] = sns10m_col

    if not cols:
        return None

    labels_by_key = {key: label for key, label, _ in MUD_PARAM_COLUMNS}
    labels_by_key["sns10s"] = "СНС 10 сек"
    labels_by_key["sns10m"] = "СНС 10 мин"
    unit_fallback_by_key = {key: u for key, _, u in MUD_PARAM_COLUMNS}
    unit_fallback_by_key["sns10s"] = "фунт/100фт²"
    unit_fallback_by_key["sns10m"] = "фунт/100фт²"

    # План — максимум одна строка, ищем в пределах небольшого диапазона под
    # единицами (перестраховка: если её нет, plan_row останется None).
    plan_row = None
    fact_rows = []
    r = units_row + 1
    limit = units_row + 1 + 30
    blanks_after_data = 0
    while r < limit:
        b_val = _clean_str(_cell(ws, 2, r))
        if b_val == "План" and plan_row is None:
            plan_row = r
            blanks_after_data = 0
        elif b_val == "Факт":
            fact_rows.append(r)
            blanks_after_data = 0
        elif fact_rows or plan_row:
            blanks_after_data += 1
            if blanks_after_data >= 2:
                break
        r += 1

    params = []
    for key, col in cols.items():
        plan_text = _clean_str(_cell(ws, col, plan_row)) if plan_row else ""
        unit = _clean_str(_cell(ws, col, units_row)) or unit_fallback_by_key.get(key, "")
        params.append({
            "key": key,
            "label": labels_by_key.get(key, key),
            "unit": unit,
            "plan_text": plan_text,
            "bounds": _parse_mud_plan_bounds(plan_text) if plan_text else None,
        })
    # Сохраняем порядок как в MUD_PARAM_COLUMNS (+ СНС сразу после ДНС).
    order = [k for k, _, _ in MUD_PARAM_COLUMNS]
    if "sns10s" in cols:
        order.insert(order.index("dns") + 1 if "dns" in order else len(order), "sns10s")
    if "sns10m" in cols:
        order.insert(order.index("sns10s") + 1 if "sns10s" in order else len(order), "sns10m")
    params.sort(key=lambda p: order.index(p["key"]) if p["key"] in order else 999)

    facts = []
    mud_type = None
    system = None
    for fr in fact_rows:
        values = {}
        statuses = {}
        for p in params:
            v = _cell(ws, cols[p["key"]], fr)
            values[p["key"]] = v
            statuses[p["key"]] = _mud_value_status(v, p["bounds"])
        facts.append({
            "time": _cell(ws, time_col, fr) if time_col else None,
            "values": values,
            "statuses": statuses,
        })
        if mud_type is None and type_col:
            mud_type = _clean_str(_cell(ws, type_col, fr)) or None
        if system is None and system_col:
            system = _clean_str(_cell(ws, system_col, fr)) or None

    return {
        "mud_type": mud_type,
        "system": system,
        "params": params,
        "facts": facts,
    }


def extract_npv(wb):
    """Лист 'НПВ' — непроизводительное время по ответственным сторонам.
    Строка 6 — агрегированный итог (используем как общий итог по скважине),
    строки 7+ — разбивка по организациям, пока столбец B (организация) не пуст."""
    if NPV_SHEET not in wb.sheetnames:
        return {"total_hours": None, "by_org": []}
    ws = SheetGrid(wb[NPV_SHEET], max_row=7 + NPV_SEARCH_RANGE)
    total_hours = _cell(ws, 4, 6)  # D6 — суммарный итог по всем сторонам
    by_org = []
    r = 7
    while r < 7 + NPV_SEARCH_RANGE:
        org = _cell(ws, 2, r)  # B
        if org is None or _clean_str(org) == "":
            break
        service = _cell(ws, 3, r)   # C
        hours = _cell(ws, 4, r)     # D
        if hours or _clean_str(org):
            by_org.append({"org": org, "service": service, "hours": hours})
        r += 1
    return {"total_hours": total_hours, "by_org": by_org}


def extract_commercial_speed(wb):
    """Коммерческая скорость (план/факт, м/ст.мес) с листа 'Программа работ'.
    Заголовок 'Коммерческая скорость...' ищется по началу строки в первых
    строках листа, план/факт — по точным подписям 'План' и 'Факт' в той же
    строке (не 'План на текущей операции' — она стоит между ними)."""
    if PROGRAM_SHEET not in wb.sheetnames:
        return None
    ws = SheetGrid(wb[PROGRAM_SHEET], max_row=15)
    anchor = None
    for r in range(1, 15):
        for c in range(1, 30):
            v = _cell(ws, c, r)
            if isinstance(v, str) and v.strip().startswith("Коммерческая скорость"):
                anchor = (r, c)
                break
        if anchor:
            break
    if not anchor:
        return None
    r, start_c = anchor
    plan_col = fact_col = None
    for c in range(start_c, start_c + 20):
        v = _cell(ws, c, r)
        if v == "План" and plan_col is None:
            plan_col = c
        elif v == "Факт" and fact_col is None and plan_col is not None:
            fact_col = c
    if plan_col is None:
        return None
    return {
        "plan": _cell(ws, plan_col, r + 1),
        "fact": _cell(ws, fact_col, r + 1) if fact_col else None,
    }


def extract_stage_deviation(wb):
    """Разбивка отклонения (план/факт, часы) по этапам работ с листа 'ТЭП'
    (таблица 'Распределение времени по этапам работ'). Возвращает список
    {'stage':..., 'plan_h', 'fact_h', 'deviation_h'} для этапов, где есть
    хоть какие-то данные."""
    if TEP_SHEET not in wb.sheetnames:
        return []
    ws = SheetGrid(wb[TEP_SHEET], max_row=100)
    anchor = _find_row_by_label(ws, 2, "Распределение времени по этапам работ", max_row=60)
    if not anchor:
        return []
    header_row = anchor + 1
    stages = []
    r = header_row + 1
    while r < header_row + 25:
        stage = _cell(ws, 3, r)   # C
        num = _cell(ws, 2, r)     # B
        if stage is None and num is None:
            break
        dev_h = _cell(ws, 10, r)  # J - Отклонение (ч)
        if stage is not None and _is_number(dev_h) and dev_h != 0:
            stages.append({
                "stage": stage,
                "plan_h": _cell(ws, 6, r),   # F
                "fact_h": _cell(ws, 7, r),   # G
                "deviation_h": dev_h,
            })
        r += 1
    return stages


def extract_current_stage(wb):
    """Лист 'Хронология работ' — 'Этап' для сводной таблицы, определяется по
    самой нижней заполненной строке таблицы (последняя по времени запись
    хронологии = текущее состояние скважины на момент рапорта).

    Этап определяется классификатором stage_classifier.classify_stage() по
    паре столбцов (Вид работ, Операция, Детализация операции) той же строки
    — см. пояснение в stage_classifier.py про ручные правила и справочник.
    Если классификатор не нашёл совпадения (например, старый формат листа
    без этих столбцов, либо сочетание, отсутствующее и в ручных правилах, и
    в справочнике), используется как запасной вариант сырой текст ячейки
    'Этап работ' той же строки. Отдельно классификатор может вернуть не
    этап, а ПРЕДУПРЕЖДЕНИЕ (для видов работ, где по одной лишь пустой
    "Операции" однозначно определить этап нельзя) — тогда сырой текст не
    используется, чтобы не выдавать гадательный этап за точный.

    Возвращает словарь {"text": этап_или_None, "warning": подсказка_или_None}
    либо None, если строку хронологии определить вообще не удалось (лист
    отсутствует, нестандартная структура и т.п.) — не прерывая разбор файла."""
    if CHRONOLOGY_SHEET not in wb.sheetnames:
        return None
    try:
        ws = SheetGrid(wb[CHRONOLOGY_SHEET], max_row=CHRONOLOGY_MAX_ROW)
        # строка заголовка таблицы — та, где в столбце C стоит 'Начало'
        # (в столбце B выше есть ещё одна строка с меткой 'Дата' — это шапка
        # листа с датой отчёта, не заголовок самой таблицы хронологии).
        header_row = _find_row_by_label(ws, 3, "Начало", max_row=20)
        if not header_row:
            return None
        stage_col = _find_col_in_row(ws, header_row, "Этап работ", max_col=30)
        vid_col = _find_col_in_row(ws, header_row, "Вид работ", max_col=30)
        op_col = _find_col_in_row(ws, header_row, "Операция", max_col=30)
        det_col = _find_col_in_row(ws, header_row, "Детализация операции", max_col=30)
        if not stage_col and not vid_col:
            return None
        last_stage = None
        last_vid = last_op = last_det = None
        for r in range(header_row + 1, ws.nrows + 1):
            date_v = ws.cell(r, 2)
            time_v = ws.cell(r, 3)
            if date_v is None and time_v is None:
                continue
            stage_v = ws.cell(r, stage_col) if stage_col else None
            vid_v = ws.cell(r, vid_col) if vid_col else None
            det_v = ws.cell(r, det_col) if det_col else None
            op_v = ws.cell(r, op_col) if op_col else None
            has_any = any(
                v is not None and _clean_str(v) != ""
                for v in (stage_v, vid_v, op_v, det_v)
            )
            if has_any:
                last_stage = stage_v
                last_vid, last_op, last_det = vid_v, op_v, det_v
        raw_stage = last_stage if (last_stage is not None and _clean_str(last_stage) != "") else None
        classified, warning = stage_classifier.classify_stage(last_vid, last_op, last_det)
        if warning:
            return {"text": None, "warning": warning}
        if classified:
            return {"text": classified, "warning": None}
        if raw_stage:
            return {"text": raw_stage, "warning": None}
        return None
    except Exception:
        return None


def _parse_component_name(name):
    """Разбирает текст наименования элемента КНБК вида 'ВЗД 7LZ244х7.0-6,0
    (Ляньхэ)' или 'МВР-121ТУ(Титан)' на модель и производителя (производитель
    — в скобках в конце названия, см. пояснение пользователя). Используется
    и для ВЗД, и для осциллятора — формат подписи в таблице у них общий."""
    name = _clean_str(name)
    name = re.sub(r"\s+", " ", name)  # в названиях иногда встречаются переносы строк
    m = re.match(r"^(.*?)\(([^)]+)\)\s*$", name)
    if m:
        model, maker = m.group(1).strip(), m.group(2).strip()
    else:
        model, maker = name, None
    model = re.sub(r"^ВЗД\s+", "", model).strip()
    return model, maker


def _normalize_bha_name(s):
    """Нормализация названия элемента КНБК для сопоставления одного и того
    же физического элемента между листом 'Наработка' (где к названию может
    быть приписан суффикс '(производитель/сервисная компания)') и таблицей
    'КНБК и БТ' листа 'Суточный отчет' (где такого суффикса нет, но зато
    есть колонка 'Наруж. Ø' с диаметром) — см. extract_bha_diameters()."""
    s = _clean_str(s)
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def extract_bha_diameters(ws):
    """Лист 'Суточный отчет' — таблица 'КНБК и БТ' (геометрия компоновки по
    стволу) — для каждого элемента компоновки берём наружный диаметр
    (колонка 'Наруж. Ø', мм). Используется как запасной способ узнать
    типоразмер ВЗД/осциллятора для подбора ресурса (vzd_reference.py),
    когда сам типоразмер не удаётся распознать по тексту названия модели
    (см. пояснение пользователя — там же, где просил использовать эту
    таблицу как пример).

    Сопоставление с записью о том же элементе на листе 'Наработка' — по
    нормализованному названию (см. _normalize_bha_name); там к названию
    приписан суффикс "(...)", который _parse_component_name уже отрезает,
    так что сравнивать нужно с уже "очищенным" от суффикса именем (model).

    Возвращает {нормализованное_название: диаметр_мм}. При любой
    неожиданности в структуре листа просто возвращает {} — это
    вспомогательный, необязательный источник данных."""
    try:
        rows = _find_all_rows_startswith(ws, 2, "КНБК и БТ", max_row=SCAN_MAX_ROW)
        if not rows:
            return {}
        anchor = rows[0]
        header_row = anchor + 1
        name_col = _find_col_startswith(ws, header_row, "Наименование") or 2
        diam_col = _find_col_startswith(ws, header_row, "Наруж")
        if not diam_col:
            return {}
        result = {}
        r = header_row + 2  # header_row+1 — строка единиц измерения ("мм" и т.п.)
        limit = min(anchor + 200, ws.nrows)
        while r <= limit:
            name_val = _cell(ws, name_col, r)
            diam_val = _cell(ws, diam_col, r)
            name_str = _clean_str(name_val)
            if not name_str and diam_val is None:
                break  # пустая строка — конец таблицы
            if name_str and _is_number(diam_val):
                result[_normalize_bha_name(name_str)] = diam_val
            r += 1
        return result
    except Exception:
        return {}


def _looks_like_vzd(name):
    name = _clean_str(name)
    if not name:
        return False
    return any(marker in name for marker in VZD_NAME_MARKERS)


def _looks_like_osc(name):
    """Проверяет, похоже ли название элемента КНБК на осциллятор (см.
    OSC_SUBSTR_MARKERS/OSC_WORD_MARKERS)."""
    name = _clean_str(name)
    if not name:
        return False
    upper = name.upper()
    if any(marker in upper for marker in OSC_SUBSTR_MARKERS):
        return True
    tokens = re.findall(r"[0-9A-Za-zА-Яа-яЁё]+", upper)
    return any(tok in OSC_WORD_MARKERS for tok in tokens)


def _looks_like_sbt(name):
    """СБТ (стальные бурильные трубы) — определяется просто по началу
    названия элемента компоновки с 'СБТ' (в отличие от 'ТБТ'/'ТБПК' —
    утяжелённых/толстостенных труб, которые сюда не относятся, см.
    пояснение пользователя: 'только СБТ')."""
    name = _clean_str(name)
    return name.upper().startswith("СБТ")


def _parse_duration_hours(v):
    """Разбирает значение столбца 'Продолжительность, час' листа
    'Хронология работ' в часы (float). В реальных файлах это ТЕКСТ вида
    'ЧЧ:ММ' (в т.ч. '24:00' и больше — не ограничено сутками), но на
    всякий случай поддержаны и число (уже готовые часы), и datetime.time/
    timedelta (если Excel в конкретном файле отформатирует ячейку иначе)."""
    if _is_number(v):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        m = re.match(r"^(\d{1,4}):(\d{2})$", s)
        if m:
            return int(m.group(1)) + int(m.group(2)) / 60.0
        try:
            return float(s.replace(",", "."))
        except ValueError:
            return None
    if isinstance(v, datetime.time):
        return v.hour + v.minute / 60.0 + v.second / 3600.0
    if isinstance(v, datetime.timedelta):
        return v.total_seconds() / 3600.0
    return None


# Типы операций листа "Хронология работ" (столбец "Операция"), которые
# считаются "наработкой с вращением" СБТ во вкладке "Инструмент" — по
# указанию пользователя. Итоговый перечень (только для секции "Хвостовик"):
# механическое бурение, проработка, перезапись (в т.ч. ГК), промежуточная
# промывка, ориентирование КНБК, замер инклинометрии. В реальных файлах
# бурение записано ровно как "Механическое бурение" (точное совпадение), а
# остальные категории почти всегда записаны развёрнуто ("Проработка
# интервала слайдирования", "Подъем КНБК с перезаписью", "Промежуточная
# промывка в обсаженном стволе" и т.п.), а не одним словом — поэтому для
# них используется поиск по вхождению корня/фразы.
_SBT_OP_SUBSTR_MARKERS = (
    "проработ", "перезапис",
    "промежуточная промывка", "ориентирование кнбк", "замер инклинометрии",
)


def _op_matches_sbt_target(op_text):
    op = _clean_str(op_text).lower()
    if not op:
        return False
    if op == "механическое бурение":
        return True
    return any(marker in op for marker in _SBT_OP_SUBSTR_MARKERS)


def _trip_numbers_match(a, b):
    if a is None or b is None:
        return False
    try:
        return float(a) == float(b)
    except (TypeError, ValueError):
        return _clean_str(a) == _clean_str(b)


def extract_chronology_rows(wb):
    """Лист 'Хронология работ' — построчный список записей истории (лист
    накопительный, за все сутки с начала бурения, не только за текущие
    сутки рапорта). Нужен для пересчёта наработки циркуляции СБТ на
    вкладке 'Инструмент' по типам операций (см. compute_sbt_trip_hours).
    Каждая запись: {'trip', 'section', 'operation', 'duration_h'}.
    Возвращает [], если листа нет, нет нужных столбцов или структура
    неожиданная — не прерывая разбор файла."""
    if CHRONOLOGY_SHEET not in wb.sheetnames:
        return []
    try:
        ws = SheetGrid(wb[CHRONOLOGY_SHEET], max_row=CHRONOLOGY_MAX_ROW)
        header_row = _find_row_by_label(ws, 3, "Начало", max_row=20)
        if not header_row:
            return []
        section_col = _find_col_in_row(ws, header_row, "Секция", max_col=30)
        trip_col = _find_col_startswith(ws, header_row, "№", max_col=30)
        op_col = _find_col_in_row(ws, header_row, "Операция", max_col=30)
        dur_col = _find_col_startswith(ws, header_row, "Продолжительность", max_col=30)
        if not op_col or not dur_col:
            return []
        rows = []
        for r in range(header_row + 1, ws.nrows + 1):
            date_v = ws.cell(r, 2)
            time_v = ws.cell(r, 3)
            if date_v is None and time_v is None:
                continue
            rows.append({
                "trip": ws.cell(r, trip_col) if trip_col else None,
                "section": ws.cell(r, section_col) if section_col else None,
                "operation": ws.cell(r, op_col),
                "duration_h": _parse_duration_hours(ws.cell(r, dur_col)),
            })
        return rows
    except Exception:
        return []


def compute_sbt_trip_hours(chronology_rows, trip_number):
    """Наработка на инструмент с вращением для одного рейса секции
    'Хвостовик' — по указанию пользователя, взамен прежнего чтения
    'Наработка цирк. общ. за рейс' напрямую со строки СБТ листа 'Наработка'.
    Считается как сумма часов ('Продолжительность, час') по строкам листа
    'Хронология работ', где: секция — 'Хвостовик' (любой вариант, включая
    'Хвостовик II' — по вхождению), '№ рейса' совпадает с trip_number
    (границы рейса берём из блока листа 'Наработка' — см.
    extract_narabotka), а тип операции — механическое бурение, проработка
    (любой вариант), перезапись (в т.ч. ГК), промежуточная промывка,
    ориентирование КНБК или замер инклинометрии — см.
    _op_matches_sbt_target. Показатель считается только за рейс (не
    накопительно по скважине/инструменту между скважинами).
    Возвращает сумму часов (float) либо None, если trip_number не задан
    или подходящих строк не нашлось вовсе (тогда в таблице покажется
    прочерк, а не обманчивый 0)."""
    if not chronology_rows or trip_number is None:
        return None
    total = 0.0
    found = False
    for row in chronology_rows:
        if not _trip_numbers_match(row.get("trip"), trip_number):
            continue
        if "хвостовик" not in _clean_str(row.get("section")).lower():
            continue
        if not _op_matches_sbt_target(row.get("operation")):
            continue
        dur = row.get("duration_h")
        if dur is None:
            continue
        total += dur
        found = True
    return total if found else None


def extract_narabotka(wb, bha_diameters=None):
    """Лист 'Наработка' — таблица(ы) 'КНБК и БТ' (компоновка низа бурильной
    колонны текущего/последних рейсов). В файле может быть несколько таких
    таблиц подряд, если в течение суток КНБК поднимали и меняли — тогда
    возвращает запись по КАЖДОЙ из них (в порядке появления в файле).

    Для каждого блока берём первую строку (это всегда "Долото...") и, по
    указанию пользователя, СЛЕДУЮЩУЮ ЗА НЕЙ строку как кандидата в ВЗД —
    подтверждаем совпадением по характерным меткам в названии (см.
    VZD_NAME_MARKERS), чтобы не ошибиться, если порядок вдруг не совпадёт.

    Осциллятор (см. OSC_SUBSTR_MARKERS/OSC_WORD_MARKERS), в отличие от ВЗД,
    жёсткой позиции в компоновке не имеет — поэтому ищем его по всему блоку
    (от долота до начала следующего блока/конца таблицы), берём первое
    совпадение.

    bha_diameters — {нормализованное_название: диаметр_мм} с листа
    'Суточный отчет' (см. extract_bha_diameters) — запасной способ узнать
    типоразмер ВЗД/осциллятора для подбора ресурса, когда в самом названии
    модели типоразмер не распознать (см. vzd_reference.py).

    Возвращает список словарей:
        {"trip_number", "section", "zaboi_from", "zaboi_to",
         "bit": {"name","serial","circ_daily","circ_trip","time_drill","footage_trip"},
         "vzd": {...такие же поля + "model","maker","resource"} или None,
         "osc": {...такие же поля + "model","maker","resource"} или None,
         "sbt": {...такие же поля + "model","maker","resource"} или None
                (заполняется только для секции "Хвостовик" — вкладка
                "Инструмент", см. _looks_like_sbt)}
    """
    bha_diameters = bha_diameters or {}
    if NARABOTKA_SHEET not in wb.sheetnames:
        return []
    ws = SheetGrid(wb[NARABOTKA_SHEET], max_row=NARABOTKA_SEARCH_RANGE)

    # Для пересчёта наработки циркуляции СБТ (см. compute_sbt_trip_hours) —
    # читаем один раз на файл, используем только если в каком-то блоке
    # найдётся СБТ секции "Хвостовик".
    chronology_rows = None

    anchors = _find_all_rows_startswith(ws, 2, "КНБК и БТ", max_row=NARABOTKA_SEARCH_RANGE)
    blocks = []
    for anchor_idx, anchor in enumerate(anchors):
        header_row = anchor + 1
        name_col = _find_col_startswith(ws, header_row, "Наименование") or 2
        serial_col = _find_col_in_row(ws, header_row, "Серийный номер")
        circ_daily_col = _find_col_startswith(ws, header_row, "Наработка цирк сут")
        circ_trip_col = (_find_col_startswith(ws, header_row, "Наработка цирк. общ")
                          or _find_col_startswith(ws, header_row, "Наработка цирк общ"))
        time_col = _find_col_startswith(ws, header_row, "Время бурения")
        footage_col = _find_col_startswith(ws, header_row, "Проходка за рейс")
        if not serial_col:
            continue

        data_start = header_row + 2  # header_row+1 — строка единиц измерения ("ч.")
        bit_row = None
        for r in range(data_start, data_start + 3):
            name_val = _clean_str(_cell(ws, name_col, r))
            if name_val.startswith("Долото"):
                bit_row = r
                break
        if bit_row is None:
            continue

        def read_component(r):
            return {
                "name": _cell(ws, name_col, r),
                "serial": _cell(ws, serial_col, r) if serial_col else None,
                "circ_daily": _cell(ws, circ_daily_col, r) if circ_daily_col else None,
                "circ_trip": _cell(ws, circ_trip_col, r) if circ_trip_col else None,
                "time_drill": _cell(ws, time_col, r) if time_col else None,
                "footage_trip": _cell(ws, footage_col, r) if footage_col else None,
            }

        bit = read_component(bit_row)

        # ВЗД обычно идёт сразу после долота, но между ними бывает переводник
        # (крестовина/сab) — поэтому просматриваем несколько строк вперёд, а
        # не только следующую, и подтверждаем находку меткой в названии
        # (VZD_NAME_MARKERS), чтобы не спутать с рулевой/роторной системой
        # (РУС) или другим элементом КНБК, где мотора нет вовсе.
        vzd = None
        for r in range(bit_row + 1, bit_row + 9):
            name_val = _clean_str(_cell(ws, name_col, r))
            if name_val.startswith("Долото") or name_val.startswith("КНБК и БТ"):
                break
            if _looks_like_vzd(name_val):
                vzd = read_component(r)
                model, maker = _parse_component_name(name_val)
                # Текст в скобках после названия в некоторых рапортах — не
                # производитель ВЗД, а сервисная/подрядная компания по
                # другому элементу компоновки (ошибочно/по инерции
                # скопированная ячейка) — поэтому для "ДРУ..." он по
                # умолчанию игнорируется. Приоритет: точное совпадение со
                # справочником (там есть исключения вроде "ДРУ 3 120Н" —
                # это ВНИИБТ, а не Радиус-Сервис, несмотря на обозначение
                # "ДРУ"); если модель в справочнике не нашлась, но название
                # всё равно начинается с "ДРУ" — по умолчанию считаем
                # Радиус-Сервис (единственный производитель, использующий
                # это обозначение в общем случае).
                matched_maker = vzd_reference.find_vzd_maker(model)
                if matched_maker:
                    maker = matched_maker
                elif model.strip().lower().startswith("дру"):
                    maker = "Радиус-Сервис"
                vzd["model"] = model
                vzd["maker"] = maker
                resource = vzd_reference.find_vzd_resource(model)
                if resource is None:
                    diam = bha_diameters.get(_normalize_bha_name(model))
                    resource = vzd_reference.find_vzd_resource_by_diameter(diam)
                vzd["resource"] = resource
                break

        # Осциллятор может стоять где угодно в компоновке (не привязан к
        # позиции долота) — просматриваем весь блок, до начала следующего
        # блока КНБК (если он есть) либо до конца области поиска.
        block_end = (anchors[anchor_idx + 1] - 1) if anchor_idx + 1 < len(anchors) else NARABOTKA_SEARCH_RANGE
        osc = None
        for r in range(bit_row + 1, block_end + 1):
            name_val = _clean_str(_cell(ws, name_col, r))
            if name_val.startswith("КНБК и БТ"):
                break
            if _looks_like_osc(name_val):
                osc = read_component(r)
                model, maker = _parse_component_name(name_val)
                osc["model"] = model
                osc["maker"] = maker
                resource = vzd_reference.find_osc_resource(model, maker)
                if resource is None:
                    diam = bha_diameters.get(_normalize_bha_name(model))
                    resource = vzd_reference.find_osc_resource_by_diameter(diam, maker)
                osc["resource"] = resource
                break

        # Забой (от/до) и номер/секция рейса — из мини-таблицы, идущей сразу
        # перед этим блоком (см. docstring extract_narabotka).
        trip_number = section = zaboi_from = zaboi_to = None
        if anchor - 2 >= 1:
            subheader_row = anchor - 2
            data_row = anchor - 1
            from_col = _find_col_in_row(ws, subheader_row, "От")
            to_col = _find_col_in_row(ws, subheader_row, "До")
            trip_number = _cell(ws, 2, data_row)
            section = _cell(ws, 3, data_row)
            zaboi_from = _cell(ws, from_col, data_row) if from_col else None
            zaboi_to = _cell(ws, to_col, data_row) if to_col else None

        # СБТ (вкладка "Инструмент") — только для секции "Хвостовик" (по
        # указанию пользователя). В блоке КНБК нередко встречается НЕСКОЛЬКО
        # строк "СБТ..." (разные типоразмеры/интервалы одной и той же
        # колонны труб на разных этапах рейса) — берём ПОСЛЕДНЮЮ из них
        # (обычно самая нижняя строка блока, см. пояснение пользователя:
        # "обычно нижняя строчка, бывает и чуть выше").
        sbt = None
        section_norm = _clean_str(section).lower()
        if "хвостовик" in section_norm:
            for r in range(bit_row + 1, block_end + 1):
                name_val = _clean_str(_cell(ws, name_col, r))
                if name_val.startswith("КНБК и БТ"):
                    break
                if _looks_like_sbt(name_val):
                    candidate = read_component(r)
                    model, maker = _parse_component_name(name_val)
                    candidate["model"] = model
                    candidate["maker"] = maker
                    candidate["resource"] = SBT_RESOURCE_HOURS
                    sbt = candidate
            if sbt is not None:
                # Наработка циркуляции СБТ — не берётся напрямую со строки
                # листа "Наработка" (как для ВЗД/осциллятора), а пересчитывается
                # из листа "Хронология работ" по типам операций, см.
                # compute_sbt_trip_hours (по указанию пользователя).
                if chronology_rows is None:
                    chronology_rows = extract_chronology_rows(wb)
                sbt["circ_trip"] = compute_sbt_trip_hours(chronology_rows, trip_number)

        blocks.append({
            "trip_number": trip_number,
            "section": section,
            "zaboi_from": zaboi_from,
            "zaboi_to": zaboi_to,
            "bit": bit,
            "vzd": vzd,
            "sbt": sbt,
            "osc": osc,
        })
    return blocks


def extract_plan_msp(wb, section_name, current_depth=None):
    """План МСП (мех. скорости проходки) для операции 'Механическое бурение'
    текущей секции — с листа 'Программа работ'. Возвращает словарь
    {'plan_from','plan_to','plan_interval','plan_time_h','plan_msp'} либо
    None, если данных нет.

    В плане на одну секцию бывает НЕСКОЛЬКО строк 'Механическое бурение'
    (например, пилотный ствол и потом основной); если current_depth задан,
    выбираем строку, чья плановая глубина ('до') ближе всего к нему —
    иначе берём последнюю найденную."""
    if PROGRAM_SHEET not in wb.sheetnames or not section_name:
        return None
    ws = SheetGrid(wb[PROGRAM_SHEET], max_row=1200)

    header_row = None
    section_col = op_col = None
    for r in range(1, 15):
        sc = _find_col_in_row(ws, r, "Секция")
        oc = _find_col_in_row(ws, r, "Операция")
        if sc and oc:
            header_row = r
            section_col, op_col = sc, oc
            break
    if not header_row:
        return None

    plan_group_row = None
    depth_col = dur_col = None
    for r in range(header_row, header_row + 3):
        pc = _find_col_in_row(ws, r, "План")
        if pc:
            plan_group_row = r
            break
    if plan_group_row:
        sub_row = plan_group_row + 1
        depth_col = _find_col_startswith(ws, sub_row, "Глубина")
        dur_col = _find_col_startswith(ws, sub_row, "Продолжительность")
    if not depth_col or not dur_col:
        return None

    section_name_norm = _clean_str(section_name)
    candidates = []
    r = header_row + 2
    limit = min(ws.nrows, header_row + 1000)
    while r <= limit:
        sec = _clean_str(_cell(ws, section_col, r))
        op = _clean_str(_cell(ws, op_col, r))
        if sec == section_name_norm and op == "Механическое бурение":
            to_val = _cell(ws, depth_col, r)
            if _is_number(to_val):
                candidates.append((r, to_val))
        r += 1

    if not candidates:
        return None

    if current_depth is not None and _is_number(current_depth):
        target_row, to_val = min(candidates, key=lambda c: abs(c[1] - current_depth))
    else:
        target_row, to_val = candidates[-1]

    from_val = None
    rr = target_row - 1
    while rr >= header_row + 2:
        v = _cell(ws, depth_col, rr)
        if _is_number(v) and v != to_val:
            from_val = v
            break
        rr -= 1

    plan_time = _cell(ws, dur_col, target_row)
    plan_interval = (to_val - from_val) if (from_val is not None and _is_number(to_val)) else None
    plan_msp = None
    if plan_interval is not None and _is_number(plan_time) and plan_time:
        plan_msp = plan_interval / plan_time

    return {
        "plan_from": from_val,
        "plan_to": to_val,
        "plan_interval": plan_interval,
        "plan_time_h": plan_time,
        "plan_msp": plan_msp,
    }


def parse_sr_file(path, build_chart=True):
    """Разбирает один файл СР и возвращает словарь с данными для отчёта.
    Бросает ParseError с понятным сообщением при неустранимой проблеме.

    build_chart — если True (по умолчанию), строит график "глубина-день"
    средствами самой программы (модуль depth_chart, без внешних
    зависимостей — не требует LibreOffice) и кладёт готовую SVG-разметку в
    data["depth_chart_svg"]. Если False — график не строится вовсе."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        raise ParseError(f"не удалось открыть книгу Excel: {e}")

    if MAIN_SHEET not in wb.sheetnames:
        raise ParseError(f"в книге нет листа «{MAIN_SHEET}»")

    ws = SheetGrid(wb[MAIN_SHEET], max_row=SCAN_MAX_ROW + OPS_SEARCH_RANGE)

    try:
        header = extract_header(ws)
        construction = extract_construction(ws)
        depth_block = extract_depth_block(ws)
        bit = extract_bit(ws)
        ops = extract_ops(ws)
        mud_params = extract_mud_params(ws)
        npv = extract_npv(wb)
        commercial_speed = extract_commercial_speed(wb)
        stage_deviation = extract_stage_deviation(wb)
        bha_diameters = extract_bha_diameters(ws)
        narabotka = extract_narabotka(wb, bha_diameters)
        current_stage = extract_current_stage(wb)
    except Exception as e:
        raise ParseError(f"ошибка разбора содержимого: {e}")

    # Текущий (последний) блок КНБК с листа "Наработка" — используется, чтобы
    # дополнить блок "Долото" рейсовыми показателями (проходка/МСП за рейс)
    # и определить активную секцию для плановой МСП с листа "Программа работ".
    current_block = narabotka[-1] if narabotka else None
    current_section = current_block.get("section") if current_block else None

    if bit is not None:
        trip_bit = (current_block or {}).get("bit") or {}
        trip_footage = trip_bit.get("footage_trip")
        trip_time = trip_bit.get("time_drill")
        bit["serial"] = trip_bit.get("serial")
        bit["trip_footage"] = trip_footage
        bit["trip_time_h"] = trip_time
        bit["trip_rop"] = (
            trip_footage / trip_time
            if _is_number(trip_footage) and _is_number(trip_time) and trip_time
            else None
        )
        bit["plan_msp"] = extract_plan_msp(wb, current_section, bit.get("interval_to"))

    depth_chart_svg = None
    chart_error = None
    if build_chart:
        depth_chart_svg, chart_error = depth_chart.build_depth_chart_svg(wb)

    data = {
        "source_file": os.path.basename(path),
        "header": header,
        "construction": construction,
        "depth": depth_block,
        "bit": bit,
        "ops": ops,
        "mud_params": mud_params,
        "npv": npv,
        "commercial_speed": commercial_speed,
        "stage_deviation": stage_deviation,
        "narabotka": narabotka,
        "current_stage": current_stage,
        "depth_chart_svg": depth_chart_svg,
        "chart_error": chart_error,
    }
    return data


def parse_folder(folder, on_progress=None, build_chart=True):
    """Разбирает все СР-файлы в папке.
    on_progress(index, total, filename) — необязательный колбэк прогресса.
    Возвращает (список_успешных_данных, список_ошибок[(filename, message)])."""
    files = find_xlsx_files(folder)
    return _parse_file_list(files, on_progress, build_chart)


def parse_source(input_path, on_progress=None, build_chart=True):
    """Универсальная точка входа: принимает папку, отдельный файл .xlsx,
    либо архив .zip/.rar с файлами СР. Архивы распаковываются во временную
    папку, которая удаляется по завершении (в т.ч. при ошибке).
    on_progress(index, total, filename) — необязательный колбэк прогресса.
    build_chart — построить график "глубина-день" средствами самой
    программы для каждой скважины (см. parse_sr_file).
    Возвращает (список_успешных_данных, список_ошибок[(filename, message)]).
    Бросает ParseError, если путь не найден, тип файла не поддерживается,
    либо архив не удалось распаковать."""
    files, cleanup_dirs = resolve_xlsx_sources(input_path)
    try:
        if not files:
            raise ParseError("Не найдено ни одного файла .xlsx по указанному пути.")
        return _parse_file_list(files, on_progress, build_chart)
    finally:
        for d in cleanup_dirs:
            shutil.rmtree(d, ignore_errors=True)


def _parse_file_list(files, on_progress=None, build_chart=True):
    results = []
    errors = []
    total = len(files)
    for i, path in enumerate(files, start=1):
        if on_progress:
            on_progress(i, total, os.path.basename(path))
        try:
            data = parse_sr_file(path, build_chart=build_chart)
            results.append(data)
        except ParseError as e:
            errors.append((os.path.basename(path), str(e)))
        except Exception as e:
            errors.append((os.path.basename(path), f"непредвиденная ошибка: {e}"))
    return results, errors
