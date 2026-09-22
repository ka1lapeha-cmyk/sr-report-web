\
# -*- coding: utf-8 -*-
"""
report_html.py
Сборка сводного HTML-отчёта по данным, извлечённым sr_parser.py из файлов СР.

Фирменный стиль — тёмно-зелёный/графитовый (по логотипу ООО "ИНК") с
белыми карточками; логотип встраивается как base64, внешних ресурсов нет.
"""

import base64
import datetime
import html as html_lib
import json
import os
import re
import sys

import vzd_reference


def _base_dir():
    """Папка с ресурсами программы: при обычном запуске — рядом с этим .py
    файлом; при запуске из .exe, собранного PyInstaller-ом — временная папка,
    куда PyInstaller распаковывает приложенные данные (assets/ и т.п.),
    см. build.bat."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
    return os.path.dirname(os.path.abspath(__file__))


ASSETS_DIR = os.path.join(_base_dir(), "assets")
LOGO_PATH = os.path.join(ASSETS_DIR, "inc_logo.png")
# Запасной путь — файл лежит прямо рядом со скриптом, без подпапки assets/.
# Нужно для веб-версии (Streamlit Cloud): загрузить файл в подпапку через
# веб-интерфейс GitHub без git неочевидно, поэтому логотип ищем и там,
# и там — что найдётся.
LOGO_PATH_FLAT = os.path.join(_base_dir(), "inc_logo.png")

REPORT_VERSION = "1.13"

NUM_PREFIX_RE = re.compile(r"\(\d+(?:\.\d+)?\)\s*")


# ---------------------------------------------------------------------------
# Форматирование значений
# ---------------------------------------------------------------------------

def esc(v):
    if v is None:
        return ""
    return html_lib.escape(str(v))


def fmt_num(v, digits=2):
    """Число в русском формате (запятая как разделитель дробной части)."""
    if v is None or v == "":
        return "—"
    if isinstance(v, bool):
        return "—"
    if isinstance(v, (int, float)):
        f = float(v)
        if f.is_integer():
            return str(int(f))
        s = f"{f:.{digits}f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")
    return esc(v)


def fmt_signed(v, digits=1):
    """Число со знаком (+/−) и запятой, для отклонений: '+95,6', '−10,3'."""
    if v is None or v == "":
        return "—"
    try:
        f = float(v)
    except (TypeError, ValueError):
        return esc(v)
    sign = "+" if f >= 0 else "−"
    s = f"{abs(f):.{digits}f}"
    if "." in s:
        s = s.rstrip("0").rstrip(".")
        if s == "":
            s = "0"
    return f"{sign}{s.replace('.', ',')}"


def fmt_date(v):
    if v is None or v == "":
        return "—"
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d.%m.%Y")
    return esc(v)


def fmt_datetime_like(v):
    """Дата/время: понимает и datetime-объекты, и уже готовые строки из Excel
    (в файлах СР даты типа 'Начало работ' часто хранятся как текст)."""
    if v is None or v == "":
        return "—"
    if isinstance(v, datetime.datetime):
        return v.strftime("%d.%m.%Y %H:%M")
    if isinstance(v, datetime.date):
        return v.strftime("%d.%m.%Y")
    return esc(v)


def parse_ru_datetime(v):
    """Пытается получить datetime из значения ячейки (объект или строка
    'ДД.ММ.ГГГГ ЧЧ:ММ' / 'ДД.ММ.ГГГГ')."""
    if isinstance(v, datetime.datetime):
        return v
    if isinstance(v, datetime.date):
        return datetime.datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y"):
            try:
                return datetime.datetime.strptime(s, fmt)
            except ValueError:
                continue
    return None


def split_ru_entries(raw):
    """Разбивает 'сырые' ячейки вида '(1.1) ООО "X"\\n(Заказчик)\\n\\n(2.1) ...'
    на список аккуратных текстовых фрагментов без номеров-пометок и переносов
    строк (такие пометки встречаются в столбцах Категория/Продолж./Ответственный
    блока 'Суточные операции', где НПВ может относиться сразу к нескольким
    пунктам за один интервал времени)."""
    if raw is None or raw == "":
        return []
    s = str(raw).replace("\r", "")
    parts = NUM_PREFIX_RE.split(s)
    out = []
    for p in parts:
        p = re.sub(r"\n+", " ", p)
        p = re.sub(r"\s{2,}", " ", p).strip(" ;")
        if p:
            out.append(p)
    return out


def fmt_multi(raw, sep="; "):
    entries = split_ru_entries(raw)
    return esc(sep.join(entries)) if entries else "—"


def safe_key(v):
    """Ключ для сортировки, устойчивый к None и смешанным типам."""
    if v is None:
        return (1, "")
    return (0, str(v))


# Порядок месторождений в отчёте — не алфавитный, а заданный явно (по
# указанию пользователя): сначала перечисленные ниже месторождения именно
# в этом порядке, затем остальные (не из списка) — по алфавиту, и в самом
# конце — Большетирские/Верхнетирские скважины.
FIELD_ORDER_FIRST = ("маччобинское", "мирнинское", "кийское", "аянское", "ярактинское")
FIELD_ORDER_LAST = ("большетирское", "верхнетирское")


def field_group_key(field):
    """Ключ сортировки месторождения — см. FIELD_ORDER_FIRST/FIELD_ORDER_LAST."""
    name = _clean_str_local(field).lower()
    if not name:
        return (1, "")
    for i, kw in enumerate(FIELD_ORDER_FIRST):
        if name.startswith(kw):
            return (0, i)
    for i, kw in enumerate(FIELD_ORDER_LAST):
        if name.startswith(kw):
            return (2, i)
    return (1, name)


def _clean_str_local(v):
    return "" if v is None else str(v).strip()


# ---------------------------------------------------------------------------
# Кластеры месторождений
# ---------------------------------------------------------------------------
# Базовый справочник получен из выгрузки "Скважины/Кластера" (список скважин
# с месторождением, участком недр и кластером). Матчинг — по значимому
# фрагменту названия месторождения (устойчиво к сокращениям вида "НГКМ" vs
# "нефтегазоконденсатное месторождение" и к написанию "ё"/"е"), поэтому
# работает и для более коротких названий полей, которые встречаются в самих
# файлах СР (там название часто в сокращённом виде).
#
# Если в СР попадётся месторождение, которого нет ни в этих правилах, ни в
# сохранённых пользователем уточнениях (см. main.py — там при обнаружении
# такого месторождения программа спрашивает пользователя и запоминает
# ответ) — оно попадает в отдельный блок "Прочие месторождения" в самом
# низу сводной таблицы, см. group_by_cluster().
CLUSTER_ORDER = ("Северный", "Центральный", "Западный")
OTHER_CLUSTER = "Прочие месторождения"

FIELD_CLUSTER_RULES = (
    ("маччобинск", "Северный"),
    ("мирнинск", "Северный"),
    ("марковск", "Центральный"),
    ("мышевск", "Центральный"),
    ("ярактинск", "Центральный"),
    ("кийск", "Центральный"),
    ("аянск", "Центральный"),
    ("большетирск", "Западный"),
    ("верхнетирск", "Западный"),
    ("ичедин", "Западный"),
)


def normalize_field_name(v):
    """Нормализация названия месторождения для сопоставления с кластером:
    в нижний регистр, без лишних пробелов, 'ё' -> 'е' (в разных файлах и
    справочниках встречается по-разному)."""
    return _clean_str_local(v).lower().replace("ё", "е")


def cluster_for_field(field_name, overrides=None):
    """Кластер месторождения: сперва проверяются сохранённые пользователем
    уточнения (overrides, словарь normalize_field_name(месторождение) ->
    кластер), затем встроенные правила FIELD_CLUSTER_RULES (сопоставление
    по вхождению значимого фрагмента). None, если кластер не определён —
    такие месторождения показываются в блоке 'Прочие месторождения'."""
    norm = normalize_field_name(field_name)
    if not norm:
        return None
    if overrides and norm in overrides:
        return overrides[norm]
    for root, cluster in FIELD_CLUSTER_RULES:
        if root in norm:
            return cluster
    return None


def unknown_fields(results, overrides=None):
    """Уникальные месторождения (в порядке первого появления) из списка
    разобранных рапортов, для которых не удалось определить кластер — ни
    по встроенным правилам, ни по сохранённым уточнениям. main.py использует
    это перед сборкой отчёта, чтобы спросить пользователя про такие
    месторождения (см. модуль main.py)."""
    seen_norm = set()
    out = []
    for d in results:
        field = (d.get("header") or {}).get("field")
        norm = normalize_field_name(field)
        if not norm or norm in seen_norm:
            continue
        seen_norm.add(norm)
        if cluster_for_field(field, overrides) is None:
            out.append(field)
    return out


def group_by_cluster(items, overrides=None):
    """items — список (idx, d), УЖЕ отсортированный (см. build_html.sort_key).
    Возвращает список (название_кластера, [(idx, d), ...]) в порядке
    CLUSTER_ORDER, с блоком OTHER_CLUSTER в конце для месторождений без
    определённого кластера (если такие есть). Относительный порядок скважин
    внутри каждого кластера сохраняется как во входном items."""
    buckets = {c: [] for c in CLUSTER_ORDER}
    other = []
    for idx, d in items:
        field = (d.get("header") or {}).get("field")
        cluster = cluster_for_field(field, overrides)
        if cluster in buckets:
            buckets[cluster].append((idx, d))
        else:
            other.append((idx, d))
    groups = [(c, buckets[c]) for c in CLUSTER_ORDER if buckets[c]]
    if other:
        groups.append((OTHER_CLUSTER, other))
    return groups


def numeric_or_text_key(v):
    if v is None:
        return (1, 0, "")
    s = str(v)
    m = re.match(r"^\s*(\d+)", s)
    if m:
        return (0, int(m.group(1)), s)
    return (0, 10**9, s)


def collect_responsible_parties(ops):
    """Уникальный список ответственных сторон за НПВ из суточных операций
    (для строки 'Виновная сторона' над свёрнутыми блоками НПВ)."""
    seen = []
    for e in ops or []:
        for entry in split_ru_entries(e.get("responsible")):
            if entry not in seen:
                seen.append(entry)
    return seen


# ---------------------------------------------------------------------------
# Раскрывающийся блок <details>
# ---------------------------------------------------------------------------

def render_details(summary_html, body_html, open_=False, extra_class=""):
    open_attr = " open" if open_ else ""
    cls = f"collapsible {extra_class}".strip()
    return (
        f'<details class="{cls}"{open_attr}>'
        f'<summary>{summary_html}</summary>'
        f'<div class="collapsible-body">{body_html}</div>'
        f'</details>'
    )


# ---------------------------------------------------------------------------
# HTML-фрагменты для одного рапорта (одной скважины)
# ---------------------------------------------------------------------------

def render_construction_table(rows):
    if not rows:
        return '<p class="muted">Данные по конструкции скважины не заполнены в рапорте.</p>'
    body = []
    for r in rows:
        body.append(
            f"<tr><td>{esc(r['name'])}</td><td>{fmt_num(r['diameter'])}</td>"
            f"<td>{fmt_num(r['plan'])}</td><td>{fmt_num(r['fact'])}</td></tr>"
        )
    return f"""
<table class="tbl">
  <thead><tr><th>Секция</th><th>&#216; ствола, мм</th><th>План, м</th><th>Факт, м</th></tr></thead>
  <tbody>{''.join(body)}</tbody>
</table>"""


def render_bit_footage_cell(bit):
    """Проходка долотом: 'за сутки / за рейс' (рейсовая — с листа 'Наработка',
    None, если для текущего рейса нет данных о подъёме/наработке)."""
    daily = fmt_num(bit.get("daily_footage"))
    trip = bit.get("trip_footage")
    if trip is None:
        return daily
    return f"{daily} / {fmt_num(trip)}"


def render_plan_msp_trigger(plan, bit):
    """Кликабельное значение план. МСП операции 'Механическое бурение' по
    текущей секции (лист 'Программа работ') — при клике открывает всплывающее
    окно с деталями: план МСП, план 'от'/'до', план интервал и факт. 'от'
    (начало интервала бурения текущим долотом)."""
    if not plan or plan.get("plan_msp") is None:
        return "—"
    data = {
        "msp": fmt_num(plan.get("plan_msp")),
        "from": fmt_num(plan.get("plan_from")),
        "to": fmt_num(plan.get("plan_to")),
        "interval": fmt_num(plan.get("plan_interval")),
        "fact_from": fmt_num((bit or {}).get("interval_from")),
    }
    data_attr = esc(json.dumps(data, ensure_ascii=False))
    return (
        f'<span class="plan-msp-trigger sr-trigger" tabindex="0" role="button" '
        f'data-plan="{data_attr}" onclick="srShowPlanPopover(event, this)">'
        f'{fmt_num(plan.get("plan_msp"))}</span>'
    )


def render_bit_speed_cell(bit):
    """Мех. скорость долотом: 'за сутки / за рейс / план' — план кликабелен
    (см. render_plan_msp_trigger)."""
    daily_rop = fmt_num(bit.get("rop"))
    trip_rop = fmt_num(bit.get("trip_rop")) if bit.get("trip_rop") is not None else "—"
    plan_html = render_plan_msp_trigger(bit.get("plan_msp"), bit)
    return f"{daily_rop} / {trip_rop} / {plan_html}"


def render_bit_block(bit):
    if not bit:
        return '<p class="muted">Данные по долоту за сутки не зафиксированы (спуск/подъём КНБК не проводился).</p>'
    items = [
        ("Изготовитель", esc(bit.get("maker")) or "—"),
        ("Диаметр, мм", fmt_num(bit.get("diameter"))),
        ("Тип", esc(bit.get("type")) or "—"),
        ("Модель", esc(bit.get("model")) or "—"),
        ("Серийный номер", esc(bit.get("serial")) or "—"),
        ("Интервал бурения, м", f"{fmt_num(bit.get('interval_from'))} – {fmt_num(bit.get('interval_to'))}"),
        ("Проходка, м (сутки / рейс)", render_bit_footage_cell(bit)),
        ("МСП, м/ч (сутки / рейс / план)", render_bit_speed_cell(bit)),
    ]
    cells = "".join(f'<div class="kv"><span class="k">{k}</span><span class="v">{v}</span></div>' for k, v in items)
    return f'<div class="kv-grid">{cells}</div>'


def render_comment_trigger(comment):
    """Кликабельная кнопка 'Комментарий' — при клике открывает всплывающее
    окно (тот же механизм, что и у render_plan_msp_trigger) с полным
    текстом комментария к строке суточных операций (лист 'Суточный отчет',
    столбец 'Комментарий')."""
    text = _clean_str_local(comment)
    if not text:
        return '<span class="muted">&mdash;</span>'
    data_attr = esc(json.dumps({"comment": text}, ensure_ascii=False))
    return (
        f'<button type="button" class="comment-trigger sr-trigger" '
        f'data-comment="{data_attr}" onclick="srShowCommentPopover(event, this)">Комментарий</button>'
    )


def render_ops_events(events):
    if not events:
        return '<p class="muted">Событий НПВ в суточных операциях не зафиксировано.</p>'
    body = []
    for e in events:
        body.append(
            "<tr>"
            f"<td>{esc(e.get('time_start'))}&ndash;{esc(e.get('time_end'))}</td>"
            f"<td>{fmt_multi(e.get('category'))}</td>"
            f"<td>{fmt_multi(e.get('duration'))}</td>"
            f"<td>{fmt_multi(e.get('responsible'))}</td>"
            f"<td>{render_comment_trigger(e.get('comment'))}</td>"
            "</tr>"
        )
    return f"""
<table class="tbl">
  <thead><tr><th>Время</th><th>Категория НПВ</th><th>Продолж.</th><th>Ответственный</th><th>Комментарий</th></tr></thead>
  <tbody>{''.join(body)}</tbody>
</table>"""


def render_npv_table(npv):
    by_org = npv.get("by_org") or []
    total = npv.get("total_hours")
    if not by_org:
        return f'<p class="muted">Накопленное НПВ по организациям не детализировано. Итого по скважине: {fmt_num(total)} ч.</p>'
    body = []
    for row in by_org:
        service = f" ({esc(row['service'])})" if row.get("service") else ""
        body.append(f"<tr><td>{esc(row['org'])}{service}</td><td>{fmt_num(row['hours'])}</td></tr>")
    return f"""
<table class="tbl">
  <thead><tr><th>Ответственная сторона</th><th>Всего, час (накопительно)</th></tr></thead>
  <tbody>{''.join(body)}</tbody>
  <tfoot><tr><td>Итого</td><td>{fmt_num(total)}</td></tr></tfoot>
</table>"""


def _fmt_mud_time(v):
    if v is None or v == "":
        return "—"
    if isinstance(v, (datetime.datetime, datetime.time)):
        return v.strftime("%H:%M")
    return esc(v)


def _render_mud_value_cell(value, unit, status):
    """Ячейка факт. значения параметра промывочной жидкости: подсвечивается
    красным, если факт вышел за плановую границу (status == 'bad'), зелёным —
    если в пределах плана (status == 'ok'); без подсветки, если сравнивать
    не с чем (плановая граница не задана/не разобрана или факта нет)."""
    if value is None or value == "":
        return "—"
    text = f"{fmt_num(value)} {unit}".strip() if unit else fmt_num(value)
    if status == "bad":
        return f'<span class="circ-over">{text}</span>'
    if status == "ok":
        return f'<span class="circ-ok">{text}</span>'
    return text


def render_mud_params_block(mud_params):
    """Блок 'Параметры промывочной жидкости' (план/факт) — лист 'Суточный
    отчет', см. sr_parser.extract_mud_params. Факт выходит за плановую
    границу — ячейка красная, в пределах плана — зелёная; для параметров без
    плана (или там, где план не задан числом/диапазоном) подсветки нет."""
    if not mud_params or not mud_params.get("params"):
        return '<p class="muted">Блок «Параметры промывочной жидкости» не найден или не заполнен в рапорте.</p>'

    params = mud_params["params"]
    facts = mud_params.get("facts") or []

    info_bits = []
    if mud_params.get("mud_type"):
        info_bits.append(f"Тип раствора: <b>{esc(mud_params['mud_type'])}</b>")
    if mud_params.get("system"):
        info_bits.append(f"Система: <b>{esc(mud_params['system'])}</b>")
    info_html = f'<p class="mud-info">{" &middot; ".join(info_bits)}</p>' if info_bits else ""

    if not facts:
        return info_html + '<p class="muted">Замеров параметров бурового раствора за сутки не зафиксировано.</p>'

    head_cols = "".join(f"<th>Факт, {_fmt_mud_time(f.get('time'))}</th>" for f in facts)
    thead = f"<tr><th>Параметр</th><th>План</th>{head_cols}</tr>"

    body = []
    for p in params:
        plan_text = esc(p["plan_text"]) if p["plan_text"] else "—"
        cells = "".join(
            f"<td>{_render_mud_value_cell(f['values'].get(p['key']), p['unit'], f['statuses'].get(p['key']))}</td>"
            for f in facts
        )
        unit_suffix = f", {esc(p['unit'])}" if p["unit"] else ""
        label = f"{esc(p['label'])}{unit_suffix}"
        body.append(f"<tr><td>{label}</td><td>{plan_text}</td>{cells}</tr>")

    return f"""
{info_html}
<table class="tbl mud-tbl">
  <thead>{thead}</thead>
  <tbody>{''.join(body)}</tbody>
</table>
<p class="mud-legend"><span class="circ-over">красным</span> — факт вне плановой границы, <span class="circ-ok">зелёным</span> — в пределах плана.</p>"""


def render_contacts(h):
    def person(role_label, name, phone, ip, email):
        sub_bits = []
        if phone:
            sub_bits.append(esc(phone))
        if ip:
            sub_bits.append(f"IP: {esc(ip)}")
        if email:
            sub_bits.append(esc(email))
        sub_html = f'<span class="contact-sub">{" &middot; ".join(sub_bits)}</span>' if sub_bits else ""
        return (
            f'<div class="kv"><span class="k">{role_label}</span>'
            f'<span class="v">{esc(name) or "—"}</span>{sub_html}</div>'
        )

    blocks = [
        f'<div class="kv"><span class="k">Буровая установка</span><span class="v">{esc(h.get("rig")) or "—"}</span></div>',
        f'<div class="kv"><span class="k">Буровая бригада</span><span class="v">{esc(h.get("brigade")) or "—"}</span></div>',
        person("Буровой мастер", h.get("driller"), h.get("driller_phone"), h.get("driller_ip"), h.get("driller_email")),
        person("Супервайзер", h.get("supervisor"), h.get("supervisor_phone"), h.get("supervisor_ip"), h.get("supervisor_email")),
        person("Куратор объекта", h.get("curator"), h.get("curator_phone"), h.get("curator_ip"), h.get("curator_email")),
    ]
    return f'<div class="kv-grid contacts-grid">{"".join(blocks)}</div>'


def render_schedule_block(header, commercial_speed, stage_deviation, depth_chart_svg=None):
    """Блок 'Отклонение от графика (глубина-день)' — короткая выжимка на
    основе шапки рапорта (норматив/начало/прогноз), коммерческой скорости
    (лист 'Программа работ') и разбивки отклонения по этапам (лист 'ТЭП')."""
    lag = header.get("lag_days")
    norm_days = header.get("norm_days")
    start_dt = parse_ru_datetime(header.get("work_start"))
    forecast_end = header.get("forecast_end")

    lines = []
    status = None  # 'отставание' / 'опережение' / None (по графику или нет данных)

    if isinstance(lag, (int, float)) and start_dt and isinstance(norm_days, (int, float)):
        plan_end_dt = start_dt + datetime.timedelta(days=float(norm_days))
        plan_end_str = f"~{plan_end_dt.strftime('%d.%m.%Y')}"
        start_str = start_dt.strftime("%d.%m.%Y")
        forecast_str = fmt_datetime_like(forecast_end)

        if lag > 0.005:
            status = "отставание"
        elif lag < -0.005:
            status = "опережение"

        if status:
            lines.append(
                f"<b>{fmt_signed(lag, 2)} сут</b> — {status} от нормативного графика "
                f"(норматив {fmt_num(norm_days)} сут от начала работ {start_str} &rarr; план {plan_end_str}; "
                f"факт. прогноз окончания — {forecast_str})."
            )
        else:
            lines.append(
                f"Бурение идёт точно по нормативному графику "
                f"(норматив {fmt_num(norm_days)} сут от начала работ {start_str} &rarr; план {plan_end_str}; "
                f"факт. прогноз окончания — {forecast_str})."
            )
    else:
        lines.append('<span class="muted">Недостаточно данных для расчёта отклонения от нормативного графика.</span>')

    if commercial_speed and (commercial_speed.get("plan") not in (None, "") or commercial_speed.get("fact") not in (None, "")):
        lines.append(
            f"Коммерческая скорость: план {fmt_num(commercial_speed.get('plan'))} м/ст.мес, "
            f"факт {fmt_num(commercial_speed.get('fact'))} м/ст.мес."
        )

    if stage_deviation:
        # deviation_h > 0 = факт дольше плана (тянет к отставанию),
        # deviation_h < 0 = факт быстрее плана (тянет к опережению).
        overruns = sorted([s for s in stage_deviation if s["deviation_h"] and s["deviation_h"] > 0],
                           key=lambda s: -s["deviation_h"])[:3]
        savings = sorted([s for s in stage_deviation if s["deviation_h"] and s["deviation_h"] < 0],
                          key=lambda s: s["deviation_h"])[:3]

        def fmt_list(items, n):
            return ", ".join(f"{esc(s['stage'])} ({fmt_signed(s['deviation_h'], 1)} ч)" for s in items[:n]) or None

        if status == "отставание":
            main_txt = fmt_list(overruns, 3)
            comp_txt = fmt_list(savings, 2)
            if main_txt and comp_txt:
                lines.append(f"Основной вклад в отставание — {main_txt}; частично компенсировано опережением по: {comp_txt}.")
            elif main_txt:
                lines.append(f"Основной вклад в отставание — {main_txt}.")
        elif status == "опережение":
            main_txt = fmt_list(savings, 3)
            comp_txt = fmt_list(overruns, 2)
            if main_txt and comp_txt:
                lines.append(f"Опережение обеспечено в основном за счёт: {main_txt}; частично компенсировано отставанием по: {comp_txt}.")
            elif main_txt:
                lines.append(f"Опережение обеспечено в основном за счёт: {main_txt}.")
        else:
            any_txt = fmt_list(overruns, 2)
            any_neg = fmt_list(savings, 2)
            if any_txt or any_neg:
                bits = [t for t in (any_txt, any_neg) if t]
                lines.append(f"Отклонения по отдельным этапам: {'; '.join(bits)} — в целом график выдерживается.")

    body = "".join(f"<p>{ln}</p>" for ln in lines)
    chart_html = ""
    if depth_chart_svg:
        chart_html = f'<div class="depth-chart">{depth_chart_svg}</div>'
    return f'<div class="schedule-block">{body}</div>{chart_html}'


def render_npv_summary_line(depth, ops):
    daily_npv = depth.get("daily_npv")
    parties = collect_responsible_parties(ops)
    parties_txt = "; ".join(esc(p) for p in parties) if parties else "—"
    return (
        '<div class="npv-summary">'
        f'<div class="npv-summary-item"><span class="k">НПВ за сутки</span><span class="v strong">{fmt_num(daily_npv)} ч</span></div>'
        f'<div class="npv-summary-item"><span class="k">Виновная сторона</span><span class="v strong">{parties_txt}</span></div>'
        '</div>'
    )


def _rig_brigade_line(h):
    """Вторая строка в заголовке карточки скважины — бригада и тип буровой
    установки (лицевая страница СР, поля 'Буровая бригада'/'Буровая
    установка'), второстепенным по значимости шрифтом под названием
    месторождения/куста. Если оба поля пусты — строка не выводится.

    Значение поля 'Буровая бригада' в источнике уже само по себе номер вида
    "ББ-22" (ББ = Буровая Бригада), поэтому слово "Бригада" перед ним НЕ
    добавляется — иначе получилось бы задвоение ("Бригада ББ-22" читается
    как "Бригада Буровая Бригада-22")."""
    brigade = _clean_str_local(h.get("brigade"))
    rig = _clean_str_local(h.get("rig"))
    parts = []
    if brigade:
        parts.append(esc(brigade))
    if rig:
        parts.append(f"БУ {esc(rig)}")
    if not parts:
        return ""
    return f'<div class="title-sub">{" &middot; ".join(parts)}</div>'


def render_well_card(idx, d):
    h = d["header"]
    depth = d["depth"]
    anchor = f"well-{idx}"
    title = f"{esc(h.get('field'))} — куст {esc(h.get('pad'))}, скв. {esc(h.get('well_no'))}"
    rig_brigade_line = _rig_brigade_line(h)

    daily_footage = depth.get("daily_footage")
    daily_npv = depth.get("daily_npv")

    contacts_summary = "Буровая установка, бригада, мастер, супервайзер — контакты"
    schedule_summary = "Отклонение от графика (глубина-день)"
    ops_summary = "Суточные операции — НПВ по категориям"
    npv_summary = "НПВ по ответственным сторонам (накопительно)"
    mud_summary = "Параметры промывочной жидкости (план/факт)"

    return f"""
<section class="card" id="{anchor}">
  <button type="button" class="scroll-top-btn" title="Наверх страницы" aria-label="Наверх страницы"
    onclick="window.scrollTo({{top: 0, behavior: 'smooth'}});">&#8593;</button>
  <div class="card-head">
    <div class="title-block">
      <h2>{title}</h2>
      {rig_brigade_line}
    </div>
    <div class="badges">
      <span class="badge">Рапорт &#8470;{esc(h.get('report_no'))} от {fmt_date(h.get('report_date'))}</span>
      <span class="badge metric">Проходка за сутки: {fmt_num(daily_footage)} м</span>
      <span class="badge metric warn">НПВ за сутки: {fmt_num(daily_npv)} ч</span>
    </div>
  </div>

  <div class="meta-grid">
    <div class="kv"><span class="k">Заказчик</span><span class="v">{esc(h.get('customer')) or '—'}</span></div>
    <div class="kv"><span class="k">Месторождение</span><span class="v">{esc(h.get('field')) or '—'}</span></div>
    <div class="kv"><span class="k">Куст</span><span class="v">{esc(h.get('pad')) or '—'}</span></div>
    <div class="kv"><span class="k">&#8470; скважины</span><span class="v">{esc(h.get('well_no')) or '—'}</span></div>
  </div>

  {render_details(contacts_summary, render_contacts(h))}

  {render_details(schedule_summary, render_schedule_block(h, d.get("commercial_speed"), d.get("stage_deviation"), d.get("depth_chart_svg")))}

  <div class="two-col">
    <div>
      <h3>Конструкция скважины (план/факт)</h3>
      {render_construction_table(d["construction"])}
    </div>
    <div>
      <h3>Долото (текущее)</h3>
      {render_bit_block(d["bit"])}
    </div>
  </div>

  <div class="kv-grid two-line">
    <div class="kv"><span class="k">Забой на 24:00, м</span><span class="v">{fmt_num(depth.get('depth_24'))}</span></div>
    <div class="kv"><span class="k">Выполняемые работы на 24:00</span><span class="v">{esc(depth.get('work_at_24')) or '—'}</span></div>
    <div class="kv"><span class="k">Забой на 06:00, м</span><span class="v">{fmt_num(depth.get('depth_06'))}</span></div>
    <div class="kv"><span class="k">Выполняемые работы на 06:00</span><span class="v">{esc(depth.get('work_at_06')) or '—'}</span></div>
    <div class="kv wide"><span class="k">Планируемые работы на сутки</span><span class="v">{fmt_multi(depth.get('planned_work'), sep=', ') if depth.get('planned_work') else '—'}</span></div>
  </div>

  {render_npv_summary_line(depth, d["ops"])}
  {render_details(ops_summary, render_ops_events(d["ops"]))}
  {render_details(npv_summary, render_npv_table(d["npv"]))}
  {render_details(mud_summary, render_mud_params_block(d.get("mud_params")))}

  <div class="src">Источник: {esc(d['source_file'])}</div>
</section>"""


# ---------------------------------------------------------------------------
# Сводная таблица (шапка отчёта)
# ---------------------------------------------------------------------------

def render_lag_cell(lag):
    """Ячейка 'Опережение/отставание от графика' для сводной таблицы:
    красный текст — отставание (lag > 0), зелёный — опережение (lag < 0),
    серый — точно по графику или нет данных."""
    if not isinstance(lag, (int, float)):
        return '<span class="lag-cell lag-none">—</span>'
    if lag > 0.005:
        cls, label = "lag-behind", "отставание"
    elif lag < -0.005:
        cls, label = "lag-ahead", "опережение"
    else:
        cls, label = "lag-none", "по графику"
    return (
        f'<span class="lag-cell {cls}">'
        f'<span class="lag-value">{fmt_signed(lag, 2)} сут</span>'
        f'<span class="lag-label">{label}</span>'
        f'</span>'
    )


def _sum_daily_footage(items):
    """Суммарная суточная проходка по списку (idx, d) — None, если ни у
    одной скважины нет числового значения (чтобы не показывать обманчивый
    0 там, где на самом деле просто нет данных)."""
    total = 0
    has_value = False
    for _, d in items:
        v = d["depth"].get("daily_footage")
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            total += v
            has_value = True
    return total if has_value else None


def render_stage_cell(stage):
    """Ячейка колонки 'Этап': либо просто текст этапа, либо (когда
    stage_classifier не смог однозначно определить этап по пустой
    'Операции' в хронологии) предупреждающий значок "!" с всплывающей
    подсказкой при наведении — см. stage_classifier.MANUAL_STAGE_RULES."""
    if not stage:
        return '<span class="muted">&mdash;</span>'
    text = stage.get("text")
    warning = stage.get("warning")
    parts = []
    if text:
        parts.append(esc(text))
    if warning:
        parts.append(f'<span class="stage-warn" title="{esc(warning)}">&#33;</span>')
    return "".join(parts) if parts else '<span class="muted">&mdash;</span>'


def _summary_rows_html(group_items):
    rows = []
    for idx, d in group_items:
        h = d["header"]
        depth = d["depth"]
        anchor = f"well-{idx}"
        npv = depth.get("daily_npv") or 0
        npv_cls = ' class="warn"' if isinstance(npv, (int, float)) and npv > 0 else ""
        stage = d.get("current_stage")
        parties = collect_responsible_parties(d.get("ops"))
        parties_txt = "; ".join(esc(p) for p in parties) if parties else "—"
        rows.append(
            "<tr>"
            f"<td>{esc(h.get('field'))}</td>"
            f"<td>{esc(h.get('pad'))}</td>"
            f"<td>{esc(h.get('well_no'))}</td>"
            f"<td>{fmt_num(depth.get('daily_footage'))}</td>"
            f"<td>{render_stage_cell(stage)}</td>"
            f"<td{npv_cls}>{fmt_num(depth.get('daily_npv'))}</td>"
            f"<td>{parties_txt}</td>"
            f"<td>{render_lag_cell(h.get('lag_days'))}</td>"
            f'<td><a href="#{anchor}">Подробнее &rarr;</a></td>'
            "</tr>"
        )
    return "".join(rows)


SUMMARY_TABLE_HEAD = """
    <tr>
      <th>Месторождение</th><th>Куст</th><th>&#8470; скважины</th>
      <th>Проходка, м/сут</th><th>Этап</th><th>НПВ, ч/сут</th><th>Ответственная сторона за НПВ</th>
      <th>Опережение/отставание</th><th></th>
    </tr>"""

# Фиксированная ширина колонок (в %, см. .tbl.summary { table-layout: fixed }
# в CSS) — без неё браузер сам подбирает ширину под содержимое КАЖДОЙ
# таблицы кластера отдельно, из-за чего одноимённые колонки в блоках
# разных кластеров съезжают и не выстраиваются друг под другом.
SUMMARY_TABLE_COLS = """
    <colgroup>
      <col style="width:16%"><col style="width:7%"><col style="width:9%">
      <col style="width:9%"><col style="width:8%"><col style="width:8%">
      <col style="width:23%"><col style="width:11%"><col style="width:9%">
    </colgroup>"""

# Минимальная ширина таблицы в пикселях — на узких экранах (телефон) не даёт
# table-layout:fixed сжать колонки до нечитаемого состояния (текст типа
# "Куст" переносится по одной букве в строку); вместо этого таблица шире
# экрана и прокручивается по горизонтали в обёртке .table-scroll (см. JS
# в конце страницы), проценты колонок при этом остаются одинаковыми.
SUMMARY_TABLE_MINWIDTH = 760


def render_summary_table(items, cluster_overrides=None):
    if not items:
        return '<p class="muted">Нет ни одного успешно разобранного рапорта.</p>'

    total_footage = _sum_daily_footage(items)
    total_html = (
        '<div class="summary-total">'
        f'<span class="summary-total-label">Общая проходка за сутки по всем скважинам</span>'
        f'<span class="summary-total-value">{fmt_num(total_footage)} м</span>'
        f'<span class="summary-total-count">{len(items)} скв.</span>'
        '</div>'
    )

    groups = group_by_cluster(items, cluster_overrides)
    blocks = []
    for cluster_label, group_items in groups:
        cluster_footage = _sum_daily_footage(group_items)
        rows_html = _summary_rows_html(group_items)
        blocks.append(f"""
<div class="cluster-block">
  <div class="cluster-header collapsible" onclick="srToggleClusterBlock(this)">
    <span class="cluster-toggle-icon">&#9660;</span>
    <span class="cluster-name">Кластер: {esc(cluster_label)}</span>
    <span class="cluster-metric">Проходка за сутки: <b>{fmt_num(cluster_footage)} м</b></span>
    <span class="cluster-metric">Скважин: {len(group_items)}</span>
  </div>
  <table class="tbl summary" style="min-width:{SUMMARY_TABLE_MINWIDTH}px">
    {SUMMARY_TABLE_COLS}
    <thead>{SUMMARY_TABLE_HEAD}</thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>""")

    return total_html + "".join(blocks)


def _bits_rows_html(group_items):
    rows = []
    for idx, d in group_items:
        h = d["header"]
        depth = d["depth"]
        bit = d.get("bit")
        anchor = f"well-{idx}"
        footage_cell = render_bit_footage_cell(bit) if bit else fmt_num(depth.get("daily_footage"))
        speed_cell = render_bit_speed_cell(bit) if bit else "—"
        rows.append(
            "<tr>"
            f"<td>{esc(h.get('field'))}</td>"
            f"<td>{esc(h.get('pad'))}</td>"
            f"<td>{esc(h.get('well_no'))}</td>"
            f"<td>{footage_cell}</td>"
            f"<td>{speed_cell}</td>"
            f"<td>{esc(bit.get('maker')) if bit and bit.get('maker') else '—'}</td>"
            f"<td>{esc(depth.get('work_at_06')) or '—'}</td>"
            f'<td><a href="#{anchor}" onclick="srShowTab(\'wells\')">Подробнее &rarr;</a></td>'
            "</tr>"
        )
    return "".join(rows)


BITS_TABLE_HEAD = """
    <tr>
      <th>Месторождение</th><th>Куст</th><th>&#8470; скважины</th>
      <th>Проходка, м (сутки / рейс)</th><th>МСП, м/ч (сутки / рейс / план)</th><th>Изготовитель долота</th>
      <th>Выполняемые работы на 06:00</th><th></th>
    </tr>"""

BITS_TABLE_COLS = """
    <colgroup>
      <col style="width:14%"><col style="width:7%"><col style="width:9%">
      <col style="width:14%"><col style="width:14%"><col style="width:14%">
      <col style="width:19%"><col style="width:9%">
    </colgroup>"""

BITS_TABLE_MINWIDTH = 720


def _render_circ_trip_cell(circ_trip, resource):
    """Ячейка 'Наработка цирк. общ. за рейс' — подсвечивается зелёным, если
    накопленная наработка ещё в пределах ресурса ВЗД/осциллятора/СБТ данной
    модели, и красным, если ресурс уже превышен. Если ресурс для этой
    модели неизвестен (не нашлось однозначного совпадения со справочником —
    см. vzd_reference.py) или само значение наработки отсутствует —
    подсветка не применяется, показывается как обычно."""
    text = fmt_num(circ_trip)
    if (
        isinstance(circ_trip, (int, float)) and not isinstance(circ_trip, bool)
        and isinstance(resource, (int, float)) and not isinstance(resource, bool)
    ):
        cls = "circ-ok" if circ_trip <= resource else "circ-over"
        return f'<span class="{cls}">{text}</span>'
    return text


def _component_rows_html(group_items, component_key):
    """Общая функция построения строк сводной таблицы по элементу КНБК
    (ВЗД/осциллятор), структура которых одинакова — см.
    render_vzd_table/render_osc_table."""
    rows = []
    for idx, d in group_items:
        h = d["header"]
        anchor = f"well-{idx}"
        blocks = d.get("narabotka") or []
        for b in blocks:
            comp = b.get(component_key)
            if not comp:
                continue
            resource = comp.get("resource")
            rows.append(
                "<tr>"
                f"<td>{esc(h.get('field'))}</td>"
                f"<td>{esc(h.get('pad'))}</td>"
                f"<td>{esc(h.get('well_no'))}</td>"
                f"<td>{esc(comp.get('model')) or '—'}</td>"
                f"<td>{esc(comp.get('maker')) or '—'}</td>"
                f"<td>{esc(comp.get('serial')) or '—'}</td>"
                f"<td>{fmt_num(b.get('zaboi_from'))} – {fmt_num(b.get('zaboi_to'))}</td>"
                f"<td>{fmt_num(comp.get('circ_daily'))}</td>"
                f"<td>{_render_circ_trip_cell(comp.get('circ_trip'), resource)}</td>"
                f"<td>{fmt_num(resource) if resource else '—'}</td>"
                f'<td><a href="#{anchor}" onclick="srShowTab(\'wells\')">Подробнее &rarr;</a></td>'
                "</tr>"
            )
    return "".join(rows)


def _component_table_head(title_col):
    return f"""
    <tr>
      <th>Месторождение</th><th>Куст</th><th>&#8470; скважины</th>
      <th>{title_col}</th><th>Изготовитель</th><th>Серийный номер</th>
      <th>Забой рейса от&ndash;до, м</th>
      <th>Наработка цирк. сут., ч</th><th>Наработка цирк. общ. за рейс, ч</th>
      <th>Ресурс, ч</th>
      <th></th>
    </tr>"""


_COMPONENT_TABLE_COLS = """
    <colgroup>
      <col style="width:12%"><col style="width:6%"><col style="width:7%">
      <col style="width:10%"><col style="width:9%"><col style="width:9%">
      <col style="width:11%">
      <col style="width:9%"><col style="width:10%">
      <col style="width:8%">
      <col style="width:9%">
    </colgroup>"""

_COMPONENT_TABLE_MINWIDTH = 980


def _sbt_rows_html(group_items):
    """Строки сводной таблицы по СБТ (вкладка 'Инструмент') — только для
    блоков секции 'Хвостовик' (см. extract_narabotka: поле 'sbt' заполняется
    только для таких блоков). Столбец суточной наработки не показывается —
    только наработка цирк. общ. за рейс (по указанию пользователя)."""
    rows = []
    for idx, d in group_items:
        h = d["header"]
        anchor = f"well-{idx}"
        blocks = d.get("narabotka") or []
        for b in blocks:
            comp = b.get("sbt")
            if not comp:
                continue
            resource = comp.get("resource")
            rows.append(
                "<tr>"
                f"<td>{esc(h.get('field'))}</td>"
                f"<td>{esc(h.get('pad'))}</td>"
                f"<td>{esc(h.get('well_no'))}</td>"
                f"<td>{esc(comp.get('model')) or '—'}</td>"
                f"<td>{esc(comp.get('serial')) or '—'}</td>"
                f"<td>{fmt_num(b.get('zaboi_from'))} – {fmt_num(b.get('zaboi_to'))}</td>"
                f"<td>{_render_circ_trip_cell(comp.get('circ_trip'), resource)}</td>"
                f"<td>{fmt_num(resource) if resource else '—'}</td>"
                f'<td><a href="#{anchor}" onclick="srShowTab(\'wells\')">Подробнее &rarr;</a></td>'
                "</tr>"
            )
    return "".join(rows)


SBT_TABLE_HEAD = """
    <tr>
      <th>Месторождение</th><th>Куст</th><th>&#8470; скважины</th>
      <th>Компоновка (СБТ)</th><th>Серийный номер</th>
      <th>Забой рейса от&ndash;до, м</th>
      <th>Наработка на инструмент с вращением, ч</th>
      <th>Ресурс, ч</th>
      <th></th>
    </tr>"""

SBT_TABLE_COLS = """
    <colgroup>
      <col style="width:14%"><col style="width:7%"><col style="width:8%">
      <col style="width:16%"><col style="width:12%">
      <col style="width:12%">
      <col style="width:14%">
      <col style="width:8%">
      <col style="width:9%">
    </colgroup>"""

SBT_TABLE_MINWIDTH = 760


def _render_clustered_table(items, cluster_overrides, rows_fn, thead_html, empty_msg, cols_html="", min_width=0):
    """Оборачивает построчную функцию rows_fn(group_items) -> html<tr>...
    в блоки по кластерам (см. group_by_cluster), в едином стиле с
    render_summary_table. Кластеры, где rows_fn не вернула ни одной строки
    (нет данных по этому элементу ни у одной скважины кластера), не
    показываются вовсе. cols_html — общий <colgroup> с фиксированными
    ширинами колонок (см. *_TABLE_COLS), чтобы одноимённые колонки во всех
    кластерах этой вкладки были одной ширины и не съезжали друг относительно
    друга (см. .tbl.summary { table-layout: fixed } в CSS). min_width —
    минимальная ширина таблицы в пикселях (см. *_TABLE_MINWIDTH) — не даёт
    колонкам схлопнуться до нечитаемого состояния на телефоне; на узких
    экранах таблица вместо этого просто шире экрана и скроллится по
    горизонтали в обёртке .table-scroll."""
    groups = group_by_cluster(items, cluster_overrides)
    blocks = []
    minw_attr = f' style="min-width:{min_width}px"' if min_width else ""
    for cluster_label, group_items in groups:
        rows_html = rows_fn(group_items)
        if not rows_html:
            continue
        well_count = len(set(re.findall(r'href="#(well-\d+)"', rows_html)))
        blocks.append(f"""
<div class="cluster-block">
  <div class="cluster-header">
    <span class="cluster-name">Кластер: {esc(cluster_label)}</span>
    <span class="cluster-metric">Скважин: {well_count}</span>
  </div>
  <table class="tbl summary"{minw_attr}>
    {cols_html}
    <thead>{thead_html}</thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>""")
    if not blocks:
        return f'<p class="muted">{empty_msg}</p>'
    return "".join(blocks)


def render_bits_table(items, cluster_overrides=None):
    return _render_clustered_table(
        items, cluster_overrides, _bits_rows_html, BITS_TABLE_HEAD,
        "Нет ни одного успешно разобранного рапорта.",
        cols_html=BITS_TABLE_COLS, min_width=BITS_TABLE_MINWIDTH,
    )


def render_vzd_table(items, cluster_overrides=None):
    """Сводная таблица по ВЗД (винтовым забойным двигателям) с листа
    'Наработка' — по одной строке на каждый обнаруженный блок 'КНБК и БТ' с
    ВЗД (может быть несколько строк на одну скважину, если КНБК меняли в
    течение суток — см. extract_narabotka в sr_parser.py), сгруппированная
    по кластерам месторождений."""
    return _render_clustered_table(
        items, cluster_overrides,
        lambda gi: _component_rows_html(gi, "vzd"),
        _component_table_head("Модель ВЗД"),
        "Нет данных по ВЗД ни по одной скважине (за эти сутки подъём/спуск КНБК с ВЗД не проводился).",
        cols_html=_COMPONENT_TABLE_COLS, min_width=_COMPONENT_TABLE_MINWIDTH,
    )


def render_osc_table(items, cluster_overrides=None):
    """Сводная таблица по осцилляторам с листа 'Наработка' — по одной строке
    на каждый обнаруженный блок 'КНБК и БТ' с осциллятором (аналогично
    render_vzd_table), сгруппированная по кластерам месторождений."""
    return _render_clustered_table(
        items, cluster_overrides,
        lambda gi: _component_rows_html(gi, "osc"),
        _component_table_head("Модель осциллятора"),
        "Нет данных по осциллятору ни по одной скважине (за эти сутки подъём/спуск КНБК с осциллятором не проводился).",
        cols_html=_COMPONENT_TABLE_COLS, min_width=_COMPONENT_TABLE_MINWIDTH,
    )


def render_sbt_table(items, cluster_overrides=None):
    """Сводная таблица по СБТ (вкладка 'Инструмент') — только секция
    'Хвостовик' (см. _sbt_rows_html), сгруппированная по кластерам
    месторождений."""
    return _render_clustered_table(
        items, cluster_overrides, _sbt_rows_html, SBT_TABLE_HEAD,
        "Нет данных по инструменту (СБТ) в секции «Хвостовик» ни по одной скважине.",
        cols_html=SBT_TABLE_COLS, min_width=SBT_TABLE_MINWIDTH,
    )


_ROLE_SUFFIX_RE = re.compile(r"\s*\(([^()]*)\)\s*$")
_ORG_FORM_RE = re.compile(r"^\s*(ООО|ОАО|ЗАО|ПАО|АО|ИП)\b[\s.]*", re.IGNORECASE)


def _split_company_and_suffix(name):
    """Отделяет завершающую скобочную пометку (роль/вид сервиса/служба) от
    названия компании, напр. 'АО "Технологии ОФС" (ННБ)' ->
    ('АО "Технологии ОФС"', 'ННБ'). Если скобок нет — (имя, None)."""
    s = _clean_str_local(name)
    m = _ROLE_SUFFIX_RE.search(s)
    if not m:
        return s, None
    base = s[:m.start()].strip()
    suffix = _clean_str_local(m.group(1))
    return (base or s), (suffix or None)


def _is_inkservis(base_company):
    """True, если base_company (без скобочной пометки) — это ООО
    'ИНК-СЕРВИС', независимо от вида кавычек/написания."""
    s = _ORG_FORM_RE.sub("", _clean_str_local(base_company))
    s = re.sub(r"[^А-ЯЁA-Z]", "", s.upper()).replace("Ё", "Е")
    return s == "ИНКСЕРВИС"


_HOURS_TEXT_RE = re.compile(r"^(\d{1,4}):(\d{2})$")


def _parse_hours_text(s):
    """Разбирает значение столбца 'Продолж.' блока 'Суточные операции'
    (лист 'Суточный отчет', см. extract_ops) в часы (float) — в реальных
    файлах это текст 'ЧЧ:ММ'; на всякий случай поддержано и обычное число."""
    s = _clean_str_local(s)
    if not s:
        return None
    m = _HOURS_TEXT_RE.match(s)
    if m:
        return int(m.group(1)) + int(m.group(2)) / 60.0
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


_WS_RE = re.compile(r"\s+")
_REASON_LIMIT = 140


def _short_reason(comment):
    """Краткая причина НПВ для строки в раскрывающемся списке рейтинга —
    из столбца 'Комментарий' блока 'Суточные операции' (см. extract_ops),
    очищенная от переносов строк и обрезанная до разумной длины (полный
    текст комментария всегда доступен на самой карточке скважины)."""
    s = _clean_str_local(comment)
    if not s:
        return ""
    s = _WS_RE.sub(" ", s).strip()
    if len(s) > _REASON_LIMIT:
        s = s[:_REASON_LIMIT].rstrip() + "…"
    return s


def _collect_npv_by_company(group_items):
    """Рейтинг виновных сторон по НПВ за отчётные сутки (вкладка 'Рейтинг
    НПВ') — источник: события блока 'Суточные операции' листа 'Суточный
    отчет' (те же, что и 'Виновная сторона'/'Суточные операции — НПВ по
    категориям' на карточке скважины), а НЕ накопительный лист 'НПВ' (там
    нет разбивки по конкретным суткам).

    Внутри одной строки блока 'Категория'/'Продолж.'/'Ответственный' могут
    содержать несколько пронумерованных вида '(N.M)' под-записей (см.
    split_ru_entries) — продолжительность сопоставляется с ответственным
    по порядку; если количество записей не совпадает (в предоставленных
    файлах такого не встречалось, но на всякий случай) — при одной
    продолжительности на несколько сторон она приписывается целиком
    каждой, при нескольких продолжительностях на одну сторону — суммируются.

    Группировка по строкам рейтинга — компания ВМЕСТЕ с видом сервиса/ролью
    в скобках (напр. 'АО "Технологии ОФС" (ННБ)' и 'АО "Технологии ОФС"
    (Цементирование)' — это две разные строки), кроме ООО 'ИНК-СЕРВИС': для
    неё скобочная пометка — это внутренняя служба (см. _is_inkservis), и
    строка рейтинга — одна ('ООО «ИНК-СЕРВИС»'), а разбивка по службам
    уходит на второй уровень раскрывающегося списка.

    В отличие от предыдущей версии, каждое событие НПВ остаётся отдельной
    записью (не суммируется по скважине) — если у стороны на одной скважине
    за сутки несколько событий НПВ, каждое показывается отдельной строкой
    со своей причиной (см. render_ops_events/_npv_rating_rows_html).

    Возвращает список записей рейтинга, отсортированный по убыванию 'hours':
      - обычная сторона: {'company', 'hours', 'is_inkservis': False,
        'entries': [{'anchor','label','hours','reason'}, ...]}
        (entries — по убыванию часов);
      - ООО 'ИНК-СЕРВИС': {'company', 'hours', 'is_inkservis': True,
        'services': [{'service','hours','entries':[...]}, ...]}
        (services и entries внутри — по убыванию часов)."""
    by_company = {}
    for idx, d in group_items:
        h = d["header"]
        anchor = f"well-{idx}"
        well_label = f"{esc(h.get('field'))} — куст {esc(h.get('pad'))}, скв. {esc(h.get('well_no'))}"
        for e in d.get("ops") or []:
            resp_entries = split_ru_entries(e.get("responsible"))
            dur_entries = split_ru_entries(e.get("duration"))
            if not resp_entries or not dur_entries:
                continue
            if len(resp_entries) == len(dur_entries):
                pairs = list(zip(resp_entries, dur_entries))
            elif len(dur_entries) == 1:
                pairs = [(r, dur_entries[0]) for r in resp_entries]
            elif len(resp_entries) == 1:
                pairs = [(resp_entries[0], du) for du in dur_entries]
            else:
                pairs = list(zip(resp_entries, dur_entries))
            reason = _short_reason(e.get("comment"))
            for resp_raw, dur_raw in pairs:
                hours = _parse_hours_text(dur_raw)
                if hours is None:
                    continue
                base, suffix = _split_company_and_suffix(resp_raw)
                if not base:
                    continue
                if _is_inkservis(base):
                    bucket = by_company.setdefault(base, {"hours": 0.0, "is_inkservis": True, "services": {}})
                    bucket["hours"] += hours
                    service = suffix or "Без указания службы"
                    sbucket = bucket["services"].setdefault(service, {"hours": 0.0, "entries": []})
                    sbucket["hours"] += hours
                    sbucket["entries"].append({"anchor": anchor, "label": well_label, "hours": hours, "reason": reason})
                else:
                    display = _clean_str_local(resp_raw)
                    bucket = by_company.setdefault(display, {"hours": 0.0, "is_inkservis": False, "entries": []})
                    bucket["hours"] += hours
                    bucket["entries"].append({"anchor": anchor, "label": well_label, "hours": hours, "reason": reason})

    rating = []
    for company, data in by_company.items():
        entry = {"company": company, "hours": data["hours"], "is_inkservis": data["is_inkservis"]}
        if data["is_inkservis"]:
            services = []
            for svc_name, svc_data in data["services"].items():
                ents = sorted(svc_data["entries"], key=lambda x: -x["hours"])
                services.append({"service": svc_name, "hours": svc_data["hours"], "entries": ents})
            services.sort(key=lambda s: -s["hours"])
            entry["services"] = services
            entry["well_count"] = len({e["anchor"] for s in services for e in s["entries"]})
        else:
            ents = sorted(data["entries"], key=lambda x: -x["hours"])
            entry["entries"] = ents
            entry["well_count"] = len({e["anchor"] for e in ents})
        rating.append(entry)
    rating.sort(key=lambda r: -r["hours"])
    return rating


def _npv_entry_li(e):
    reason_html = f' &mdash; {esc(e["reason"])}' if e.get("reason") else ""
    return (
        f'<li><a href="#{e["anchor"]}" onclick="srShowTab(\'wells\')">{e["label"]}</a> '
        f'&mdash; <span class="warn">{fmt_num(e["hours"])} ч</span>{reason_html}</li>'
    )


def _npv_rating_rows_html(rating):
    rows = []
    for rank, r in enumerate(rating, start=1):
        if r["is_inkservis"]:
            service_blocks = "".join(
                f'<details class="npv-rating-details npv-rating-details-nested">'
                f'<summary>{esc(svc["service"])} &mdash; <span class="warn">{fmt_num(svc["hours"])} ч</span></summary>'
                f'<ul>{"".join(_npv_entry_li(e) for e in svc["entries"])}</ul></details>'
                for svc in r["services"]
            )
            drilldown = (
                f'<details class="npv-rating-details"><summary>Подробнее '
                f'({r["well_count"]} скв., служб: {len(r["services"])})</summary>'
                f'{service_blocks}</details>'
            )
        else:
            items_html = "".join(_npv_entry_li(e) for e in r["entries"])
            drilldown = (
                f'<details class="npv-rating-details"><summary>Подробнее ({r["well_count"]} скв.)</summary>'
                f'<ul>{items_html}</ul></details>'
            )
        rows.append(
            "<tr>"
            f"<td>{rank}</td>"
            f"<td>{esc(r['company'])}</td>"
            f'<td class="warn">{fmt_num(r["hours"])}</td>'
            f"<td>{drilldown}</td>"
            "</tr>"
        )
    return "".join(rows)


NPV_RATING_TABLE_HEAD = """
    <tr>
      <th>&#8470;</th><th>Виновная сторона</th><th>Часы НПВ (общие)</th><th>Скважины с НПВ</th>
    </tr>"""

NPV_RATING_TABLE_COLS = """
    <colgroup>
      <col style="width:5%"><col style="width:28%">
      <col style="width:15%"><col style="width:52%">
    </colgroup>"""

NPV_RATING_TABLE_MINWIDTH = 560


def render_npv_rating_tab(items, cluster_overrides=None):
    """Вкладка 'Рейтинг НПВ' — рейтинг компаний-подрядчиков по суммарному
    НПВ за отчётные сутки, отдельно по каждому кластеру месторождений (как
    и остальные сводные вкладки), по убыванию часов. См.
    _collect_npv_by_company."""
    if not items:
        return '<p class="muted">Нет ни одного успешно разобранного рапорта.</p>'

    groups = group_by_cluster(items, cluster_overrides)
    blocks = []
    for cluster_label, group_items in groups:
        rating = _collect_npv_by_company(group_items)
        if not rating:
            continue
        total_hours = sum(r["hours"] for r in rating)
        rows_html = _npv_rating_rows_html(rating)
        blocks.append(f"""
<div class="cluster-block">
  <div class="cluster-header">
    <span class="cluster-name">Кластер: {esc(cluster_label)}</span>
    <span class="cluster-metric">НПВ за сутки: <b>{fmt_num(total_hours)} ч</b></span>
    <span class="cluster-metric">Компаний: {len(rating)}</span>
  </div>
  <table class="tbl summary" style="min-width:{NPV_RATING_TABLE_MINWIDTH}px">
    {NPV_RATING_TABLE_COLS}
    <thead>{NPV_RATING_TABLE_HEAD}</thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>""")

    if not blocks:
        return '<p class="muted">НПВ за эти сутки ни по одной скважине не зафиксировано.</p>'
    return "".join(blocks)


def render_reference_modal():
    """Содержимое модального окна 'Справочная' — вся исходная таблица по
    ВЗД (производитель/тип/коэф. об-литр/оборот на литр/ресурс) и таблица по
    осцилляторам, как их предоставил пользователь (см. vzd_reference.py)."""
    vzd_rows = "".join(
        "<tr>"
        f"<td>{esc(maker)}</td><td>{esc(model)}</td>"
        f"<td>{fmt_num(koef)}</td><td>{fmt_num(oborot)}</td><td>{fmt_num(resource)}</td>"
        "</tr>"
        for maker, model, koef, oborot, resource in vzd_reference.VZD_RESOURCE_TABLE
    )
    osc_rows = "".join(
        f"<tr><td>{esc(maker)}</td><td>{esc(label)}</td><td>{fmt_num(resource)}</td></tr>"
        for maker, label, resource in vzd_reference.OSC_RESOURCE_TABLE
    )
    return f"""
<p class="summary-title" style="margin-top:0">Ресурс ВЗД (расчёт оборотов на литр промывочной жидкости)</p>
<table class="tbl">
  <thead><tr><th>Производитель ВЗД</th><th>Тип ВЗД</th><th>Коэф. об/литр</th><th>Оборот на литр</th><th>Ресурс, ч</th></tr></thead>
  <tbody>{vzd_rows}</tbody>
</table>
<p class="summary-title">Ресурс осциллятора</p>
<table class="tbl">
  <thead><tr><th>Производитель</th><th>Типоразмер</th><th>Ресурс, ч</th></tr></thead>
  <tbody>{osc_rows}</tbody>
</table>
"""


# ---------------------------------------------------------------------------
# Логотип
# ---------------------------------------------------------------------------

def _logo_data_uri():
    for path in (LOGO_PATH, LOGO_PATH_FLAT):
        try:
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            return f"data:image/png;base64,{b64}"
        except OSError:
            continue
    return None


# ---------------------------------------------------------------------------
# Полная страница
# ---------------------------------------------------------------------------

CSS = """
:root {
  --brand-dark: #1b1f22;      /* графит логотипа */
  --brand-green: #265841;     /* зелёный логотипа ИНК */
  --brand-green-dark: #163325;
  --brand-green-tint: #e9f1ec;
  --accent-warm: #c14a26;     /* тёплый акцент (терракот irkutskoil.ru) */
  --bg: #f2f5f3;
  --card-bg: #ffffff;
  --text: #1b201d;
  --muted: #667066;
  --border: #dfe6e1;
  --warn-bg: #fbeae4;
  --warn-text: #a13c1f;
}
* { box-sizing: border-box; }
body {
  margin: 0; background: var(--bg); color: var(--text);
  font-family: "Segoe UI", Arial, sans-serif; font-size: 14px; line-height: 1.45;
}
a { color: var(--brand-green); }
header.top {
  background: linear-gradient(135deg, var(--brand-dark) 0%, var(--brand-green-dark) 62%, var(--brand-green) 130%);
  color: #fff; padding: 22px 28px;
}
.top-inner { display: flex; align-items: center; justify-content: space-between; gap: 20px; max-width: 1180px; margin: 0 auto; }
.top-text h1 { margin: 0 0 6px 0; font-size: 22px; font-weight: 700; }
.top-text .sub { color: #cfe0d6; font-size: 13px; }
.logo-plate { background: #fff; border-radius: 12px; padding: 6px 10px; box-shadow: 0 2px 8px rgba(0,0,0,.25); flex-shrink: 0; }
.logo-plate img { display: block; height: 46px; width: auto; }
main { max-width: 1180px; margin: 0 auto; padding: 24px 20px 60px; }
h2 { margin: 0; font-size: 18px; color: var(--brand-dark); }
h3 { font-size: 13px; color: var(--brand-green); margin: 18px 0 8px; text-transform: uppercase; letter-spacing: .04em; font-weight: 700; }
.muted { color: var(--muted); font-style: italic; margin: 4px 0; }
.tbl { width: 100%; border-collapse: collapse; font-size: 13px; background: var(--card-bg); }
.tbl th, .tbl td { border: 1px solid var(--border); padding: 6px 10px; text-align: left; vertical-align: top; }
.tbl thead th { background: var(--brand-green-tint); font-weight: 700; color: var(--brand-green-dark); }
.tbl tfoot td { font-weight: 700; background: #f2f7f4; }
.tbl.summary { margin-bottom: 4px; table-layout: fixed; }
.tbl.summary th, .tbl.summary td { overflow-wrap: break-word; word-break: break-word; }
.tbl td.warn { color: var(--warn-text); font-weight: 700; }
.warn { color: var(--warn-text); font-weight: 700; }
.circ-ok { color: #1f9d5a; font-weight: 700; }
.circ-over { color: var(--warn-text); font-weight: 700; }
.mud-info { margin: 2px 0 8px; }
.mud-legend { font-size: 12px; color: var(--muted); margin: 6px 0 0; }
.mud-tbl th, .mud-tbl td { white-space: nowrap; }
.npv-rating-details summary { cursor: pointer; color: var(--brand-green-dark); font-weight: 700; font-size: 13px; }
.npv-rating-details ul { margin: 8px 0 2px; padding-left: 18px; }
.npv-rating-details li { margin: 3px 0; }
.npv-rating-details-nested { margin: 6px 0 6px 4px; }
.npv-rating-details-nested summary { font-size: 12.5px; color: var(--text); font-weight: 600; }
.section-block { background: var(--card-bg); border: 1px solid var(--border); border-radius: 10px; padding: 16px 20px; margin-bottom: 22px; }
.summary-title { font-size: 17px; font-weight: 800; color: var(--brand-dark); margin: 0 0 12px 0; text-transform: none; letter-spacing: 0; }

.summary-total {
  display: flex; flex-wrap: wrap; align-items: baseline; gap: 10px;
  background: var(--brand-green-tint); border: 1px solid #cfe2d7; border-radius: 8px;
  padding: 10px 16px; margin-bottom: 16px;
}
.summary-total-label { font-weight: 700; color: var(--brand-green-dark); }
.summary-total-value { font-size: 18px; font-weight: 800; color: var(--brand-green-dark); }
.summary-total-count { color: var(--muted); font-size: 12px; margin-left: auto; }

.cluster-block { margin-bottom: 18px; }
.cluster-block:last-child { margin-bottom: 0; }
.cluster-header {
  display: flex; flex-wrap: wrap; align-items: baseline; gap: 14px;
  background: var(--brand-dark); color: #fff; border-radius: 8px 8px 0 0;
  padding: 8px 14px;
}
.cluster-name { font-weight: 800; font-size: 14px; letter-spacing: .02em; }
.cluster-metric { font-size: 12.5px; color: #d7e0da; }
.cluster-metric b { color: #fff; }
.cluster-block .tbl.summary { border-radius: 0 0 8px 8px; overflow: hidden; }
.cluster-block .tbl.summary thead th:first-child { border-top-left-radius: 0; }
.cluster-header.collapsible { cursor: pointer; user-select: none; }
.cluster-toggle-icon { display: inline-block; width: 12px; font-size: 11px; }
.cluster-block.collapsed .tbl.summary { display: none; }

.lag-cell { display: inline-flex; flex-direction: column; line-height: 1.2; }
.lag-cell .lag-value { font-weight: 700; }
.lag-cell .lag-label { font-size: 11px; text-transform: uppercase; letter-spacing: .02em; }
.lag-cell.lag-behind { color: #b3261e; }
.lag-cell.lag-ahead { color: #1e7a3d; }
.lag-cell.lag-none { color: var(--muted); }

.tabs {
  display: flex; gap: 6px; margin-bottom: 16px;
  overflow-x: auto; overflow-y: hidden; -webkit-overflow-scrolling: touch;
  scrollbar-width: thin;
}
.tabs::-webkit-scrollbar { height: 5px; }
.tabs::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
.reference-btn { margin-left: auto; background: var(--brand-green-tint); color: var(--brand-green-dark); }
.reference-btn:hover { background: #dcecE3; }
.table-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch; }

.sr-modal-backdrop {
  position: fixed; inset: 0; z-index: 1100; display: none;
  background: rgba(20, 26, 22, .45); padding: 30px 16px;
  overflow-y: auto;
}
.sr-modal-backdrop.open { display: block; }
.sr-modal {
  background: var(--card-bg); border-radius: 12px; max-width: 780px;
  margin: 0 auto; padding: 24px 28px 28px; box-shadow: 0 20px 60px rgba(0,0,0,.3);
  position: relative;
}
.sr-modal-close {
  position: absolute; top: 14px; right: 16px; border: none; background: none;
  font-size: 22px; line-height: 1; cursor: pointer; color: var(--muted);
}
.sr-modal-close:hover { color: var(--text); }
.tab-btn {
  background: #e7ede9; border: 1px solid var(--border); border-bottom: none;
  padding: 10px 20px; border-radius: 9px 9px 0 0; cursor: pointer;
  font-weight: 700; font-size: 13px; color: var(--muted);
  flex: 0 0 auto; white-space: nowrap;
}
.tab-btn.active { background: var(--card-bg); color: var(--brand-green-dark); box-shadow: 0 -2px 0 var(--brand-green) inset; }
.tab-panel { display: none; }
.tab-panel.active { display: block; }
.card { position: relative; background: var(--card-bg); border: 1px solid var(--border); border-radius: 10px; padding: 18px 22px; margin-bottom: 22px; scroll-margin-top: 16px; border-top: 3px solid var(--brand-green); }
.scroll-top-btn {
  position: absolute; top: 12px; right: 14px; width: 30px; height: 30px; border-radius: 50%;
  background: var(--brand-green-tint); color: var(--brand-green-dark); border: 1px solid #cfe2d7;
  cursor: pointer; font-size: 15px; font-weight: 800; line-height: 1;
  display: inline-flex; align-items: center; justify-content: center; z-index: 2;
}
.scroll-top-btn:hover { background: #dcecE3; }
.tool-note { background: var(--brand-green-tint); border: 1px solid #cfe2d7; border-radius: 8px; padding: 8px 14px; margin-bottom: 14px; font-size: 13px; color: var(--brand-green-dark); }
.card-head { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 10px; border-bottom: 1px solid var(--border); padding-bottom: 10px; padding-right: 38px; margin-bottom: 12px;}
.title-block { display: flex; flex-direction: column; gap: 2px; }
.title-block h2 { margin: 0; }
.title-sub { font-size: 12.5px; color: var(--muted); font-weight: 600; }
.badges { display: flex; flex-wrap: wrap; gap: 8px; }
.badge { background: var(--brand-green-tint); color: var(--brand-green-dark); border-radius: 999px; padding: 4px 12px; font-size: 12px; font-weight: 700; white-space: nowrap;}
.badge.warn { background: var(--warn-bg); color: var(--warn-text); }
.meta-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px 20px; margin-bottom: 6px; }
.meta-grid .v { font-weight: 700; }
.kv-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 6px 24px; }
.kv-grid.two-line { grid-template-columns: repeat(2, 1fr); }
.kv-grid.two-line .wide { grid-column: 1 / -1; }
.kv-grid.contacts-grid { grid-template-columns: repeat(2, 1fr); }
.kv { display: flex; flex-direction: column; padding: 4px 0; border-bottom: 1px dashed var(--border); }
.kv .k { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: .02em; }
.kv .v { font-size: 14px; }
.kv .v.strong { font-weight: 700; }
.contact-sub { font-size: 12px; color: var(--muted); }
.two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; align-items: start; }
.src { margin-top: 10px; font-size: 11px; color: var(--muted); text-align: right; }
.errors { background: #fff3f2; border: 1px solid #f3c8c3; border-radius: 10px; padding: 14px 18px; margin-bottom: 22px; color: #8a2f26; }
.errors li { margin-bottom: 4px; }
footer { text-align: center; color: var(--muted); font-size: 12px; padding: 20px; }

.schedule-block { background: var(--brand-green-tint); border: 1px solid #cfe2d7; border-radius: 8px; padding: 10px 16px; }
.schedule-block p { margin: 6px 0; }

.depth-chart { margin-top: 12px; text-align: center; }
.depth-chart img { max-width: 100%; height: auto; border: 1px solid var(--border); border-radius: 6px; background: #fff; }
.depth-chart svg { max-width: 100%; height: auto; border: 1px solid var(--border); border-radius: 6px; background: #fff; }

.npv-summary { display: flex; flex-wrap: wrap; gap: 20px; background: var(--warn-bg); border: 1px solid #f0cdbe; border-radius: 8px; padding: 10px 16px; margin: 14px 0 8px; }
.npv-summary-item { display: flex; flex-direction: column; }
.npv-summary-item .k { font-size: 11px; color: var(--warn-text); text-transform: uppercase; letter-spacing: .02em; }
.npv-summary-item .v { color: var(--warn-text); }

details.collapsible { border: 1px solid var(--border); border-radius: 8px; margin: 10px 0; background: #fbfcfb; }
details.collapsible > summary {
  cursor: pointer; padding: 9px 14px; font-weight: 700; font-size: 13px;
  color: var(--brand-green-dark); list-style: none; display: flex; align-items: center; gap: 8px;
}
details.collapsible > summary::-webkit-details-marker { display: none; }
details.collapsible > summary::before {
  content: "\\25B8"; display: inline-block; transition: transform .15s ease; color: var(--brand-green);
}
details.collapsible[open] > summary::before { transform: rotate(90deg); }
details.collapsible > .collapsible-body { padding: 4px 14px 14px; }

.plan-msp-trigger {
  color: var(--brand-green-dark); font-weight: 700; cursor: pointer;
  text-decoration: underline dotted; text-underline-offset: 2px;
}
.plan-msp-trigger:hover, .plan-msp-trigger:focus { color: var(--accent-warm); outline: none; }

.comment-trigger {
  cursor: pointer; font: inherit; font-size: 12px; font-weight: 700;
  color: var(--brand-green-dark); background: var(--brand-green-tint);
  border: 1px solid #cfe2d7; border-radius: 6px; padding: 3px 10px;
  white-space: nowrap;
}
.comment-trigger:hover, .comment-trigger:focus { background: #dcecE3; outline: none; }

.stage-warn {
  display: inline-flex; align-items: center; justify-content: center;
  width: 17px; height: 17px; margin-left: 6px; border-radius: 50%;
  background: var(--accent-warm); color: #fff; font-size: 12px; font-weight: 800;
  line-height: 1; cursor: help; vertical-align: middle;
}

.sr-popover {
  position: fixed; z-index: 1000; display: none;
  background: var(--card-bg); border: 1px solid var(--border); border-radius: 10px;
  box-shadow: 0 10px 30px rgba(0,0,0,.2); padding: 12px 16px; min-width: 230px;
  font-size: 13px;
}
.sr-popover.open { display: block; }
.sr-popover.wide { max-width: 480px; max-height: 60vh; overflow-y: auto; }
.sr-popover .pop-comment-text { white-space: normal; line-height: 1.5; color: var(--text); }
.sr-popover .pop-head {
  display: flex; align-items: center; justify-content: space-between; gap: 12px;
  margin: -2px 0 8px; font-weight: 700; color: var(--brand-green-dark);
  border-bottom: 1px solid var(--border); padding-bottom: 6px;
}
.sr-popover .pop-close {
  cursor: pointer; border: none; background: none; font-size: 17px; line-height: 1;
  color: var(--muted); padding: 0 2px;
}
.sr-popover .pop-close:hover { color: var(--text); }
.sr-popover .pop-row { display: flex; justify-content: space-between; gap: 18px; padding: 3px 0; }
.sr-popover .pop-row .pop-k { color: var(--muted); }
.sr-popover .pop-row .pop-v { font-weight: 700; white-space: nowrap; }

@media (max-width: 900px) {
  .meta-grid, .two-col, .kv-grid { grid-template-columns: 1fr; }
  .top-inner { flex-direction: column; align-items: flex-start; }
}
@media print {
  body { background: #fff; }
  .card { break-inside: avoid; }
  details.collapsible { break-inside: avoid; }
}
"""


def build_html(results, errors, folder_path, generated_at=None, cluster_overrides=None):
    generated_at = generated_at or datetime.datetime.now()

    def sort_key(item):
        d = item
        h = d["header"]
        return (field_group_key(h.get("field")), numeric_or_text_key(h.get("pad")), numeric_or_text_key(h.get("well_no")))

    ordered = sorted(results, key=sort_key)
    items = list(enumerate(ordered, start=1))

    report_date = None
    for d in results:
        rd = d["header"].get("report_date")
        if rd:
            report_date = rd
            break

    errors_html = ""
    if errors:
        li = "".join(f"<li><b>{esc(fn)}</b> — {esc(msg)}</li>" for fn, msg in errors)
        errors_html = f"""
<div class="errors">
  <b>Не удалось разобрать {len(errors)} файл(ов):</b>
  <ul>{li}</ul>
</div>"""

    cards_html = "".join(render_well_card(idx, d) for idx, d in items)
    summary_html = render_summary_table(items, cluster_overrides)
    npv_rating_html = render_npv_rating_tab(items, cluster_overrides)
    bits_html = render_bits_table(items, cluster_overrides) if items else '<p class="muted">Нет ни одного успешно разобранного рапорта.</p>'
    vzd_html = render_vzd_table(items, cluster_overrides) if items else '<p class="muted">Нет ни одного успешно разобранного рапорта.</p>'
    osc_html = render_osc_table(items, cluster_overrides) if items else '<p class="muted">Нет ни одного успешно разобранного рапорта.</p>'
    sbt_html = render_sbt_table(items, cluster_overrides) if items else '<p class="muted">Нет ни одного успешно разобранного рапорта.</p>'
    reference_modal_html = render_reference_modal()

    logo_uri = _logo_data_uri()
    logo_html = f'<div class="logo-plate"><img src="{logo_uri}" alt="ООО «ИНК»"></div>' if logo_uri else ""

    html_doc = f"""<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Сводный суточный рапорт</title>
<style>{CSS}</style>
</head>
<body>
<header class="top">
  <div class="top-inner">
    <div class="top-text">
      <h1>Сводный суточный рапорт по бурению</h1>
      <div class="sub">
        Дата рапортов: {fmt_date(report_date)} &nbsp;|&nbsp;
        Скважин в отчёте: {len(results)} &nbsp;|&nbsp;
        Источник: {esc(folder_path)} &nbsp;|&nbsp;
        Сформировано: {generated_at.strftime('%d.%m.%Y %H:%M')} &nbsp;|&nbsp;
        Версия программы: {REPORT_VERSION}
      </div>
    </div>
    {logo_html}
  </div>
</header>
<main>
  {errors_html}

  <div class="tabs">
    <button type="button" class="tab-btn active" data-tab="wells" onclick="srShowTab('wells')">Скважины</button>
    <button type="button" class="tab-btn" data-tab="npvrating" onclick="srShowTab('npvrating')">Рейтинг НПВ</button>
    <button type="button" class="tab-btn" data-tab="bits" onclick="srShowTab('bits')">Долота</button>
    <button type="button" class="tab-btn" data-tab="vzd" onclick="srShowTab('vzd')">ВЗД</button>
    <button type="button" class="tab-btn" data-tab="osc" onclick="srShowTab('osc')">Осциллятор</button>
    <button type="button" class="tab-btn" data-tab="tool" onclick="srShowTab('tool')">Инструмент</button>
    <button type="button" class="tab-btn reference-btn" id="sr-reference-btn" hidden onclick="srShowReferenceModal()">Справочная</button>
  </div>

  <div class="tab-panel active" data-tab="wells">
    <div class="section-block">
      <p class="summary-title">Общая сводка по всем скважинам</p>
      {summary_html}
    </div>
    {cards_html}
  </div>

  <div class="tab-panel" data-tab="npvrating">
    <div class="section-block">
      <p class="summary-title">Рейтинг НПВ — компании-подрядчики по объёму НПВ за отчётные сутки</p>
      {npv_rating_html}
    </div>
  </div>

  <div class="tab-panel" data-tab="bits">
    <div class="section-block">
      <p class="summary-title">Долота — сводка по всем скважинам</p>
      {bits_html}
    </div>
  </div>

  <div class="tab-panel" data-tab="vzd">
    <div class="section-block">
      <p class="summary-title">ВЗД — сводка по всем скважинам</p>
      {vzd_html}
    </div>
  </div>

  <div class="tab-panel" data-tab="osc">
    <div class="section-block">
      <p class="summary-title">Осциллятор — сводка по всем скважинам</p>
      {osc_html}
    </div>
  </div>

  <div class="tab-panel" data-tab="tool">
    <div class="section-block">
      <p class="summary-title">Инструмент — сводка по всем скважинам</p>
      <div class="tool-note">Вкладка «Инструмент» показывает наработку СБТ (стальных бурильных труб) и относится только к секции «Хвостовик» — по остальным секциям строительства скважины данные здесь не приводятся.</div>
      {sbt_html}
    </div>
  </div>
</main>
<footer>Отчёт сформирован автоматически из файлов СР (суточных рапортов по бурению). Версия программы {REPORT_VERSION}.</footer>
<div id="sr-popover" class="sr-popover" role="dialog" aria-live="polite"></div>
<div id="sr-reference-modal" class="sr-modal-backdrop" onclick="if (event.target === this) srHideReferenceModal();">
  <div class="sr-modal" role="dialog" aria-modal="true" aria-label="Справочная">
    <button type="button" class="sr-modal-close" onclick="srHideReferenceModal()" aria-label="Закрыть">&times;</button>
    {reference_modal_html}
  </div>
</div>
<script>
function srShowTab(name) {{
  document.querySelectorAll('.tab-panel').forEach(function (p) {{
    p.classList.toggle('active', p.dataset.tab === name);
  }});
  document.querySelectorAll('.tab-btn').forEach(function (b) {{
    if (b.id !== 'sr-reference-btn') b.classList.toggle('active', b.dataset.tab === name);
  }});
  var refBtn = document.getElementById('sr-reference-btn');
  if (refBtn) refBtn.hidden = (name !== 'vzd' && name !== 'osc');
}}

function srToggleClusterBlock(headerEl) {{
  var block = headerEl.closest('.cluster-block');
  if (!block) return;
  block.classList.toggle('collapsed');
  var icon = headerEl.querySelector('.cluster-toggle-icon');
  if (icon) icon.innerHTML = block.classList.contains('collapsed') ? '&#9654;' : '&#9660;';
}}

function srShowReferenceModal() {{
  var modal = document.getElementById('sr-reference-modal');
  if (modal) modal.classList.add('open');
}}

function srHideReferenceModal() {{
  var modal = document.getElementById('sr-reference-modal');
  if (modal) modal.classList.remove('open');
}}

document.addEventListener('keydown', function (e) {{
  if (e.key === 'Escape') srHideReferenceModal();
}});

function srShowPlanPopover(evt, el) {{
  evt.stopPropagation();
  var data;
  try {{
    data = JSON.parse(el.getAttribute('data-plan'));
  }} catch (e) {{
    return;
  }}
  var pop = document.getElementById('sr-popover');
  if (!pop) return;
  pop.innerHTML =
    '<div class="pop-head"><span>План МСП (мех. бурение)</span>' +
    '<button type="button" class="pop-close" onclick="srHidePlanPopover()" aria-label="Закрыть">&times;</button></div>' +
    '<div class="pop-row"><span class="pop-k">План МСП, м/ч</span><span class="pop-v">' + data.msp + '</span></div>' +
    '<div class="pop-row"><span class="pop-k">План «от», м</span><span class="pop-v">' + data.from + '</span></div>' +
    '<div class="pop-row"><span class="pop-k">План «до», м</span><span class="pop-v">' + data.to + '</span></div>' +
    '<div class="pop-row"><span class="pop-k">План интервал, м</span><span class="pop-v">' + data.interval + '</span></div>' +
    '<div class="pop-row"><span class="pop-k">Факт «от», м</span><span class="pop-v">' + data.fact_from + '</span></div>';
  pop.style.display = 'block';
  pop.classList.add('open');
  var rect = el.getBoundingClientRect();
  var popW = pop.offsetWidth, popH = pop.offsetHeight;
  var left = rect.left, top = rect.bottom + 8;
  if (left + popW > window.innerWidth - 10) left = window.innerWidth - popW - 10;
  if (top + popH > window.innerHeight - 10) top = rect.top - popH - 8;
  pop.style.left = Math.max(10, left) + 'px';
  pop.style.top = Math.max(10, top) + 'px';
}}

function srHidePlanPopover() {{
  var pop = document.getElementById('sr-popover');
  if (pop) {{
    pop.classList.remove('open');
    pop.classList.remove('wide');
    pop.style.display = 'none';
  }}
}}

function srShowCommentPopover(evt, el) {{
  evt.stopPropagation();
  var data;
  try {{
    data = JSON.parse(el.getAttribute('data-comment'));
  }} catch (e) {{
    return;
  }}
  var pop = document.getElementById('sr-popover');
  if (!pop) return;
  var text = String(data.comment || '').split('\\n').map(function (line) {{
    var d = document.createElement('div');
    d.textContent = line;
    return d.innerHTML || '&nbsp;';
  }}).join('<br>');
  pop.innerHTML =
    '<div class="pop-head"><span>Комментарий</span>' +
    '<button type="button" class="pop-close" onclick="srHidePlanPopover()" aria-label="Закрыть">&times;</button></div>' +
    '<div class="pop-comment-text">' + text + '</div>';
  pop.classList.add('wide');
  pop.style.display = 'block';
  pop.classList.add('open');
  var rect = el.getBoundingClientRect();
  var popW = pop.offsetWidth, popH = pop.offsetHeight;
  var left = rect.left, top = rect.bottom + 8;
  if (left + popW > window.innerWidth - 10) left = window.innerWidth - popW - 10;
  if (top + popH > window.innerHeight - 10) top = Math.max(10, rect.top - popH - 8);
  pop.style.left = Math.max(10, left) + 'px';
  pop.style.top = Math.max(10, top) + 'px';
}}

document.addEventListener('click', function (e) {{
  var pop = document.getElementById('sr-popover');
  if (pop && pop.classList.contains('open') && !pop.contains(e.target) && !e.target.classList.contains('sr-trigger')) {{
    srHidePlanPopover();
    pop.classList.remove('wide');
  }}
}});
document.addEventListener('keydown', function (e) {{
  if (e.key === 'Escape') srHidePlanPopover();
}});

// Широкие таблицы на узких экранах (телефон) иначе раздвигают всю страницу
// в ширину (это особенно заметно в мобильном Safari на iPhone — вся страница
// уменьшается по масштабу, чтобы вместить самую широкую таблицу, и вкладки
// "Осциллятор"/"Инструмент" вверху уезжают за край экрана без возможности
// пролистать к ним) — оборачиваем каждую таблицу в свой прокручиваемый
// по горизонтали контейнер, не трогая остальную вёрстку страницы.
document.querySelectorAll('table.tbl').forEach(function (t) {{
  if (t.parentElement && t.parentElement.classList.contains('table-scroll')) return;
  var wrap = document.createElement('div');
  wrap.className = 'table-scroll';
  t.parentNode.insertBefore(wrap, t);
  wrap.appendChild(t);
}});
</script>
</body>
</html>"""
    return html_doc
