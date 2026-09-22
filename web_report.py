# -*- coding: utf-8 -*-
"""
Логика веб-версии генератора отчёта, не зависящая от Streamlit.

Вынесена в отдельный модуль от streamlit_app.py специально, чтобы её можно
было проверить обычным Python-скриптом/тестом на реальных файлах СР, без
установленного Streamlit и без поднятия веб-страницы — сам streamlit_app.py
только собирает вокруг этой функции интерфейс страницы.
"""

import os
import shutil
import tempfile

from sr_parser import parse_source, ParseError
import report_html as rh

WEB_SOURCE_LABEL = "Веб-версия (файлы загружены вручную)"


def generate_report(saved_items, build_chart=True, on_progress=None):
    """saved_items: список пар (имя_файла, содержимое_в_байтах) — то, что
    пришло от st.file_uploader (f.name, f.getvalue()) или от любого другого
    источника с тем же интерфейсом (например, тестового скрипта).

    Принимает либо несколько .xlsx, либо ровно один .zip-архив с ними
    внутри (смешивать нельзя — тот же принцип, что и в десктопной
    программе, где источник — это либо папка, либо один архив).

    Возвращает (results, errors, html_doc). html_doc — None, если ни один
    файл не разобрался (тогда смотреть errors). Бросает ParseError при
    некорректном/пустом/смешанном наборе файлов."""
    zip_items = [it for it in saved_items if it[0].lower().endswith(".zip")]
    xlsx_items = [it for it in saved_items if it[0].lower().endswith(".xlsx")]

    if zip_items and xlsx_items:
        raise ParseError(
            "Загружены одновременно и .zip, и отдельные .xlsx файлы. "
            "Загрузите что-то одно: либо один .zip-архив, либо несколько .xlsx файлов."
        )
    if len(zip_items) > 1:
        raise ParseError("Загрузите только один .zip-архив за раз.")

    tmpdir = tempfile.mkdtemp(prefix="sr_web_")
    try:
        if zip_items:
            name, data = zip_items[0]
            input_path = os.path.join(tmpdir, name)
            with open(input_path, "wb") as f:
                f.write(data)
        else:
            if not xlsx_items:
                raise ParseError("Не выбрано ни одного файла .xlsx или .zip.")
            for name, data in xlsx_items:
                dest = os.path.join(tmpdir, os.path.basename(name))
                with open(dest, "wb") as f:
                    f.write(data)
            input_path = tmpdir

        results, errors = parse_source(input_path, on_progress=on_progress, build_chart=build_chart)
        html_doc = None
        if results:
            html_doc = rh.build_html(results, errors, WEB_SOURCE_LABEL)
        return results, errors, html_doc
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
