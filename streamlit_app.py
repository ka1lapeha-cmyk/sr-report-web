# -*- coding: utf-8 -*-
"""
Веб-версия генератора сводного СР.

Использует ровно ту же логику разбора файлов и сборки отчёта, что и
десктопная программа (main.py + sr_parser.py + report_html.py), но без
tkinter — работает как обычная страница в браузере на Streamlit, поэтому
доступна в том числе с телефона, без установки чего-либо на компьютер.

Важно: этот файл НЕ импортирует main.py (там tkinter, который на сервере
не нужен и может быть не установлен) — только sr_parser.py и
report_html.py, у которых из сторонних пакетов нужен только openpyxl.

Запуск у себя локально для проверки (на компьютере с Python):
    pip install streamlit openpyxl
    streamlit run streamlit_app.py

Как выложить в интернет (чтобы работало с телефона без компьютера) —
см. README_WEB.txt рядом с этим файлом.
"""

import datetime
import traceback

import streamlit as st

from sr_parser import ParseError
from web_report import generate_report

WEB_APP_VERSION = "1.0"


def _run_ui():
    st.set_page_config(page_title="СР — сводный отчёт", page_icon="\U0001F4CB", layout="centered")

    st.title("Генератор сводного отчёта СР")
    st.caption(f"Веб-версия {WEB_APP_VERSION} · та же логика, что и в программе для компьютера")

    st.markdown(
        "Загрузите файлы суточных рапортов (.xlsx) за нужную дату — по одному или "
        "сразу несколько — либо один .zip-архив с ними внутри, и нажмите "
        "«Сформировать отчёт»."
    )

    uploaded = st.file_uploader(
        "Файлы СР",
        type=["xlsx", "zip"],
        accept_multiple_files=True,
        help="Можно выбрать сразу несколько .xlsx файлов, либо один .zip-архив с ними внутри.",
    )

    build_chart = st.checkbox("Строить график «глубина-день»", value=True)

    generate = st.button("Сформировать отчёт", type="primary", disabled=not uploaded)

    if not generate:
        return
    if not uploaded:
        st.warning("Сначала выберите файлы.")
        return

    saved_items = [(f.name, f.getvalue()) for f in uploaded]

    progress_bar = st.progress(0.0)
    status = st.empty()

    def on_progress(i, total, name):
        progress_bar.progress(i / max(total, 1))
        status.text(f"Обработка {i}/{total}: {name}")

    try:
        with st.spinner("Разбираю файлы..."):
            results, errors, html_doc = generate_report(saved_items, build_chart=build_chart, on_progress=on_progress)
    except ParseError as e:
        progress_bar.empty()
        status.empty()
        st.error(str(e))
        return
    except Exception:
        progress_bar.empty()
        status.empty()
        st.error("Непредвиденная ошибка при обработке файлов:")
        st.code(traceback.format_exc())
        return

    progress_bar.empty()
    status.empty()

    if not results:
        st.error("Ни один файл не удалось разобрать. Проверьте, что это файлы СР нужного формата.")
        if errors:
            with st.expander(f"Ошибки при разборе ({len(errors)})"):
                for name, msg in errors:
                    st.write(f"**{name}**: {msg}")
        return

    st.success(f"Готово: {len(results)} скважин обработано, ошибок: {len(errors)}.")
    if errors:
        with st.expander(f"Ошибки при разборе ({len(errors)})"):
            for name, msg in errors:
                st.write(f"**{name}**: {msg}")

    report_date = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    st.download_button(
        "Скачать отчёт (HTML)",
        data=html_doc.encode("utf-8"),
        file_name=f"SR_Общий_{report_date}.html",
        mime="text/html",
        type="primary",
    )
    st.caption(
        "Скачанный файл откроется в браузере телефона как обычная веб-страница — "
        "все вкладки и раскрывающиеся карточки работают в нём без интернета, "
        "как и в отчёте из десктопной программы."
    )


# Streamlit исполняет весь файл сверху вниз при каждом "streamlit run" /
# перерисовке страницы — отдельного __main__ не нужно.
_run_ui()
