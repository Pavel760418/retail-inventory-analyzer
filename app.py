# -*- coding: utf-8 -*-
"""Streamlit UI — Анализ итогов инвентаризации торгового зала.

Оборачивает существующую бизнес-логику (parser/metrics/export) без её переписывания.
"""
from __future__ import annotations

import datetime
import io
import sys
import tempfile
import traceback
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.settings import PERIOD_DAYS, load_settings
from src import __version__
from src.analysis_service import ensure_excel_bytes, run_ui_analysis
from src.master_hierarchy import get_master_hierarchy
from src.ui_helpers import (
    detail_table,
    discrepancy_structure_df,
    executive_insights,
    kpi_cards,
    status_risk_table,
    store_risk_chart_df,
    top_shortage_chart_df,
)
from src.validators import (
    EXPECTED_FORMAT_HINT,
    validate_capitalization_bytes,
    validate_uploaded_inventory_bytes,
)

st.set_page_config(
    page_title="Анализ итогов инвентаризации торгового зала",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DISPLAY_ROW_LIMIT = 5_000


def _save_bytes(data: bytes, suffix: str) -> Path:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(data)
    tmp.flush()
    tmp.close()
    return Path(tmp.name)


def _df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def _df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def _plot_bar_h(df: pd.DataFrame, x: str, y: str, title: str):
    try:
        import plotly.express as px
    except ImportError:
        st.warning("Plotly не установлен — график недоступен.")
        return
    if df is None or df.empty:
        st.info("Недостаточно данных для построения графика.")
        return
    fig = px.bar(df, x=x, y=y, orientation="h", title=title, text_auto=".2s")
    fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=420, margin=dict(l=10, r=10, t=40, b=10))
    st.plotly_chart(fig, use_container_width=True)


def _plot_pie(df: pd.DataFrame, names: str, values: str, title: str):
    try:
        import plotly.express as px
    except ImportError:
        st.warning("Plotly не установлен — график недоступен.")
        return
    if df is None or df.empty or float(df[values].sum()) <= 0:
        st.info("Недостаточно данных для построения графика.")
        return
    fig = px.pie(df, names=names, values=values, title=title, hole=0.35)
    fig.update_layout(height=420, margin=dict(l=10, r=10, t=40, b=10))
    st.plotly_chart(fig, use_container_width=True)


def render_sidebar():
    settings = load_settings()
    st.sidebar.title("Параметры")
    st.sidebar.caption(f"Розница v{__version__}")

    try:
        master = get_master_hierarchy()
        st.sidebar.success(
            f"Эталон иерархии: {master.source_file}\n"
            f"{len(master.items):,} SKU"
        )
    except FileNotFoundError as exc:
        st.sidebar.error(str(exc))

    st.sidebar.subheader("Загрузка данных")
    inv = st.sidebar.file_uploader(
        "Инвентаризация (.xlsx)",
        type=["xlsx", "xls"],
        key="inv_file",
    )
    cap = st.sidebar.file_uploader(
        "Оприходование излишков (.xlsx)",
        type=["xlsx", "xls"],
        key="cap_file",
    )

    period_days = st.sidebar.number_input(
        "Период, дней",
        min_value=1,
        max_value=365,
        value=int(settings.period_days or PERIOD_DAYS),
    )
    # Default OFF in UI — cross-store is O(n²) and often causes Cloud timeouts
    enable_cross = st.sidebar.checkbox(
        "Аналитика между магазинами",
        value=False,
        help="Медленно на больших файлах. Включайте только при необходимости сравнения магазинов.",
    )
    if enable_cross:
        st.sidebar.warning("Межмагазинная аналитика сильно замедляет расчёт.")
    use_custom_end = st.sidebar.checkbox("Задать конечную дату", value=False)
    end_date = None
    if use_custom_end:
        end_date = st.sidebar.date_input("Конечная дата", value=datetime.date.today())

    show_tech = st.sidebar.checkbox("Показывать технические детали", value=False)
    run = st.sidebar.button("Запустить анализ", type="primary", use_container_width=True)

    st.sidebar.markdown("---")
    st.sidebar.caption(
        "Ревизор = документ ровно в 08:00:00. Проверка цен отключена. "
        "Файлы не сохраняются на сервере после сессии."
    )
    return inv, cap, period_days, enable_cross, end_date, show_tech, run


def render_instruction_tab():
    st.markdown(
        """
### Как подготовить входной файл

1. Выгрузите из 1С **исходную** инвентаризацию сети/торгового зала (не готовый «Анализ_…»).
2. Формат: **Excel .xlsx** (лист с иерархией Документ.Подразделение / номенклатура).
3. Обязательные колонки-маркеры в шапке:
   - Документ.Подразделение
   - Количество книжное / Количество фактическое
   - Сумма излишек / Сумма недостача
4. Второй файл — оприходование излишков при закрытии смены с колонками
   «Закрытие смены количество» и «Закрытие смены сумма».

### Статусы

- **Ревизор** — время документа ровно 08:00:00.
- **Оператор** — любое другое распознанное время.
- **Аномалии книжных сумм** — пустые нормативная/фактическая суммы (см. лист «Аномалии»).
- **Чистые недостачи** — не объяснены однородным пересортом.

### Excel-отчёт

После анализа откройте вкладку «Excel-отчёт». Отчёт собирается **один раз** при первом открытии вкладки
(та же `build_report`, что и CLI). Метрики на главной появляются сразу после расчёта, не дожидаясь Excel.

### Поддержка

Канал поддержки задаётся организацией. При ошибке структуры файла сверьте выгрузку с инструкцией выше.
        """
    )
    st.info(EXPECTED_FORMAT_HINT)


def render_home(result):
    st.subheader("Ключевые показатели")
    cards = kpi_cards(result)
    cols = st.columns(3)
    for i, card in enumerate(cards):
        with cols[i % 3]:
            st.metric(card["label"], card["value"], delta=card["delta"], help=card["help"])

    st.subheader("Вывод для руководителя")
    for line in executive_insights(result):
        st.markdown(f"- {line}")

    st.subheader("Статусы и риски")
    st.dataframe(status_risk_table(result), use_container_width=True, hide_index=True)

    st.subheader("Визуализации")
    c1, c2 = st.columns(2)
    with c1:
        _plot_bar_h(
            top_shortage_chart_df(result),
            x="Недостача, ₽",
            y="Позиция",
            title="Top-10 позиций по недостаче",
        )
    with c2:
        _plot_pie(
            discrepancy_structure_df(result),
            names="Тип",
            values="Сумма, ₽",
            title="Структура расхождений",
        )
    _plot_bar_h(
        store_risk_chart_df(result),
        x="Чистые недостачи, ₽",
        y="Магазин",
        title="Рейтинг магазинов по чистым недостачам",
    )


def render_detail(result):
    det = detail_table(result)
    st.caption(f"Строк в результате: {len(det):,}. На экране не более {DISPLAY_ROW_LIMIT:,}.")

    stores = sorted(det["магазин"].dropna().astype(str).unique()) if "магазин" in det.columns else []
    authors = sorted(det["автор_дока"].dropna().astype(str).unique()) if "автор_дока" in det.columns else []
    f1, f2, f3 = st.columns(3)
    store_f = f1.multiselect("Магазин", stores)
    author_f = f2.multiselect("Автор документа", authors)
    q = f3.text_input("Поиск по наименованию / SKU")

    view = det
    if store_f and "магазин" in view.columns:
        view = view[view["магазин"].isin(store_f)]
    if author_f and "автор_дока" in view.columns:
        view = view[view["автор_дока"].isin(author_f)]
    if q and "наименование" in view.columns:
        mask = view["наименование"].astype(str).str.contains(q, case=False, na=False)
        view = view[mask]

    if len(view) > DISPLAY_ROW_LIMIT:
        st.warning(f"Показаны первые {DISPLAY_ROW_LIMIT:,} строк из {len(view):,}. Скачайте CSV/Excel для полного набора.")
        show = view.head(DISPLAY_ROW_LIMIT)
    else:
        show = view
    st.dataframe(show, use_container_width=True, hide_index=True)

    d1, d2 = st.columns(2)
    d1.download_button(
        "Скачать фильтр CSV",
        data=_df_to_csv_bytes(view),
        file_name="detail_filtered.csv",
        mime="text/csv",
    )
    d2.download_button(
        "Скачать фильтр Excel",
        data=_df_to_xlsx_bytes(view),
        file_name="detail_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_quality(result):
    q = result.quality
    if q.passed and not q.issues:
        st.success("✅ Проверки пройдены")
    elif q.passed:
        st.warning("⚠️ Анализ выполнен, есть замечания по качеству данных")
    else:
        st.error("❌ Обнаружены проблемы во входных данных")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Строк после периода", f"{q.after_period_rows:,}")
    m2.metric("Исключено периодом", f"{q.excluded_by_period:,}")
    m3.metric("Вне справочника", f"{q.hierarchy_unmatched_rows:,}")
    m4.metric("Аномалий / крит.", f"{q.anomaly_total} / {q.anomaly_critical}")

    if q.issues:
        st.dataframe(pd.DataFrame(q.issues), use_container_width=True, hide_index=True)
        st.download_button(
            "Скачать протокол замечаний CSV",
            data=_df_to_csv_bytes(pd.DataFrame(q.issues)),
            file_name="quality_issues.csv",
            mime="text/csv",
        )
    else:
        st.info("Критичных замечаний по качеству нет.")


def render_excel_tab(result, show_tech: bool = False):
    st.markdown(
        """
Итоговый Excel формируется **той же** функцией `build_report`, что и CLI.
Состав: Сводка, Выводы, Рейтинг, Аномалии, Ревизоры vs Операторы, Цепочки пересортов,
Оприходование, топы, перекрытие, мероприятия, детализация магазинов и др.

Метрики на вкладке «Главная» уже посчитаны. Excel собирается **отдельно по кнопке**,
чтобы не блокировать экран повторным полным проходом.
        """
    )
    ready = bool(result.excel_bytes) and not getattr(result, "excel_pending", False)

    if not ready:
        st.info("Excel ещё не сформирован. Нажмите кнопку — один проход сборки отчёта.")
        if st.button("Сформировать Excel-отчёт", type="primary", key="build_excel_btn"):
            # Native Streamlit widgets only — custom HTML + rerun caused removeChild React errors
            with st.status("🤖 AI-агент формирует Excel-отчёт…", expanded=True) as status:
                progress = st.progress(0, text="Подготовка Excel…")

                def _cb(v: float, text: str) -> None:
                    status.write(text)
                    progress.progress(min(max(int(v * 100), 0), 100), text=text)

                try:
                    result = ensure_excel_bytes(result, progress_cb=_cb)
                    st.session_state["analysis_result"] = result
                    progress.progress(100, text="Excel готов")
                    status.update(label="🤖 Excel-отчёт готов", state="complete")
                    ready = True
                except Exception as exc:
                    status.update(label="Ошибка формирования Excel", state="error")
                    st.error(f"Не удалось сформировать Excel: {exc}")
                    if show_tech:
                        st.code(traceback.format_exc())
                    return
        if not ready:
            return

    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
    fname = f"Анализ_инвентаризации_{ts}.xlsx"
    st.success("Excel-отчёт готов к скачиванию.")
    st.write(
        f"Размер: {len(result.excel_bytes) / 1024:,.0f} КБ · "
        f"Листов: {result.excel_sheets} · "
        f"Время формирования: {result.excel_seconds:.1f} с"
    )
    st.caption(f"Отпечаток набора: `{result.source_fingerprint}` · период {result.period_str}")
    st.download_button(
        "Скачать Excel-отчёт",
        data=result.excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


def render_results(result, show_tech: bool = False, caption: str | None = None) -> None:
    tabs = st.tabs(["Главная", "Детализация", "Контроль качества", "Excel-отчёт", "Инструкция"])
    with tabs[0]:
        render_home(result)
    with tabs[1]:
        render_detail(result)
    with tabs[2]:
        render_quality(result)
    with tabs[3]:
        render_excel_tab(result, show_tech=show_tech)
    with tabs[4]:
        render_instruction_tab()
    if caption:
        st.caption(caption)


def main() -> None:
    st.title("Анализ итогов инвентаризации торгового зала")
    st.markdown(
        "Загрузите файл инвентаризации, получите контроль качества данных, "
        "ключевые показатели расхождений и готовый Excel-отчёт."
    )

    inv, cap, period_days, enable_cross, end_date, show_tech, run = render_sidebar()

    # Prefer cached session result if uploaders were cleared after a long run (Streamlit Cloud)
    result_cached = st.session_state.get("analysis_result")

    if inv is None or cap is None:
        if result_cached is not None:
            render_results(
                result_cached,
                show_tech=show_tech,
                caption="Файлы в загрузчике сброшены — метрики из текущей сессии. "
                        "Чтобы пересчитать, загрузите файлы снова.",
            )
            return
        st.info(
            "Загрузите в боковой панели **два** Excel-файла: инвентаризацию и оприходование излишков. "
            "Затем нажмите «Запустить анализ»."
        )
        with st.expander("Краткая инструкция", expanded=True):
            render_instruction_tab()
        return

    # Cache upload bytes in session so reruns after analysis don't lose file content
    inv_bytes = inv.getvalue()
    cap_bytes = cap.getvalue()
    upload_key = f"{inv.name}|{len(inv_bytes)}|{cap.name}|{len(cap_bytes)}|{period_days}|{enable_cross}|{end_date}"
    prev_key = st.session_state.get("upload_key")
    if prev_key != upload_key:
        # Inputs changed — drop stale result so user must re-run intentionally
        if "analysis_result" in st.session_state and prev_key is not None:
            st.session_state.pop("analysis_result", None)
        st.session_state["upload_key"] = upload_key
        st.session_state["inv_bytes"] = inv_bytes
        st.session_state["cap_bytes"] = cap_bytes
        st.session_state["inv_name"] = inv.name
        st.session_state["cap_name"] = cap.name

    inv_val = validate_uploaded_inventory_bytes(inv_bytes, inv.name)
    cap_val = validate_capitalization_bytes(cap_bytes, cap.name)
    if not inv_val.ok:
        st.error(inv_val.user_message)
        if show_tech:
            st.code("\n".join(inv_val.errors))
        return
    if not cap_val.ok:
        st.error(cap_val.user_message)
        if show_tech:
            st.code("\n".join(cap_val.errors))
        return
    for w in inv_val.warnings:
        st.warning(w)

    if run:
        st.session_state["pending_run"] = True

    if st.session_state.get("pending_run"):
        st.session_state["pending_run"] = False
        inv_path = _save_bytes(st.session_state["inv_bytes"], Path(st.session_state["inv_name"]).suffix or ".xlsx")
        cap_path = _save_bytes(st.session_state["cap_bytes"], Path(st.session_state["cap_name"]).suffix or ".xlsx")
        try:
            # Native st.status only (no custom HTML / st.rerun) — avoids React removeChild crashes
            with st.status("🤖 AI-агент выполняет работу и готовит анализ…", expanded=True) as status:
                st.caption(
                    "Парсинг, сверка иерархии, метрики и оприходование. "
                    "Excel — отдельно во вкладке «Excel-отчёт»."
                )
                progress = st.progress(0, text="AI-агент приступает к работе…")

                def _cb(v: float, text: str) -> None:
                    status.write(text)
                    progress.progress(min(max(int(v * 100), 0), 100), text=text)

                result = run_ui_analysis(
                    str(inv_path),
                    str(cap_path),
                    period_days=int(period_days),
                    end_date=end_date if isinstance(end_date, datetime.date) else None,
                    enable_cross_store=bool(enable_cross),
                    source_names=(st.session_state["inv_name"], st.session_state["cap_name"]),
                    inv_bytes=st.session_state["inv_bytes"],
                    cap_bytes=st.session_state["cap_bytes"],
                    progress_cb=_cb,
                )
                progress.progress(100, text="Анализ готов")
                st.session_state["analysis_result"] = result
                st.session_state["analysis_fp"] = result.source_fingerprint
                status.update(label="🤖 Анализ готов", state="complete")
            st.success("Анализ завершён. Показаны ключевые метрики. Excel — во вкладке «Excel-отчёт».")
        except Exception as exc:
            st.error(f"Ошибка анализа: {exc}")
            if show_tech:
                st.code(traceback.format_exc())
            return
        finally:
            for p in (inv_path, cap_path):
                try:
                    p.unlink(missing_ok=True)
                except OSError:
                    pass

    result = st.session_state.get("analysis_result")
    if result is None:
        st.success("Структура файлов распознана. Нажмите «Запустить анализ» в боковой панели.")
        return

    render_results(result, show_tech=show_tech)


if __name__ == "__main__":
    main()
