# -*- coding: utf-8 -*-
"""Streamlit UI — AI Агент: анализатор итогов инвентаризации торгового зала."""
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
    author_role_chart_df,
    author_shortage_share_df,
    conclusions_cards,
    detail_table,
    discrepancy_structure_df,
    executive_insights,
    kpi_cards,
    status_risk_table,
    store_ranking_df,
    store_risk_chart_df,
    summary_finance_block,
    top_overlap_chart_df,
    top_shortage_chart_df,
    top_surplus_chart_df,
)
from src.validators import (
    EXPECTED_FORMAT_HINT,
    validate_capitalization_bytes,
    validate_uploaded_inventory_bytes,
)

st.set_page_config(
    page_title="AI Агент — анализатор инвентаризации торгового зала",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

DISPLAY_ROW_LIMIT = 5_000

# Lightweight theme only (no animated custom DOM nodes — those caused removeChild)
_THEME_CSS = """
<style>
html, body, [class*="css"]  {
  font-size: 1.05rem;
}
h1 {
  font-weight: 800 !important;
  letter-spacing: -0.02em;
  color: #0f2942 !important;
}
h2, h3 {
  font-weight: 700 !important;
  color: #1a365d !important;
}
div[data-testid="stMetricValue"] {
  font-size: 1.55rem !important;
  font-weight: 800 !important;
}
div[data-testid="stMetricLabel"] {
  font-weight: 600 !important;
}
.nia-banner {
  padding: 0.85rem 1.15rem;
  border-radius: 10px;
  background: linear-gradient(90deg, #e8f1fb 0%, #f3f7fc 50%, #eef6f0 100%);
  border-left: 5px solid #1e5aa0;
  margin: 0.4rem 0 1rem 0;
  font-size: 1.05rem;
  font-weight: 600;
  color: #1a365d;
}
.nia-kpi {
  border-radius: 12px;
  padding: 0.75rem 0.9rem 0.35rem 0.9rem;
  margin-bottom: 0.55rem;
  border: 1px solid #d6e2ef;
  background: #fff;
  border-top: 4px solid var(--accent, #1e5aa0);
}
.nia-concl {
  border-radius: 10px;
  padding: 0.8rem 1rem;
  margin-bottom: 0.55rem;
  background: #f7fafc;
  border: 1px solid #e2e8f0;
  border-left: 4px solid #2b6cb0;
}
.nia-concl b { color: #1a365d; font-size: 1.05rem; }
.nia-agent-live {
  padding: 1rem 1.2rem;
  border-radius: 12px;
  background: linear-gradient(135deg, #0f2942 0%, #1e5aa0 55%, #2b6cb0 100%);
  color: #fff;
  margin: 0.5rem 0 1rem 0;
}
.nia-agent-live h3 { color: #fff !important; margin: 0 0 0.35rem 0; font-size: 1.35rem; }
.nia-agent-live p { margin: 0; opacity: 0.92; }
</style>
"""


def _inject_theme() -> None:
    st.markdown(_THEME_CSS, unsafe_allow_html=True)


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


def _plot_bar_h(df: pd.DataFrame, x: str, y: str, title: str, color: str = "#1e5aa0"):
    try:
        import plotly.express as px
    except ImportError:
        st.warning("Plotly не установлен — график недоступен.")
        return
    if df is None or df.empty:
        st.info("Недостаточно данных для построения графика.")
        return
    fig = px.bar(
        df, x=x, y=y, orientation="h", title=title, text_auto=".2s",
        color_discrete_sequence=[color],
    )
    fig.update_layout(
        yaxis={"categoryorder": "total ascending"},
        height=440,
        margin=dict(l=10, r=10, t=48, b=10),
        title_font=dict(size=18, color="#0f2942", family="Arial"),
        font=dict(size=13),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(247,250,252,1)",
    )
    fig.update_traces(textfont_size=12, textposition="outside", cliponaxis=False)
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
    fig = px.pie(
        df, names=names, values=values, title=title, hole=0.4,
        color_discrete_sequence=["#c53030", "#2b6cb0", "#d69e2e", "#dd6b20"],
    )
    fig.update_layout(
        height=440,
        margin=dict(l=10, r=10, t=48, b=10),
        title_font=dict(size=18, color="#0f2942"),
        font=dict(size=13),
    )
    fig.update_traces(textposition="inside", textinfo="percent+label", textfont_size=12)
    st.plotly_chart(fig, use_container_width=True)


def _plot_grouped_bar(df: pd.DataFrame, x: str, y: str, color: str, title: str):
    try:
        import plotly.express as px
    except ImportError:
        st.warning("Plotly не установлен — график недоступен.")
        return
    if df is None or df.empty:
        st.info("Недостаточно данных для построения графика.")
        return
    fig = px.bar(
        df, x=x, y=y, color=color, barmode="group", title=title, text_auto=".2s",
        color_discrete_map={"Недостачи": "#c53030", "Излишки": "#2b6cb0"},
    )
    fig.update_layout(
        height=420,
        margin=dict(l=10, r=10, t=48, b=10),
        title_font=dict(size=18, color="#0f2942"),
        font=dict(size=13),
        legend_title_text="",
    )
    st.plotly_chart(fig, use_container_width=True)


def render_sidebar():
    settings = load_settings()
    st.sidebar.markdown("### 🤖 AI Агент")
    st.sidebar.caption(f"Розница v{__version__}")

    try:
        master = get_master_hierarchy()
        st.sidebar.caption(
            f"Эталон: {master.source_file} · {len(master.items):,} SKU"
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
    # Advanced options hidden from employee UI (defaults are safe/fast):
    # enable_cross_store=False, end_date=None, show_tech=False
    enable_cross = False
    end_date = None
    show_tech = False
    run = st.sidebar.button("Запустить анализ", type="primary", use_container_width=True)

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
4. Второй файл — оприходование излишков при закрытии смены.

### Excel-отчёт

После анализа откройте вкладку «Excel-отчёт» и нажмите «Сформировать Excel-отчёт».
        """
    )
    st.info(EXPECTED_FORMAT_HINT)


def render_home(result):
    fin = summary_finance_block(result)
    st.markdown(
        f'<div class="nia-banner">Сводка · период <b>{fin["period"]}</b> · '
        f'{fin["stores"]} магазинов · {fin["docs"]:,} документов · '
        f'сальдо <b>{fin["net"]:,.0f} ₽</b></div>',
        unsafe_allow_html=True,
    )

    st.markdown("## Ключевые показатели")
    cards = kpi_cards(result)
    for row_start in range(0, len(cards), 4):
        cols = st.columns(4)
        for i, card in enumerate(cards[row_start:row_start + 4]):
            with cols[i]:
                accent = card.get("color", "#1e5aa0")
                st.markdown(
                    f'<div class="nia-kpi" style="--accent:{accent}">',
                    unsafe_allow_html=True,
                )
                st.metric(card["label"], card["value"], delta=card["delta"], help=card["help"])
                st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("## Выводы")
    concl = conclusions_cards(result)
    if concl:
        for c in concl[:6]:
            st.markdown(
                f'<div class="nia-concl"><b>{c["title"]}</b><br/>{c["text"]}</div>',
                unsafe_allow_html=True,
            )
    else:
        for line in executive_insights(result):
            st.markdown(f"- {line}")

    st.markdown("## Статусы и риски")
    st.dataframe(status_risk_table(result), use_container_width=True, hide_index=True)

    st.markdown("## Визуализации")
    c1, c2 = st.columns(2)
    with c1:
        _plot_bar_h(
            top_shortage_chart_df(result),
            x="Недостача, ₽",
            y="Позиция",
            title="Топ-10 позиций по недостаче",
            color="#c53030",
        )
    with c2:
        _plot_pie(
            discrepancy_structure_df(result),
            names="Тип",
            values="Сумма, ₽",
            title="Структура расхождений",
        )

    c3, c4 = st.columns(2)
    with c3:
        _plot_bar_h(
            top_surplus_chart_df(result),
            x="Излишек, ₽",
            y="Позиция",
            title="Топ-10 позиций по излишкам",
            color="#2b6cb0",
        )
    with c4:
        _plot_bar_h(
            top_overlap_chart_df(result),
            x="Перекрытие, ₽",
            y="Пара",
            title="Топ-10 пересортов (перекрытий)",
            color="#d69e2e",
        )

    _plot_bar_h(
        store_risk_chart_df(result),
        x="Чистые недостачи, ₽",
        y="Магазин",
        title="Рейтинг магазинов по чистым недостачам",
        color="#dd6b20",
    )


def render_rankings(result):
    st.markdown("## Рейтинг магазинов (классы A/B/C)")
    rank = store_ranking_df(result)
    if rank.empty:
        st.info("Нет данных рейтинга.")
    else:
        st.dataframe(rank, use_container_width=True, hide_index=True)
        if "Класс" in rank.columns:
            try:
                import plotly.express as px
                fig = px.histogram(
                    rank, x="Класс", title="Распределение магазинов по классам",
                    color="Класс",
                    color_discrete_map={"A": "#2f855a", "B": "#d69e2e", "C": "#c53030"},
                )
                fig.update_layout(height=360, title_font=dict(size=18), showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            except Exception:
                pass

    st.markdown("## Ревизоры vs Операторы")
    role_df = author_role_chart_df(result)
    share_df = author_shortage_share_df(result)
    r1, r2 = st.columns(2)
    with r1:
        _plot_grouped_bar(role_df, x="Роль", y="Сумма, ₽", color="Тип",
                          title="Вклад ролей: недостачи и излишки")
    with r2:
        _plot_pie(share_df, names="Роль", values="Недостачи, ₽",
                  title="Доля недостач: ревизор / оператор")

    a = result.author_sum or {}
    st.markdown(
        f'<div class="nia-banner">Доминирует: <b>{a.get("network_dominant", "—")}</b> '
        f'({a.get("network_sign", "")}) · цепочек пересортов: '
        f'<b>{int(a.get("chains_count", 0)):,}</b> на '
        f'<b>{float(a.get("chains_sum", 0) or 0):,.0f} ₽</b></div>',
        unsafe_allow_html=True,
    )

    chains = result.author_chains
    if chains is not None and not chains.empty:
        st.markdown("### Топ цепочек пересортов ревизор ↔ оператор")
        show_cols = [c for c in [
            "магазин", "наименование", "автор1", "знак1", "сумма1",
            "автор2", "знак2", "сумма2", "сумма_цепочки", "инициатор",
        ] if c in chains.columns]
        top_ch = chains.sort_values(
            "сумма_цепочки" if "сумма_цепочки" in chains.columns else show_cols[0],
            ascending=False,
        ).head(15)
        st.dataframe(top_ch[show_cols] if show_cols else top_ch.head(15),
                     use_container_width=True, hide_index=True)


def render_tops(result):
    st.markdown("## Топы сети")
    t1, t2 = st.columns(2)
    with t1:
        st.markdown("### Недостачи")
        st.dataframe(top_shortage_chart_df(result, n=20), use_container_width=True, hide_index=True)
    with t2:
        st.markdown("### Излишки")
        st.dataframe(top_surplus_chart_df(result, n=20), use_container_width=True, hide_index=True)

    st.markdown("### Пересорты / перекрытия")
    st.dataframe(top_overlap_chart_df(result, n=20), use_container_width=True, hide_index=True)

    chronic = result.chronic
    if chronic is not None and not chronic.empty:
        st.markdown("### Хронические проблемные позиции")
        cols = [c for c in ["наименование", "магазинов", "документов", "недостачи", "излишки"] if c in chronic.columns]
        st.dataframe(chronic[cols].head(20) if cols else chronic.head(20),
                     use_container_width=True, hide_index=True)

    anom = result.anom_df
    if anom is not None and not anom.empty:
        st.markdown("### Аномалии книжных сумм (критичные)")
        crit = anom
        if "критичность" in anom.columns:
            crit = anom[anom["критичность"] == "ДА"]
        st.dataframe(crit.head(20), use_container_width=True, hide_index=True)


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
        st.warning(f"Показаны первые {DISPLAY_ROW_LIMIT:,} строк из {len(view):,}.")
        show = view.head(DISPLAY_ROW_LIMIT)
    else:
        show = view
    st.dataframe(show, use_container_width=True, hide_index=True)

    d1, d2 = st.columns(2)
    d1.download_button("Скачать фильтр CSV", data=_df_to_csv_bytes(view),
                       file_name="detail_filtered.csv", mime="text/csv")
    d2.download_button("Скачать фильтр Excel", data=_df_to_xlsx_bytes(view),
                       file_name="detail_filtered.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_quality(result):
    q = result.quality
    if q.passed and not q.issues:
        st.success("Проверки пройдены")
    elif q.passed:
        st.warning("Анализ выполнен, есть замечания по качеству данных")
    else:
        st.error("Обнаружены проблемы во входных данных")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Строк после периода", f"{q.after_period_rows:,}")
    m2.metric("Исключено периодом", f"{q.excluded_by_period:,}")
    m3.metric("Вне справочника", f"{q.hierarchy_unmatched_rows:,}")
    m4.metric("Аномалий / крит.", f"{q.anomaly_total} / {q.anomaly_critical}")

    if q.issues:
        st.dataframe(pd.DataFrame(q.issues), use_container_width=True, hide_index=True)
    else:
        st.info("Критичных замечаний по качеству нет.")


def render_excel_tab(result, show_tech: bool = False):
    st.markdown(
        "Итоговый Excel = та же `build_report`, что CLI. "
        "Сборка **по кнопке**, чтобы не блокировать BI-дашборд."
    )
    ready = bool(result.excel_bytes) and not getattr(result, "excel_pending", False)

    if not ready:
        st.info("Excel ещё не сформирован.")
        if st.button("Сформировать Excel-отчёт", type="primary", key="build_excel_btn"):
            live = st.empty()
            bar = st.progress(0, text="Подготовка Excel…")

            def _cb(v: float, text: str) -> None:
                live.markdown(
                    f'<div class="nia-agent-live"><h3>🤖 AI-агент формирует Excel</h3>'
                    f"<p>{text}</p></div>",
                    unsafe_allow_html=True,
                )
                bar.progress(min(max(int(v * 100), 0), 100), text=text)

            try:
                result = ensure_excel_bytes(result, progress_cb=_cb)
                st.session_state["analysis_result"] = result
                live.markdown(
                    '<div class="nia-agent-live"><h3>🤖 Excel готов</h3>'
                    "<p>Можно скачать отчёт.</p></div>",
                    unsafe_allow_html=True,
                )
                bar.progress(100, text="Excel готов")
                ready = True
            except Exception as exc:
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
        f"Время: {result.excel_seconds:.1f} с"
    )
    st.download_button(
        "Скачать Excel-отчёт",
        data=result.excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


def render_results(result, show_tech: bool = False, caption: str | None = None) -> None:
    tabs = st.tabs([
        "Главная", "Рейтинги и роли", "Топы", "Детализация",
        "Контроль качества", "Excel-отчёт", "Инструкция",
    ])
    with tabs[0]:
        render_home(result)
    with tabs[1]:
        render_rankings(result)
    with tabs[2]:
        render_tops(result)
    with tabs[3]:
        render_detail(result)
    with tabs[4]:
        render_quality(result)
    with tabs[5]:
        render_excel_tab(result, show_tech=show_tech)
    with tabs[6]:
        render_instruction_tab()
    if caption:
        st.caption(caption)


def _run_analysis(period_days, enable_cross, end_date, show_tech) -> bool:
    """Run single-pass UI analysis. Returns True on success."""
    inv_path = _save_bytes(
        st.session_state["inv_bytes"],
        Path(st.session_state["inv_name"]).suffix or ".xlsx",
    )
    cap_path = _save_bytes(
        st.session_state["cap_bytes"],
        Path(st.session_state["cap_name"]).suffix or ".xlsx",
    )
    live = st.empty()
    bar = st.progress(0, text="AI-агент приступает к работе…")
    live.markdown(
        '<div class="nia-agent-live"><h3>🤖 AI-агент выполняет работу и готовит анализ</h3>'
        "<p>Приложение не зависло — следите за статусом ниже.</p></div>",
        unsafe_allow_html=True,
    )

    def _cb(v: float, text: str) -> None:
        # Single placeholder update — never append (avoids React removeChild)
        live.markdown(
            f'<div class="nia-agent-live"><h3>🤖 AI-агент работает</h3>'
            f"<p><b>Сейчас:</b> {text}</p></div>",
            unsafe_allow_html=True,
        )
        bar.progress(min(max(int(v * 100), 0), 100), text=text)

    try:
        # Validate only at run-time (not on upload) to avoid DOM thrash / removeChild
        inv_val = validate_uploaded_inventory_bytes(
            st.session_state["inv_bytes"], st.session_state["inv_name"]
        )
        cap_val = validate_capitalization_bytes(
            st.session_state["cap_bytes"], st.session_state["cap_name"]
        )
        if not inv_val.ok:
            st.error(inv_val.user_message)
            if show_tech:
                st.code("\n".join(inv_val.errors))
            return False
        if not cap_val.ok:
            st.error(cap_val.user_message)
            if show_tech:
                st.code("\n".join(cap_val.errors))
            return False
        for w in inv_val.warnings:
            st.warning(w)

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
        st.session_state["analysis_result"] = result
        st.session_state["analysis_fp"] = result.source_fingerprint
        live.markdown(
            '<div class="nia-agent-live"><h3>🤖 Анализ готов</h3>'
            "<p>Открыт BI-дашборд. Excel — во вкладке «Excel-отчёт».</p></div>",
            unsafe_allow_html=True,
        )
        bar.progress(100, text="Анализ готов")
        st.success("Анализ завершён.")
        return True
    except Exception as exc:
        st.error(f"Ошибка анализа: {exc}")
        if show_tech:
            st.code(traceback.format_exc())
        return False
    finally:
        for p in (inv_path, cap_path):
            try:
                p.unlink(missing_ok=True)
            except OSError:
                pass


def main() -> None:
    _inject_theme()
    st.title("AI Агент — анализатор итогов инвентаризации торгового зала")
    st.markdown(
        "Загрузите файлы → **Запустить анализ** → BI-дашборд со сводкой, выводами, "
        "топами и рейтингами. Excel формируется отдельно."
    )

    inv, cap, period_days, enable_cross, end_date, show_tech, run = render_sidebar()
    result_cached = st.session_state.get("analysis_result")

    # Stable empty-state: no expander swap on upload (prevents removeChild)
    if inv is None or cap is None:
        if result_cached is not None:
            render_results(
                result_cached,
                show_tech=show_tech,
                caption="Файлы в загрузчике сброшены — показаны метрики сессии. "
                        "Чтобы пересчитать, загрузите оба файла снова.",
            )
            return
        st.info(
            "Загрузите в боковой панели **два** Excel-файла "
            "(инвентаризация + оприходование), затем нажмите «Запустить анализ»."
        )
        st.markdown("Краткая инструкция — во вкладке после анализа или ниже.")
        with st.expander("Как подготовить файлы", expanded=False):
            render_instruction_tab()
        return

    # Cache bytes only — NO heavy validation here (validation on Run)
    inv_bytes = inv.getvalue()
    cap_bytes = cap.getvalue()
    upload_key = f"{inv.name}|{len(inv_bytes)}|{cap.name}|{len(cap_bytes)}|{period_days}|{enable_cross}|{end_date}"
    prev_key = st.session_state.get("upload_key")
    if prev_key != upload_key:
        if "analysis_result" in st.session_state and prev_key is not None:
            st.session_state.pop("analysis_result", None)
        st.session_state["upload_key"] = upload_key
        st.session_state["inv_bytes"] = inv_bytes
        st.session_state["cap_bytes"] = cap_bytes
        st.session_state["inv_name"] = inv.name
        st.session_state["cap_name"] = cap.name

    if run:
        st.session_state["pending_run"] = True

    if st.session_state.get("pending_run"):
        st.session_state["pending_run"] = False
        ok = _run_analysis(period_days, enable_cross, end_date, show_tech)
        if not ok:
            return

    result = st.session_state.get("analysis_result")
    if result is None:
        st.markdown(
            '<div class="nia-banner">Файлы загружены. Нажмите '
            "<b>«Запустить анализ»</b> в боковой панели — AI-агент посчитает метрики.</div>",
            unsafe_allow_html=True,
        )
        return

    render_results(result, show_tech=show_tech)


if __name__ == "__main__":
    main()
