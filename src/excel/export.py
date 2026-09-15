# -*- coding: utf-8 -*-
"""Excel report builders and full report pipeline."""
from __future__ import annotations

import datetime
import io
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.hyperlink import Hyperlink

from config.settings import TOP_N_NETWORK, TOP_N_STORE
from src.anomalies import (
    TYPE_1,
    TYPE_2,
    TYPE_3,
    anomalies_summary,
    detect_book_sum_anomalies,
    enrich_store_metrics_with_anomalies,
)
from src.capitalization import (
    STATUS_MATCHED,
    STATUS_MATCHED_SHORTAGE,
    STATUS_UNCERTAIN_STORE,
    STATUS_UNMATCHED_NAME,
    STATUS_UNMATCHED_STORE,
    capitalization_summary,
    enrich_store_metrics_with_capitalization,
    match_capitalization_to_inventory,
    parse_capitalization_excel,
)
from src.catalog import enrich_dataframe, load_catalog, unmatched_summary
from src.master_hierarchy import UNMATCHED_LABEL, get_master_hierarchy
from src.author import (
    annotate_overlap_authors,
    author_network_summary,
    calc_author_role_stats,
    calc_author_store_stats,
    detect_author_chains,
    enrich_store_metrics_with_authors,
)
from src.excel.styles import (
    ACCENT,
    BD,
    DF,
    GD,
    GF,
    GL,
    HD,
    HG,
    HM,
    HOMOGENEITY_RU,
    HR,
    INT,
    LB,
    LEVEL_RU,
    NUM,
    ORG,
    PCT,
    PF,
    RD,
    RF,
    SCENARIO_COLORS,
    SCENARIO_PRIORITY_COLOR,
    SF,
    SH,
    SUBT_F,
    SUM,
    TF,
    WH,
    YF,
    cs3,
    dbars,
    hrow,
    mtitle,
    nc,
    safe_sheet_name,
    set_col_widths,
    tc,
    trow,
    zebra_fill,
)
from src.metrics import (
    build_sku_store_pivot,
    calc_document_metrics,
    calc_network_summary,
    calc_sku_cross,
    calc_store_metrics,
    generate_conclusions,
)
from src.models import Config
from src.overlap import build_analytical_cross_store, build_operational_overlaps
from src.parser import filter_by_period, parse_network_excel
from src.scenarios import assign_scenario

def _display_category_value(val) -> str:
    s = "" if val is None else str(val)
    if s.startswith("__unmatched__"):
        return UNMATCHED_LABEL
    return s


def _for_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Prepare dataframe for Excel sheets: human-readable category fields."""
    if df is None or df.empty:
        return df
    out = df.copy()
    if "category_label" in out.columns:
        out["category_group"] = out["category_label"].map(_display_category_value)
    elif "category_group" in out.columns:
        out["category_group"] = out["category_group"].map(_display_category_value)
    return out


def write_df_table(ws, start_row, headers, rows_data, col_fmts=None, zebra=True, hdr_fill=None):
    hrow(ws, start_row, headers, fill=hdr_fill or HD)
    r = start_row + 1
    for i, row_vals in enumerate(rows_data):
        fill = PatternFill("solid", fgColor="FAFAFA") if zebra and i % 2 else None
        for j, val in enumerate(row_vals, 1):
            fmt = (col_fmts or {}).get(j)
            if isinstance(val, (int, float)) and fmt:
                nc(ws, r, j, val, fmt, fill)
            else:
                tc(ws, r, j, val, fill, wrap=(j == 2))
        r += 1
    return r


def sh_contents(wb, sheets_info):
    ws = wb.create_sheet("Содержание", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    for i, w in enumerate([6, 42, 55], 2):
        ws.column_dimensions[GL(i)].width = w
    mtitle(ws, 2, 2, 4, "АНАЛИТИКА ИНВЕНТАРИЗАЦИЙ СЕТИ — СОДЕРЖАНИЕ",
           Font(name="Calibri", bold=True, size=14, color=WH), HD)
    hrow(ws, 4, ["№", "Лист", "Описание"], [HM, HM, HM], ht=22, col_start=2)
    for i, (sname, desc, loc) in enumerate(sheets_info, 1):
        r = 4 + i
        nc(ws, r, 2, i, INT)
        sheet_name, cell_ref = loc.split("!")
        c = ws.cell(row=r, column=3, value=f"> {sname}")
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{sheet_name}'!{cell_ref}")
        c.font = Font(name="Calibri", size=10, color="0000FF", underline="single")
        c.border = BD
        tc(ws, r, 4, desc)
    tc(ws, 4 + len(sheets_info) + 2, 2,
       f"Сформировано: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}")


def _dash_paragraph(ws, row: int, text: str, *, height: int = 56, italic: bool = False) -> int:
    """Merged methodology/context paragraph on Сводка; returns next free row."""
    ws.merge_cells(f"A{row}:H{row}")
    c = ws.cell(row=row, column=1, value=text)
    c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
    c.font = Font(name="Calibri", size=9, italic=italic, color="404040")
    ws.row_dimensions[row].height = height
    return row + 1


def sh_dashboard(wb, summary, store_metrics, period_str, cap_summary=None,
                 anom_summary=None, author_summary=None):
    from src import __version__ as _ver

    ws = wb.create_sheet("Сводка")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {1: 28, 2: 18, 3: 18, 4: 18, 5: 18, 6: 18, 7: 18, 8: 18})
    mtitle(ws, 2, 1, 8, f"СВОДКА | {period_str}",
           Font(name="Calibri", bold=True, size=14, color=WH), HD)

    # --- Контекст релиза / AI ---
    mtitle(ws, 3, 1, 8,
           f"РЕЛИЗ 2 · retail_inventory_analyzer v{_ver} · работа AI-агента Cursor",
           Font(name="Calibri", bold=True, size=10, color=WH), HM, "left")
    ctx = (
        "Этот отчёт сформирован модулем анализа инвентаризаций РОЗНИЦЫ (без собственного производства). "
        "Модуль разработан и адаптирован AI-агентом (Cursor) на базе логики первого релиза сетевого анализа "
        "(v3.0 / анализ_сеть_последняя_с_группами) с существенным обогащением: эталонная иерархия 1С, "
        "аномалии книжных сумм, сверка оприходования закрытия смены, разрез ревизор/оператор и цепочки пересортов. "
        "Проверка цен между магазинами в розничном контуре намеренно отключена. "
        "Цифры ниже — аналитическая интерпретация выгрузки TDSheet, а не проводки в учётную систему."
    )
    next_row = _dash_paragraph(ws, 4, ctx, height=62)

    kpis = [
        ("Магазинов", summary["stores"], INT, ACCENT, LB),
        ("Документов", summary["docs"], INT, ACCENT, LB),
        ("Излишки", summary["surplus"], SUM, GD, GF),
        ("Недостачи", summary["shortage"], SUM, RD, RF),
        ("Перекрытие", summary["overlap"], SUM, ORG, YF),
        ("Чистые недостачи", summary["clean_shortage"], SUM, RD, RF),
    ]
    kpi_row = next_row
    for idx, (lbl, val, fmt, col, bg) in enumerate(kpis):
        cc = idx + 1
        for rr, vv in [(kpi_row, lbl), (kpi_row + 1, val), (kpi_row + 2, "руб." if fmt == SUM else "шт.")]:
            cell = ws.cell(row=rr, column=cc, value=vv)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = Font(name="Calibri", bold=(rr == kpi_row + 1), size=13 if rr == kpi_row + 1 else 8,
                             color=col if rr == kpi_row + 1 else "606060")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BD
            if rr == kpi_row + 1 and fmt == SUM:
                cell.number_format = SUM
    next_row = kpi_row + 3
    next_row = _dash_paragraph(
        ws, next_row,
        "Как читать KPI: «Излишки» и «Недостачи» — суммы из исходной выгрузки (факт vs учёт). "
        "«Перекрытие» — оценка однородного пересорта внутри категории эталона (жадное сопоставление "
        "излишек↔недостача, балл схожести ≥ 20). «Чистые недостачи» = Недостачи − Перекрытие — "
        "часть, которую нельзя объяснить пересортом; именно она ближе к реальному риску потерь. "
        "Сальдо (излишки − недостачи) на KPI-ленте не выводится отдельно: смотрите структуру недостач ниже.",
        height=52, italic=True,
    )
    next_row += 1

    if anom_summary and int(anom_summary.get("total", 0) or 0) > 0:
        mtitle(ws, next_row, 1, 8, "АНОМАЛИИ КНИЖНЫХ СУММ", SUBT_F, SH, "left")
        next_row += 1
        anom_kpis = [
            ("Всего аномалий", int(anom_summary.get("total", 0)), INT, ORG, YF),
            ("Критических (тип 2)", int(anom_summary.get("critical", 0)), INT, RD, RF),
            ("Тип 1 (нет норм.)", int(anom_summary.get("type1", 0)), INT, ORG, YF),
            ("Тип 3 (только факт.)", int(anom_summary.get("type3", 0)), INT, ORG, YF),
            ("Доля строк", float(anom_summary.get("share_pct", 0)) / 100.0, PCT, ACCENT, LB),
            ("Излишки от аномалий", float(anom_summary.get("surplus_impact", 0)), SUM, GD, GF),
        ]
        for idx, (lbl, val, fmt, col, bg) in enumerate(anom_kpis):
            cc = idx + 1
            for rr, vv in [(next_row, lbl), (next_row + 1, val)]:
                cell = ws.cell(row=rr, column=cc, value=vv)
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.font = Font(name="Calibri", bold=(rr == next_row + 1),
                                 size=12 if rr == next_row + 1 else 8,
                                 color=col if rr == next_row + 1 else "606060")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = BD
                if rr == next_row + 1 and fmt == SUM:
                    cell.number_format = SUM
                if rr == next_row + 1 and fmt == PCT:
                    cell.number_format = PCT
        next_row += 3
        meanings = [
            f"{TYPE_1}: нет книжной нормативной → пересчёт может уйти в излишки (риск нулевой с/с).",
            f"{TYPE_2}: нет обеих книжных сумм → нет недостач/излишков по сумме, итог искажён.",
            f"{TYPE_3}: только фактическая книжная → ручной пересчёт, нужна сверка по складам.",
        ]
        for msg in meanings:
            ws.merge_cells(f"A{next_row}:H{next_row}")
            c = ws.cell(row=next_row, column=1, value=msg)
            c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
            c.font = Font(name="Calibri", size=9, italic=True)
            next_row += 1
        next_row = _dash_paragraph(
            ws, next_row,
            "Интерпретация: аномалии — сигнал качества данных, не прямые потери. "
            "Тип 2 критичен (искажает суммы расхождений). Перед управленческими выводами "
            "по магазину с высокой долей аномалий сначала восстановите книжные суммы в 1С.",
            height=40, italic=True,
        )
        next_row += 1

    if author_summary:
        mtitle(ws, next_row, 1, 8, "РЕВИЗОРЫ VS ОПЕРАТОРЫ", SUBT_F, SH, "left")
        next_row += 1
        author_kpis = [
            ("Док. ревизора", int(author_summary.get("rev_docs", 0)), INT, ACCENT, LB),
            ("Док. оператора", int(author_summary.get("op_docs", 0)), INT, ACCENT, LB),
            ("Недост. ревизор", float(author_summary.get("rev_shortage", 0)), SUM, RD, RF),
            ("Недост. оператор", float(author_summary.get("op_shortage", 0)), SUM, RD, RF),
            ("Пересорт ревизор", float(author_summary.get("rev_overlap", 0)), SUM, ORG, YF),
            ("Пересорт оператор", float(author_summary.get("op_overlap", 0)), SUM, ORG, YF),
        ]
        for idx, (lbl, val, fmt, col, bg) in enumerate(author_kpis):
            cc = idx + 1
            for rr, vv in [(next_row, lbl), (next_row + 1, val)]:
                cell = ws.cell(row=rr, column=cc, value=vv)
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.font = Font(name="Calibri", bold=(rr == next_row + 1),
                                 size=12 if rr == next_row + 1 else 8,
                                 color=col if rr == next_row + 1 else "606060")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = BD
                if rr == next_row + 1 and fmt == SUM:
                    cell.number_format = SUM
        next_row += 3
        anote = (
            f"Доля недостач: ревизор {author_summary.get('rev_share_shortage', 0):.1f}% / "
            f"оператор {author_summary.get('op_share_shortage', 0):.1f}%. "
            f"Сальдо ревизора: {author_summary.get('rev_net', 0):,.0f} руб.; "
            f"оператора: {author_summary.get('op_net', 0):,.0f} руб. "
            f"Больший вклад в результат сети: {author_summary.get('network_dominant')} "
            f"({author_summary.get('network_sign')}). "
            f"Цепочек пересортов ревизор↔оператор: {author_summary.get('chains_count', 0)} "
            f"({author_summary.get('chains_sum', 0):,.0f} руб.)."
        )
        ws.merge_cells(f"A{next_row}:H{next_row}")
        cnote = ws.cell(row=next_row, column=1, value=anote)
        cnote.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        cnote.font = Font(name="Calibri", size=9, italic=True)
        ws.row_dimensions[next_row].height = 40
        next_row += 1
        next_row = _dash_paragraph(
            ws, next_row,
            "Методика автора документа (новое в релизе 2): время из названия "
            "«Инвентаризация … от ДД.ММ.ГГГГ ЧЧ:ММ:СС». Ровно 08:00:00 → ревизор; "
            "любое иное распознанное время (в т.ч. 8:00 без ведущего нуля, 21:00, 10:53:34) → оператор; "
            "время не извлечено → «неизвестно». Пересорт по роли атрибутируется автору документа недостачи. "
            "Цепочка: одна позиция (sku_key) в магазине, разные роли, противоположный знак во времени "
            "(недостача→излишек или излишек→недостача); инициатор = автор первого документа. "
            "Важно: разрез не меняет формулу чистых недостач и балл рейтинга A/B/C — только интерпретацию.",
            height=68, italic=True,
        )
        next_row += 1

    if cap_summary:
        mtitle(ws, next_row, 1, 8, "ОПРИХОДОВАНИЕ ИЗЛИШКОВ (ЗАКРЫТИЕ СМЕНЫ)", SUBT_F, SH, "left")
        next_row += 1
        cap_kpis = [
            ("Складов (файл)", int(cap_summary.get("warehouses", 0)), INT, ACCENT, LB),
            ("Строк оприход.", int(cap_summary.get("cap_rows", 0)), INT, ACCENT, LB),
            ("Сумма оприход.", float(cap_summary.get("cap_sum", 0)), SUM, GD, GF),
            ("Сопоставлено", int(cap_summary.get("matched", 0)) + int(cap_summary.get("matched_shortage", 0)), INT, GD, GF),
            ("Связь с недост.", int(cap_summary.get("matched_shortage", 0)), INT, RD, RF),
            ("Не сопоставлено", int(cap_summary.get("unmatched_name", 0))
             + int(cap_summary.get("unmatched_store", 0))
             + int(cap_summary.get("uncertain_store", 0)), INT, ORG, YF),
        ]
        for idx, (lbl, val, fmt, col, bg) in enumerate(cap_kpis):
            cc = idx + 1
            for rr, vv in [(next_row, lbl), (next_row + 1, val)]:
                cell = ws.cell(row=rr, column=cc, value=vv)
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.font = Font(name="Calibri", bold=(rr == next_row + 1), size=12 if rr == next_row + 1 else 8,
                                 color=col if rr == next_row + 1 else "606060")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = BD
                if rr == next_row + 1 and fmt == SUM:
                    cell.number_format = SUM
        next_row += 3
        note = (
            "Оприходование сопоставляется с инвентаризацией только по точному наименованию "
            "и только внутри своего склада/подразделения (без смешивания точек)."
        )
        ws.merge_cells(f"A{next_row}:H{next_row}")
        cnote = ws.cell(row=next_row, column=1, value=note)
        cnote.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        cnote.font = Font(name="Calibri", size=9, italic=True)
        next_row += 1
        next_row = _dash_paragraph(
            ws, next_row,
            "Интерпретация оприходования: «Сопоставлено» — строка закрытия смены нашла то же имя "
            "в инвентаризации своего склада; «↔ Недостача» — при этом по позиции есть недостача "
            "(кандидат на закрытие смены vs пересчёт). Несопоставленное требует ручной проверки "
            "карточки/склада. Модуль не создаёт документов оприходования — только сверка.",
            height=48, italic=True,
        )
        next_row += 1

    mtitle(ws, next_row, 1, 8, "МЕТОДИКА РАСЧЁТОВ И ВЫВОДОВ (РЕЛИЗ 2)", SUBT_F, SH, "left")
    next_row += 1
    method_blocks = [
        (
            "1. Категории и однородность. "
            "Категория берётся только из эталона Шаблон_иерархия_безпроизводства.xlsx "
            "(leaf_group / category_group). Keyword-эвристики первого релиза отключены. "
            "SKU вне справочника → лист «Вне справочника», отдельная группа без ложного перекрытия."
        ),
        (
            "2. Перекрытие (пересорт). "
            "Внутри магазина сопоставляются излишки и недостачи одной категории; "
            "балл по артикулу/токенам/весу/цене; жадное покрытие остатков количества и суммы. "
            "Операционный лист «Перекрытие» — для аналитики; на сырые излишки/недостачи не влияет."
        ),
        (
            "3. Рейтинг магазинов A/B/C. "
            "Балл = взвешенные процентили: коэффициент потерь (чистые/учёт), доля перекрытия, "
            "чистые недостачи, число SKU с расхождениями. Классы не пересчитываются от роли автора — "
            "роли добавлены колонками для интерпретации «кто сформировал результат»."
        ),
        (
            "4. Мероприятия СЦ-1…СЦ-7. "
            "Сценарии по порогам сумм и характеру расхождений (крупная/средняя недостача, "
            "весовые потери и др.) — план действий, не автоматическое закрытие в 1С."
        ),
        (
            "5. Как пользоваться выводами. "
            "Сначала чистые недостачи и хронические позиции; затем вклад ревизора/оператора "
            "и цепочки пересортов; аномалии книжных сумм — проверка качества данных; "
            "оприходование — сверка закрытия смены. Детали — листы «Выводы», «Мероприятия», "
            "«Ревизоры vs Операторы», «Цепочки пересортов»."
        ),
    ]
    for block in method_blocks:
        next_row = _dash_paragraph(ws, next_row, block, height=44)
    next_row += 1

    mtitle(ws, next_row, 1, 8, "ЧТО ТАКОЕ «ЧИСТЫЕ НЕДОСТАЧИ»", SUBT_F, SH, "left")
    expl = (
        "Чистые недостачи — это сумма недостач, которую нельзя объяснить пересортом внутри одной товарной группы.\n"
        "Перекрытие допустимо только при однородности категории: например, разные виды молочной группы между собой, "
        "но не мясо против напитков и не овощи против молочки.\n"
        "Если излишек неоднороден недостаче, он не должен её перекрывать автоматически — такая позиция требует "
        "перепроверки, сверки непроведённых приходов, уточнения карточки товара и анализа ошибок прошлых пересчётов.\n"
        "Интерпретация: высокая доля перекрытия → искать пересорт/ошибки взвешивания/карточек; "
        "высокая доля чистых недостач → искать пропажу, непроведённые документы, кражи, системные ошибки учёта."
    )
    ws.merge_cells(f"A{next_row + 1}:H{next_row + 2}")
    c = ws.cell(row=next_row + 1, column=1, value=expl)
    c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
    c.font = Font(name="Calibri", size=10)
    ws.row_dimensions[next_row + 1].height = 56
    ws.row_dimensions[next_row + 2].height = 56

    er_base = next_row + 4
    mtitle(ws, er_base, 1, 8, "ПРАВИЛА ОДНОРОДНОСТИ И ПРИМЕРЫ", SUBT_F, SH, "left")
    hrow(ws, er_base + 1, ["Недостача", "Излишек", "Можно перекрыть?", "Что делать"], HM, col_start=1)
    examples = [
        ("Мясо / птица", "Напитки", "НЕТ", "Перепроверка, сверка приходов, поиск ошибки пересчёта"),
        ("Овощи / фрукты", "Молочная группа", "НЕТ", "Уточнение карточки, проверка непроведённых документов"),
        ("Колбасные изделия", "Кондитерские изделия", "НЕТ", "Ручная сверка, акт расхождений"),
        ("Молоко 2,5%", "Кефир 1%", "ДА", "Оформить пересорт внутри молочной группы"),
        ("Хлеб белый", "Хлеб ржаной", "ДА", "Пересорт внутри хлебобулочной группы"),
    ]
    er = er_base + 2
    for ned, izl, ok, act in examples:
        fl = SF if ok == "ДА" else PF
        tc(ws, er, 1, ned, fl, wrap=True)
        tc(ws, er, 2, izl, fl, wrap=True)
        tc(ws, er, 3, ok, fl, bold=True)
        tc(ws, er, 4, act, fl, wrap=True)
        ws.row_dimensions[er].height = 28
        er += 1

    mtitle(ws, er + 1, 1, 4, "СТРУКТУРА НЕДОСТАЧ", SUBT_F, SH, "left")
    hrow(ws, er + 2, ["Показатель", "Сумма, руб.", "Доля"], HM, col_start=1)
    struct_rows = [
        ("Всего недостач", summary["shortage"], 1.0),
        ("Перекрыто однородным пересортом", summary["overlap"],
         (summary["overlap"] / summary["shortage"]) if summary["shortage"] else 0),
        ("Чистые недостачи (не перекрыты)", summary["clean_shortage"],
         (summary["clean_shortage"] / summary["shortage"]) if summary["shortage"] else 0),
        ("Излишки без однородного перекрытия (на проверку)", summary.get("non_homogeneous_surplus", 0), None),
    ]
    r = er + 3
    for i, (lbl, val, pct) in enumerate(struct_rows):
        fl = DF if "Чистые" in lbl else (PF if "без однородного" in lbl else (SF if "Перекрыто" in lbl else None))
        tc(ws, r, 1, lbl, fl, wrap=True)
        nc(ws, r, 2, val, SUM, fl, bold=("Чистые" in lbl))
        if pct is not None:
            nc(ws, r, 3, pct, PCT, fl)
        else:
            tc(ws, r, 3, "—", fl)
        r += 1
    dbars(ws, f"B{er + 3}:B{r-1}", "F8696B")
    r = _dash_paragraph(
        ws, r,
        "Интерпретация структуры: «Перекрыто» — потенциал оформления пересорта; "
        "«Чистые» — приоритет для расследования потерь; "
        "«Излишки без однородного перекрытия» — излишки, не нашедшие парную недостачу в своей группе "
        "(возможны ошибки карточки, чужая категория, непроведённый расход).",
        height=44, italic=True,
    )

    mtitle(ws, r + 1, 1, 8, "РЕЙТИНГ МАГАЗИНОВ (ЛУЧШИЕ / ХУДШИЕ)", SUBT_F, SH, "left")
    r = _dash_paragraph(
        ws, r + 2,
        "Класс A/B/C — относительный рейтинг внутри текущей выборки (процентили), не абсолютная оценка сети. "
        "Смотрите также долю ревизора/оператора на полном листе «Рейтинг магазинов»: "
        "магазин с хорошим баллом, но сильным вкладом операторских недостач, интерпретируется иначе, "
        "чем точка, где результат сформирован ревизией в 08:00:00.",
        height=48, italic=True,
    )
    has_cap = "оприходование_сумма" in store_metrics.columns
    if has_cap:
        hdr = ["Место", "Магазин", "Класс", "Балл", "Чист. недост.", "Оприходование", "Доля перекрытия"]
    else:
        hdr = ["Место", "Магазин", "Класс", "Балл", "Чист. недост.", "Коэфф. потерь", "Доля перекрытия"]
    hrow(ws, r, hdr, col_start=1)
    rr = r + 1
    for _, row in store_metrics.head(10).iterrows():
        fl = SF if row["класс"] == "A" else (PF if row["класс"] == "B" else DF)
        nc(ws, rr, 1, int(row["место"]), INT, fl)
        tc(ws, rr, 2, row["магазин"], fl)
        tc(ws, rr, 3, row["класс"], fl, bold=True)
        nc(ws, rr, 4, row["store_score"], SUM, fl)
        nc(ws, rr, 5, row["чистые_недостачи"], SUM, fl)
        if has_cap:
            nc(ws, rr, 6, float(row.get("оприходование_сумма", 0) or 0), SUM, fl)
        else:
            nc(ws, rr, 6, row["shrinkage_pct"] / 100, PCT, fl)
        nc(ws, rr, 7, row["recovery_pct"] / 100, PCT, fl)
        rr += 1
    if len(store_metrics) > 5:
        rr += 1
        mtitle(ws, rr, 1, 7, "ХУДШИЕ МАГАЗИНЫ", Font(name="Calibri", bold=True, size=11, color=WH), HR)
        rr += 1
        hrow(ws, rr, hdr, HR, col_start=1)
        rr += 1
        for _, row in store_metrics.tail(5).iterrows():
            nc(ws, rr, 1, int(row["место"]), INT, DF)
            tc(ws, rr, 2, row["магазин"], DF)
            tc(ws, rr, 3, row["класс"], DF, bold=True)
            nc(ws, rr, 4, row["store_score"], SUM, DF)
            nc(ws, rr, 5, row["чистые_недостачи"], SUM, DF)
            if has_cap:
                nc(ws, rr, 6, float(row.get("оприходование_сумма", 0) or 0), SUM, DF)
            else:
                nc(ws, rr, 6, row["shrinkage_pct"] / 100, PCT, DF)
            nc(ws, rr, 7, row["recovery_pct"] / 100, PCT, DF)
            rr += 1

    if not store_metrics.empty and len(store_metrics) >= 2:
        chart_row = rr + 1
        mtitle(ws, chart_row, 1, 4, "ЧИСТЫЕ НЕДОСТАЧИ ПО МАГАЗИНАМ", SUBT_F, SH, "left")
        cr = chart_row + 1
        hrow(ws, cr, ["Магазин", "Сумма"], HM, col_start=1)
        data_start = cr + 1
        for i, (_, row) in enumerate(store_metrics.head(15).iterrows()):
            tc(ws, data_start + i, 1, row["магазин"])
            nc(ws, data_start + i, 2, row["чистые_недостачи"], SUM)
        data_end = data_start + min(len(store_metrics), 15) - 1
        if data_end >= data_start:
            chart = BarChart()
            chart.title = "Чистые недостачи"
            chart.y_axis.title = "руб."
            data = Reference(ws, min_col=2, min_row=cr, max_row=data_end)
            cats = Reference(ws, min_col=1, min_row=data_start, max_row=data_end)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.height = 10
            chart.width = 20
            ws.add_chart(chart, f"E{chart_row}")


def sh_store_ranking(wb, store_metrics):
    ws = wb.create_sheet("Рейтинг магазинов")
    ws.sheet_view.showGridLines = False
    has_cap = "оприходование_сумма" in store_metrics.columns
    has_anom = "аномалий" in store_metrics.columns
    has_author = "доминанта" in store_metrics.columns
    extra = (2 if has_cap else 0) + (2 if has_anom else 0) + (5 if has_author else 0)
    ncols = 12 + extra
    widths = {i: 12 for i in range(1, ncols + 1)}
    widths[1] = 8
    widths[2] = 28
    widths[3] = 8
    widths[4] = 10
    set_col_widths(ws, widths)
    mtitle(ws, 2, 1, ncols, "РЕЙТИНГ МАГАЗИНОВ", Font(name="Calibri", bold=True, size=13, color=WH), HD)
    headers = ["Место", "Магазин", "Класс", "Балл", "Излишки", "Недостачи", "Сальдо",
               "Перекрытие", "Чист. недост.", "Коэфф. потерь", "Доля перекрытия", "Док."]
    if has_cap:
        headers.extend(["Оприходование", "Оп.↔недост."])
    if has_anom:
        headers.extend(["Аномалии", "Аном. крит."])
    if has_author:
        headers.extend([
            "Доля недост. рев.%", "Доля недост. оп.%",
            "Доля изл. рев.%", "Доля изл. оп.%",
            "Доминанта (роль)",
        ])
    hrow(ws, 4, headers, col_start=1)
    r = 5
    for _, row in store_metrics.iterrows():
        fl = SF if row["класс"] == "A" else (PF if row["класс"] == "B" else DF)
        nc(ws, r, 1, int(row["место"]), INT, fl)
        tc(ws, r, 2, row["магазин"], fl)
        tc(ws, r, 3, row["класс"], fl, bold=True)
        nc(ws, r, 4, row["store_score"], SUM, fl)
        nc(ws, r, 5, row["излишки"], SUM, fl)
        nc(ws, r, 6, row["недостачи"], SUM, fl)
        nc(ws, r, 7, row["сальдо"], SUM, fl)
        nc(ws, r, 8, row["перекрытие"], SUM, fl)
        nc(ws, r, 9, row["чистые_недостачи"], SUM, fl)
        nc(ws, r, 10, row["shrinkage_pct"] / 100, PCT, fl)
        nc(ws, r, 11, row["recovery_pct"] / 100, PCT, fl)
        nc(ws, r, 12, int(row["документов"]), INT, fl)
        col = 13
        if has_cap:
            nc(ws, r, col, float(row.get("оприходование_сумма", 0) or 0), SUM, fl)
            nc(ws, r, col + 1, int(row.get("оп_связь_с_недостачей", 0) or 0), INT, fl)
            col += 2
        if has_anom:
            nc(ws, r, col, int(row.get("аномалий", 0) or 0), INT, fl)
            nc(ws, r, col + 1, int(row.get("аномалий_критич", 0) or 0), INT, fl)
            col += 2
        if has_author:
            nc(ws, r, col, float(row.get("доля_рев_недостачи_%", 0) or 0) / 100, PCT, fl)
            nc(ws, r, col + 1, float(row.get("доля_оп_недостачи_%", 0) or 0) / 100, PCT, fl)
            nc(ws, r, col + 2, float(row.get("доля_рев_излишки_%", 0) or 0) / 100, PCT, fl)
            nc(ws, r, col + 3, float(row.get("доля_оп_излишки_%", 0) or 0) / 100, PCT, fl)
            dom = str(row.get("доминанта", "") or "")
            sign = str(row.get("доминанта_знак", "") or "")
            tc(ws, r, col + 4, f"{dom} ({sign})" if dom else "—", fl)
        r += 1
    dbars(ws, f"I5:I{r-1}", "F8696B")
    if has_cap and r > 5:
        dbars(ws, f"M5:M{r-1}", "63BE7B")
    ws.auto_filter.ref = f"A4:{GL(ncols)}{r-1}"
    ws.freeze_panes = "A5"


def sh_capitalization(wb, matched_df: pd.DataFrame, cap_summary: dict):
    """Dedicated sheet for shift-close surplus capitalization reconciliation."""
    ws = wb.create_sheet("Оприходование излишков")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {
        1: 28, 2: 24, 3: 36, 4: 12, 5: 14, 6: 18, 7: 14, 8: 14, 9: 14, 10: 40,
    })
    mtitle(ws, 2, 1, 10, "ОПРИХОДОВАНИЕ ИЗЛИШКОВ ПРИ ЗАКРЫТИИ СМЕНЫ",
           Font(name="Calibri", bold=True, size=13, color=WH), HD)
    tc(ws, 3, 1,
       "Точное сопоставление по наименованию внутри склада/подразделения. Данные разных складов не смешиваются.")

    # KPI strip
    kpis = [
        ("Строк", int(cap_summary.get("cap_rows", 0)), INT),
        ("Сумма", float(cap_summary.get("cap_sum", 0)), SUM),
        ("Matched", int(cap_summary.get("matched", 0)), INT),
        ("↔ Недостача", int(cap_summary.get("matched_shortage", 0)), INT),
        ("Unmatched name", int(cap_summary.get("unmatched_name", 0)), INT),
        ("Unmatched/uncertain store",
         int(cap_summary.get("unmatched_store", 0)) + int(cap_summary.get("uncertain_store", 0)), INT),
    ]
    hrow(ws, 5, [k[0] for k in kpis], HM, col_start=1)
    for i, (_, val, fmt) in enumerate(kpis, 1):
        nc(ws, 6, i, val, fmt, TF, bold=True)

    headers = [
        "Склад (файл опр.)", "Магазин (инвент.)", "Наименование", "Кол-во", "Сумма",
        "Статус", "Связь с недостачей", "Инв. излишек", "Инв. недостача", "Комментарий",
    ]
    hrow(ws, 8, headers, HD, col_start=1)
    if matched_df is None or matched_df.empty:
        tc(ws, 9, 1, "Нет данных оприходования.")
        return

    status_fill = {
        STATUS_MATCHED: SF,
        STATUS_MATCHED_SHORTAGE: DF,
        STATUS_UNMATCHED_NAME: PF,
        STATUS_UNMATCHED_STORE: PF,
        STATUS_UNCERTAIN_STORE: PatternFill("solid", fgColor="FCE4D6"),
    }
    r = 9
    for _, row in matched_df.iterrows():
        st = str(row.get("статус", ""))
        fl = status_fill.get(st, zebra_fill(r) or None)
        tc(ws, r, 1, row.get("склад", ""), fl, wrap=True)
        tc(ws, r, 2, row.get("магазин", "") or "—", fl, wrap=True)
        tc(ws, r, 3, row.get("наименование", ""), fl, wrap=True)
        nc(ws, r, 4, float(row.get("оп_количество", 0) or 0), NUM, fl)
        nc(ws, r, 5, float(row.get("оп_сумма", 0) or 0), SUM, fl)
        tc(ws, r, 6, st, fl, bold=True)
        link = str(row.get("связь_с_недостачей", "НЕТ"))
        link_fl = DF if link == "ДА" else fl
        tc(ws, r, 7, link, link_fl, bold=(link == "ДА"))
        nc(ws, r, 8, float(row.get("инв_излишек_сумма", 0) or 0), SUM, fl)
        nc(ws, r, 9, float(row.get("инв_недостача_сумма", 0) or 0), SUM, fl)
        tc(ws, r, 10, row.get("комментарий", ""), fl, wrap=True)
        ws.row_dimensions[r].height = 28
        r += 1
    trow(ws, r, 10, {1: f"ИТОГО ({len(matched_df)} строк)"},
         {4: (round(float(matched_df["оп_количество"].sum()), 3), NUM),
          5: (round(float(matched_df["оп_сумма"].sum()), 2), SUM)})
    dbars(ws, f"E9:E{r-1}", "63BE7B")
    ws.auto_filter.ref = f"A8:{GL(10)}{r-1}"
    ws.freeze_panes = "A9"


def sh_anomalies(wb, anom_df: pd.DataFrame, anom_summary: dict):
    """Sheet: book-sum anomalies (Release 4)."""
    ws = wb.create_sheet("Аномалии")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {
        1: 36, 2: 24, 3: 14, 4: 14, 5: 14, 6: 12, 7: 44, 8: 10, 9: 40,
    })
    mtitle(ws, 2, 1, 9, "АНОМАЛИИ КНИЖНОЙ НОРМАТИВНОЙ / ФАКТИЧЕСКОЙ СУММЫ",
           Font(name="Calibri", bold=True, size=13, color=WH), HD)
    tc(ws, 3, 1,
       "Пустые книжные суммы не скрываются: тип 1 — риск нулевой с/с; тип 2 — критично "
       "(нет расхождений по сумме); тип 3 — только фактическая, нужна сверка.")

    kpis = [
        ("Всего", int(anom_summary.get("total", 0)), INT),
        ("Критических", int(anom_summary.get("critical", 0)), INT),
        ("Тип 1", int(anom_summary.get("type1", 0)), INT),
        ("Тип 2", int(anom_summary.get("type2", 0)), INT),
        ("Тип 3", int(anom_summary.get("type3", 0)), INT),
        ("Магазинов", int(anom_summary.get("stores_affected", 0)), INT),
    ]
    hrow(ws, 5, [k[0] for k in kpis], HM, col_start=1)
    for i, (_, val, fmt) in enumerate(kpis, 1):
        fl = DF if i == 2 and val else TF
        nc(ws, 6, i, val, fmt, fl, bold=True)

    headers = [
        "Наименование", "Магазин / склад", "Тип аномалии",
        "Книжная норм. сумма", "Книжная факт. сумма",
        "Что произошло", "Критичность", "Комментарий", "Другие магазины",
    ]
    hrow(ws, 8, headers, HD, col_start=1)
    if anom_df is None or anom_df.empty:
        tc(ws, 9, 1, "Аномалий не обнаружено.")
        return

    type_fill = {
        TYPE_1: PF,
        TYPE_2: DF,
        TYPE_3: PatternFill("solid", fgColor="FCE4D6"),
    }
    r = 9
    for _, row in anom_df.iterrows():
        atype = str(row.get("тип_аномалии", ""))
        fl = type_fill.get(atype, zebra_fill(r) or None)
        tc(ws, r, 1, row.get("наименование", ""), fl, wrap=True)
        tc(ws, r, 2, row.get("магазин", ""), fl, wrap=True)
        tc(ws, r, 3, atype, fl, bold=True)
        bn = row.get("книжная_нормативная_сумма")
        bf = row.get("книжная_фактическая_сумма")
        if bn is None or (isinstance(bn, float) and np.isnan(bn)):
            tc(ws, r, 4, "— отсутствует —", fl)
        else:
            nc(ws, r, 4, float(bn), SUM, fl)
        if bf is None or (isinstance(bf, float) and np.isnan(bf)):
            tc(ws, r, 5, "— отсутствует —", fl)
        else:
            nc(ws, r, 5, float(bf), SUM, fl)
        tc(ws, r, 6, row.get("что_произошло", ""), fl, wrap=True)
        crit = str(row.get("критичность", "НЕТ"))
        tc(ws, r, 7, crit, DF if crit == "ДА" else fl, bold=(crit == "ДА"))
        tc(ws, r, 8, row.get("комментарий", ""), fl, wrap=True)
        tc(ws, r, 9, row.get("другие_магазины", "") or "—", fl, wrap=True)
        ws.row_dimensions[r].height = 36
        r += 1
    ws.auto_filter.ref = f"A8:{GL(9)}{r-1}"
    ws.freeze_panes = "A9"


def sh_authors_vs_operators(wb, role_stats: pd.DataFrame, store_stats: pd.DataFrame,
                            author_summary: dict):
    """Лист полной статистики ревизор / оператор."""
    ws = wb.create_sheet("Ревизоры vs Операторы")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {i: 16 for i in range(1, 14)})
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 28
    mtitle(ws, 2, 1, 12, "РЕВИЗОРЫ VS ОПЕРАТОРЫ",
           Font(name="Calibri", bold=True, size=13, color=WH), HD)
    tc(ws, 3, 1,
       "Ревизор = время документа ровно 08:00:00. Оператор = любое иное распознанное время. "
       "Финансовый рейтинг магазинов не пересчитывается — это интерпретационный разрез.")

    mtitle(ws, 5, 1, 10, "СВОДКА ПО СЕТИ", SUBT_F, SH, "left")
    hrow(ws, 6, [
        "Автор", "Документов", "Излишки", "Недостачи", "Перекрытие", "Сальдо",
        "Доля изл.%", "Доля недост.%", "Доля пересорт.%",
    ], HM, col_start=1)
    r = 7
    if role_stats is not None and not role_stats.empty:
        for _, row in role_stats.iterrows():
            fl = SF if row["автор"] == "ревизор" else (PF if row["автор"] == "оператор" else TF)
            tc(ws, r, 1, row["автор"], fl, bold=True)
            nc(ws, r, 2, int(row["документов"]), INT, fl)
            nc(ws, r, 3, float(row["излишки"]), SUM, fl)
            nc(ws, r, 4, float(row["недостачи"]), SUM, fl)
            nc(ws, r, 5, float(row["перекрытие"]), SUM, fl)
            nc(ws, r, 6, float(row["сальдо"]), SUM, fl)
            nc(ws, r, 7, float(row.get("доля_излишки_%", 0)) / 100, PCT, fl)
            nc(ws, r, 8, float(row.get("доля_недостачи_%", 0)) / 100, PCT, fl)
            nc(ws, r, 9, float(row.get("доля_перекрытие_%", 0)) / 100, PCT, fl)
            r += 1
    r += 1
    verdict = (
        f"По сети больший вклад: {author_summary.get('network_dominant')} "
        f"({author_summary.get('network_sign')}). "
        f"Магазинов с доминантой ревизора: {author_summary.get('stores_rev_dominant', 0)}; "
        f"оператора: {author_summary.get('stores_op_dominant', 0)}. "
        f"Цепочек пересортов: {author_summary.get('chains_count', 0)} "
        f"на {author_summary.get('chains_sum', 0):,.0f} руб."
    )
    ws.merge_cells(f"A{r}:I{r}")
    c = ws.cell(row=r, column=1, value=verdict)
    c.font = Font(name="Calibri", size=10, bold=True)
    c.alignment = Alignment(wrap_text=True)
    ws.row_dimensions[r].height = 36
    r += 2

    mtitle(ws, r, 1, 12, "СРАВНЕНИЕ ПО МАГАЗИНАМ", SUBT_F, SH, "left")
    r += 1
    hrow(ws, r, [
        "Магазин", "Рев. док.", "Оп. док.",
        "Рев. изл.", "Оп. изл.", "Рев. недост.", "Оп. недост.",
        "Рев. пересорт", "Оп. пересорт",
        "Доля недост. рев.%", "Доля недост. оп.%",
        "Доминанта", "Знак",
    ], HM, col_start=1)
    r += 1
    start = r
    if store_stats is not None and not store_stats.empty:
        for _, row in store_stats.iterrows():
            fl = SF if row.get("доминанта") == "ревизор" else (
                PF if row.get("доминанта") == "оператор" else None
            )
            tc(ws, r, 1, row["магазин"], fl)
            nc(ws, r, 2, int(row.get("рев_документов", 0)), INT, fl)
            nc(ws, r, 3, int(row.get("оп_документов", 0)), INT, fl)
            nc(ws, r, 4, float(row.get("рев_излишки", 0)), SUM, fl)
            nc(ws, r, 5, float(row.get("оп_излишки", 0)), SUM, fl)
            nc(ws, r, 6, float(row.get("рев_недостачи", 0)), SUM, fl)
            nc(ws, r, 7, float(row.get("оп_недостачи", 0)), SUM, fl)
            nc(ws, r, 8, float(row.get("рев_перекрытие", 0)), SUM, fl)
            nc(ws, r, 9, float(row.get("оп_перекрытие", 0)), SUM, fl)
            nc(ws, r, 10, float(row.get("доля_рев_недостачи_%", 0)) / 100, PCT, fl)
            nc(ws, r, 11, float(row.get("доля_оп_недостачи_%", 0)) / 100, PCT, fl)
            tc(ws, r, 12, row.get("доминанта", ""), fl, bold=True)
            tc(ws, r, 13, row.get("доминанта_знак", ""), fl)
            r += 1
    if r > start:
        ws.auto_filter.ref = f"A{start-1}:{GL(13)}{r-1}"
        ws.freeze_panes = f"A{start}"


def sh_author_chains(wb, chains_df: pd.DataFrame, author_summary: dict):
    """Лист цепочек пересортов ревизор ↔ оператор."""
    ws = wb.create_sheet("Цепочки пересортов")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {
        1: 16, 2: 36, 3: 18, 4: 34, 5: 12, 6: 12, 7: 12,
        8: 34, 9: 12, 10: 12, 11: 12, 12: 14, 13: 34, 14: 12,
    })
    mtitle(ws, 2, 1, 14, "ЦЕПОЧКИ ПЕРЕСОРТОВ РЕВИЗОР ↔ ОПЕРАТОР",
           Font(name="Calibri", bold=True, size=13, color=WH), HD)
    tc(ws, 3, 1,
       "Одна и та же позиция (sku_key) в одном магазине: сначала одна роль фиксирует недостачу/излишек, "
       "затем другая роль — противоположный знак. Инициатор = автор первого документа.")
    kpis = [
        ("Цепочек", int(author_summary.get("chains_count", 0)), INT),
        ("Сумма цепочек", float(author_summary.get("chains_sum", 0)), SUM),
    ]
    hrow(ws, 5, [k[0] for k in kpis], HM, col_start=1)
    for i, (_, val, fmt) in enumerate(kpis, 1):
        nc(ws, 6, i, val, fmt, TF, bold=True)

    headers = [
        "Магазин", "Товар", "Категория",
        "Док 1", "Дата 1", "Автор 1", "Знак 1", "Сумма 1",
        "Док 2", "Дата 2", "Автор 2", "Знак 2", "Сумма 2",
        "Сумма цепочки", "Тип", "Инициатор",
    ]
    # Fix column count - I had 14 widths but 16 headers. Adjust.
    hrow(ws, 8, headers, HD, col_start=1)
    if chains_df is None or chains_df.empty:
        tc(ws, 9, 1, "Цепочки пересортов между ревизором и оператором не обнаружены.")
        return

    r = 9
    top = chains_df.head(1000)
    for idx, (_, row) in enumerate(top.iterrows()):
        fl = SF if row.get("инициатор") == "ревизор" else PF
        tc(ws, r, 1, row.get("магазин", ""), fl)
        tc(ws, r, 2, row.get("наименование", ""), fl, wrap=True)
        tc(ws, r, 3, row.get("category_group", ""), fl)
        tc(ws, r, 4, row.get("док1", ""), fl, wrap=True)
        dt1 = row.get("дата1")
        tc(ws, r, 5, dt1.strftime("%d.%m.%Y %H:%M") if hasattr(dt1, "strftime") else str(dt1 or ""), fl)
        tc(ws, r, 6, row.get("автор1", ""), fl, bold=True)
        tc(ws, r, 7, row.get("знак1", ""), fl)
        nc(ws, r, 8, float(row.get("сумма1", 0) or 0), SUM, fl)
        tc(ws, r, 9, row.get("док2", ""), fl, wrap=True)
        dt2 = row.get("дата2")
        tc(ws, r, 10, dt2.strftime("%d.%m.%Y %H:%M") if hasattr(dt2, "strftime") else str(dt2 or ""), fl)
        tc(ws, r, 11, row.get("автор2", ""), fl, bold=True)
        tc(ws, r, 12, row.get("знак2", ""), fl)
        nc(ws, r, 13, float(row.get("сумма2", 0) or 0), SUM, fl)
        nc(ws, r, 14, float(row.get("сумма_цепочки", 0) or 0), SUM, fl)
        tc(ws, r, 15, row.get("тип_цепочки", ""), fl, wrap=True)
        tc(ws, r, 16, row.get("инициатор", ""), fl, bold=True)
        ws.row_dimensions[r].height = 32
        r += 1
    trow(ws, r, 16, {1: f"ИТОГО ({len(top)} цепочек)"},
         {14: (round(float(top["сумма_цепочки"].sum()), 2), SUM)})
    dbars(ws, f"N9:N{r-1}", ORG)
    ws.auto_filter.ref = f"A8:{GL(16)}{r-1}"
    ws.freeze_panes = "A9"


def sh_modern_analytics_sheet(wb, title, sheet_name, df_show, headers_map, used_names,
                              accent_fill=HR, bar_col_idx=None, highlight_col=None):
    """Современный аналитический лист: зебра, цветовая подсветка, полоски значений."""
    sname = safe_sheet_name(sheet_name, used_names)
    ws = wb.create_sheet(sname)
    ws.sheet_view.showGridLines = False
    cols = list(headers_map.keys())
    ncols = len(cols)
    for i, w in enumerate([14, 42, 12, 12, 14, 18, 14, 12, 12, 12, 12][:ncols], 1):
        ws.column_dimensions[GL(i)].width = w
    mtitle(ws, 2, 1, ncols, title, Font(name="Calibri", bold=True, size=13, color=WH), HD)
    hrow(ws, 4, [headers_map[c] for c in cols], accent_fill, col_start=1)
    if df_show is None or df_show.empty:
        tc(ws, 5, 1, "Нет данных.")
        return sname
    sum_cols = {j for j, c in enumerate(cols, 1)
                if any(k in c for k in ("сумма", "недост", "излиш", "сальдо", "перекрыт", "балл"))}
    r = 5
    for row_idx, (_, row) in enumerate(df_show.iterrows()):
        zf = zebra_fill(row_idx)
        for j, c in enumerate(cols, 1):
            val = row.get(c, "")
            if isinstance(val, (bool, np.bool_)):
                val = "ДА" if val else "НЕТ"
            elif c == "chronic":
                val = "Хроническая" if val else "—"
            elif c == "уровень" and val in LEVEL_RU:
                val = LEVEL_RU[val]
            elif c == "homogeneity" and val in HOMOGENEITY_RU:
                val = HOMOGENEITY_RU[val]
            elif c == "severity":
                val = str(val).replace("Critical", "Критический").replace("High", "Высокий").replace(
                    "Medium", "Средний").replace("Low", "Низкий")
            fl = zf
            if highlight_col and c == highlight_col:
                if isinstance(val, (int, float)) and val > 0:
                    fl = DF if "недост" in c else SF
            if isinstance(val, (int, float, np.integer, np.floating)) and not isinstance(val, bool):
                nc(ws, r, j, float(val), SUM if j in sum_cols else INT, fl)
            else:
                tc(ws, r, j, val, fl, wrap=(j <= 2))
        r += 1
    last = r - 1
    if bar_col_idx and last >= 5:
        dbars(ws, f"{GL(bar_col_idx)}5:{GL(bar_col_idx)}{last}",
              "F8696B" if accent_fill == HR else ("63BE7B" if accent_fill == HG else "4472C4"))
    ws.auto_filter.ref = f"A4:{GL(ncols)}{last}"
    ws.freeze_panes = "A5"
    return sname


def sh_generic_list(wb, title, sheet_name, df_show, headers_map, used_names):
    return sh_modern_analytics_sheet(wb, title, sheet_name, df_show, headers_map, used_names)


def sh_measures(wb, dd, dfp):
    ws = wb.create_sheet("Мероприятия")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {1: 4, 2: 6, 3: 38, 4: 12, 5: 8, 6: 12, 7: 14, 8: 10, 9: 22, 10: 55, 11: 8})
    pa = set(dfp["недостача_арт"].unique()) if dfp is not None and not dfp.empty else set()
    mtitle(ws, 2, 1, 11, "ПЛАН МЕРОПРИЯТИЙ ПО СЕТИ",
           Font(name="Calibri", bold=True, size=13, color=WH), HD)
    hrow(ws, 5, ["№", "Магазин", "Наименование", "Арт.", "Ед.", "Кол.", "Сумма",
                 "Сценарий", "Тип", "Мероприятия", "Приор."], ht=28, col_start=1)
    dd_s = dd.sort_values("недостача_сумма", ascending=False)
    r = 6
    for idx, (_, row) in enumerate(dd_s.iterrows(), 1):
        sc, sn, measures, priority = assign_scenario(row, pa, dfp)
        fl = PatternFill("solid", fgColor=SCENARIO_COLORS.get(sc, "FFFFFF"))
        pc = SCENARIO_PRIORITY_COLOR.get(priority, "000000")
        ws.row_dimensions[r].height = 60
        nc(ws, r, 1, idx, INT, fl)
        tc(ws, r, 2, row.get("магазин", ""), fl)
        tc(ws, r, 3, row["наименование"], fl, wrap=True)
        tc(ws, r, 4, row.get("артикул", ""), fl)
        tc(ws, r, 5, row.get("единица", ""), fl)
        nc(ws, r, 6, row["недостача_кол"], NUM, fl)
        nc(ws, r, 7, row["недостача_сумма"], SUM, fl)
        c_sc = ws.cell(row=r, column=8, value=sc)
        c_sc.font = Font(name="Calibri", size=10, bold=True, color=pc)
        c_sc.fill = fl
        c_sc.border = BD
        tc(ws, r, 9, sn, fl, wrap=True)
        c_m = ws.cell(row=r, column=10, value=measures)
        c_m.font = Font(name="Calibri", size=9)
        c_m.fill = fl
        c_m.border = BD
        c_m.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        ws.cell(row=r, column=11, value=f"P{priority}").font = Font(bold=True, color=pc)
        r += 1
    ws.auto_filter.ref = f"A5:{GL(11)}{r-1}"
    ws.freeze_panes = "A6"
    dbars(ws, f"G6:G{r-1}", "F8696B")


def sh_overlap(wb, overlap_df, title_suffix="", sheet_title="Перекрытие"):
    ws = wb.create_sheet(sheet_title)
    ws.sheet_view.showGridLines = False
    mtitle(ws, 2, 1, 14, f"АНАЛИЗ ПЕРЕКРЫТИЯ {title_suffix}",
           Font(name="Calibri", bold=True, size=13, color=WH), HD)
    if "АНАЛИТИК" in title_suffix.upper() or sheet_title == "Перекрытие сеть":
        tc(ws, 3, 1, "Аналитический расчёт между магазинами. Не заменяет физический пересорт без перемещения.")
    if overlap_df.empty:
        tc(ws, 4, 1, "Перекрытие не обнаружено.")
        return
    cols = ["уровень", "магазин_изл", "док_изл", "автор_изл", "магазин_нед", "док_нед", "автор_нед",
            "излишек_товар", "недостача_товар", "category_group", "перекрытие_сум",
            "score", "homogeneity", "fallback_flag"]
    headers = ["Уровень", "Маг. изл.", "Док. изл.", "Автор изл.", "Маг. нед.", "Док. нед.", "Автор нед.",
               "Излишек", "Недостача", "Категория", "Сумма", "Балл", "Однородность", "Резерв. кат."]
    hrow(ws, 4, headers, col_start=1)
    r = 5
    top = overlap_df.head(500)
    for row_idx, (_, row) in enumerate(top.iterrows()):
        fb = row.get("fallback_flag") == "ДА"
        fl = PF if fb else (SF if row.get("уровень") == "cross_doc" else (zebra_fill(row_idx) or TF))
        for j, c in enumerate(cols, 1):
            val = row.get(c, "")
            if c == "уровень" and val in LEVEL_RU:
                val = LEVEL_RU[val]
            elif c == "homogeneity" and val in HOMOGENEITY_RU:
                val = HOMOGENEITY_RU[val]
            if c == "перекрытие_сум":
                nc(ws, r, j, val, SUM, fl)
            elif c == "score":
                nc(ws, r, j, val, INT, fl)
            else:
                tc(ws, r, j, val, fl, wrap=(j in (8, 9)))
        r += 1
    trow(ws, r, 14, {1: f"ИТОГО ({len(top)} пар)"}, {11: (round(top["перекрытие_сум"].sum(), 2), SUM)})
    dbars(ws, f"K5:K{r-1}", ORG)
    ws.auto_filter.ref = f"A4:{GL(14)}{r-1}"
    ws.freeze_panes = "A5"


def sh_store_detail(wb, store_name, store_df, doc_metrics, used_names):
    sname = safe_sheet_name(f"Маг_{store_name}"[:31], used_names)
    ws = wb.create_sheet(sname)
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {1: 6, 2: 36, 3: 14, 4: 12, 5: 14, 6: 14, 7: 14, 8: 12, 9: 12})
    sur = store_df["излишек_сумма"].sum()
    sh = store_df["недостача_сумма"].sum()
    rev_docs = 0
    op_docs = 0
    if "автор_дока" in store_df.columns:
        rev_docs = store_df.loc[store_df["автор_дока"] == "ревизор", "документ"].nunique()
        op_docs = store_df.loc[store_df["автор_дока"] == "оператор", "документ"].nunique()
    mtitle(ws, 2, 1, 9, f"МАГАЗИН: {store_name}",
           Font(name="Calibri", bold=True, size=13, color=WH), HD)
    mtitle(ws, 3, 1, 9,
           f"Излишки: {sur:,.0f} | Недостачи: {sh:,.0f} | Сальдо: {sur-sh:,.0f} | "
           f"Док. ревизора: {rev_docs} | Док. оператора: {op_docs}",
           Font(name="Calibri", size=9, color="606060"))
    r = 5
    mtitle(ws, r, 1, 9, "РЕЙТИНГ ДОКУМЕНТОВ", SUBT_F, SH, "left")
    r += 1
    hrow(ws, r, ["№", "Документ", "Дата", "Автор", "Класс", "Недостачи", "Излишки", "Перекрытие", "Чист. недост."], col_start=1)
    r += 1
    dm = doc_metrics[doc_metrics["магазин"] == store_name].head(30)
    for _, drow in dm.iterrows():
        fl = DF if drow["класс"] in ("Критический", "Высокий") else (SF if drow["класс"] == "Излишки" else None)
        nc(ws, r, 1, int(drow["рейтинг"]), INT, fl)
        tc(ws, r, 2, drow["документ"], fl, wrap=True)
        dt = drow["дата"]
        tc(ws, r, 3, dt.strftime("%d.%m.%Y %H:%M") if hasattr(dt, "strftime") else str(dt), fl)
        tc(ws, r, 4, drow.get("автор", "неизвестно"), fl, bold=True)
        tc(ws, r, 5, drow["класс"], fl, bold=True)
        nc(ws, r, 6, drow["недостачи"], SUM, fl)
        nc(ws, r, 7, drow["излишки"], SUM, fl)
        nc(ws, r, 8, drow["перекрытие"], SUM, fl)
        nc(ws, r, 9, drow["чистые_недостачи"], SUM, fl)
        r += 1
    r += 2
    mtitle(ws, r, 1, 4, "ДОКУМЕНТЫ С НЕДОСТАЧАМИ", Font(name="Calibri", bold=True, size=11, color=WH), HR)
    r += 1
    hrow(ws, r, ["Документ", "Автор", "Недостачи", "Поз. с недост."], HR, col_start=1)
    r += 1
    for _, drow in dm[dm["недостачи"] > 0].head(15).iterrows():
        tc(ws, r, 1, drow["документ"], DF, wrap=True)
        tc(ws, r, 2, drow.get("автор", ""), DF)
        nc(ws, r, 3, drow["недостачи"], SUM, DF)
        nc(ws, r, 4, int(drow["sku_нед"]), INT, DF)
        r += 1
    r += 2
    mtitle(ws, r, 1, 4, "ДОКУМЕНТЫ С ИЗЛИШКАМИ", Font(name="Calibri", bold=True, size=11, color=WH), HG)
    r += 1
    hrow(ws, r, ["Документ", "Автор", "Излишки", "Поз. с изл."], HG, col_start=1)
    r += 1
    for _, drow in dm[dm["излишки"] > 0].head(15).iterrows():
        tc(ws, r, 1, drow["документ"], SF, wrap=True)
        tc(ws, r, 2, drow.get("автор", ""), SF)
        nc(ws, r, 3, drow["излишки"], SUM, SF)
        nc(ws, r, 4, int(drow["sku_изл"]), INT, SF)
        r += 1
    r += 2
    mtitle(ws, r, 1, 7, "ТОП-10 НЕДОСТАЧ", Font(name="Calibri", bold=True, size=11, color=WH), HR)
    r += 1
    hrow(ws, r, ["№", "Наименование", "Арт.", "Кол.", "Сумма", "Категория"], HR, col_start=1)
    r += 1
    top_sh = store_df[store_df["недостача_сумма"] > 0].sort_values("недостача_сумма", ascending=False).head(TOP_N_STORE)
    for i, (_, srow) in enumerate(top_sh.iterrows(), 1):
        nc(ws, r, 1, i, INT, DF)
        tc(ws, r, 2, srow["наименование"], DF, wrap=True)
        tc(ws, r, 3, srow.get("артикул", ""), DF)
        nc(ws, r, 4, srow["недостача_кол"], NUM, DF)
        nc(ws, r, 5, srow["недостача_сумма"], SUM, DF)
        tc(ws, r, 6, srow.get("category_label") or _display_category_value(srow.get("category_group", "")), DF)
        r += 1
    dbars(ws, f"E{r - len(top_sh)}:E{r - 1}", "F8696B")
    r += 2
    mtitle(ws, r, 1, 7, "ТОП-10 ИЗЛИШКОВ", Font(name="Calibri", bold=True, size=11, color=WH), HG)
    r += 1
    hrow(ws, r, ["№", "Наименование", "Арт.", "Кол.", "Сумма", "Категория"], HG, col_start=1)
    r += 1
    top_sur = store_df[store_df["излишек_сумма"] > 0].sort_values("излишек_сумма", ascending=False).head(TOP_N_STORE)
    sur_start = r
    for i, (_, srow) in enumerate(top_sur.iterrows(), 1):
        nc(ws, r, 1, i, INT, SF)
        tc(ws, r, 2, srow["наименование"], SF, wrap=True)
        tc(ws, r, 3, srow.get("артикул", ""), SF)
        nc(ws, r, 4, srow["излишек_кол"], NUM, SF)
        nc(ws, r, 5, srow["излишек_сумма"], SUM, SF)
        tc(ws, r, 6, srow.get("category_label") or _display_category_value(srow.get("category_group", "")), SF)
        r += 1
    if r > sur_start:
        dbars(ws, f"E{sur_start}:E{r - 1}", "63BE7B")
    ws.freeze_panes = "A6"
    return sname


def sh_conclusions(wb, conclusions):
    ws = wb.create_sheet("Выводы")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 100
    mtitle(ws, 2, 1, 2, "ВЫВОДЫ И РЕКОМЕНДАЦИИ", Font(name="Calibri", bold=True, size=13, color=WH), HD)
    r = 4
    for title, text in conclusions:
        mtitle(ws, r, 1, 2, title, SUBT_F, SH, "left")
        r += 1
        ws.merge_cells(f"A{r}:B{r}")
        c = ws.cell(row=r, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
        c.font = Font(name="Calibri", size=10)
        ws.row_dimensions[r].height = 55
        r += 2


def sh_pivot_heatmap(wb, pivot_df, used_names):
    sname = safe_sheet_name("Позиции x Магазины", used_names)
    ws = wb.create_sheet(sname)
    ws.sheet_view.showGridLines = False
    top = pivot_df[pivot_df.sum(axis=1) > 0].sum(axis=1).sort_values(ascending=False).head(100).index
    sub = pivot_df.loc[top]
    ncols = min(14, len(sub.columns) + 1)
    mtitle(ws, 2, 1, ncols,
           "СРАВНЕНИЕ ПОЗИЦИЙ — НЕДОСТАЧИ ПО МАГАЗИНАМ (ТОП-100)",
           Font(name="Calibri", bold=True, size=12, color=WH), HD)
    cols = ["Наименование"] + list(sub.columns[:ncols - 1])
    hrow(ws, 4, cols, col_start=1)
    r = 5
    for idx, row in sub.iterrows():
        tc(ws, r, 1, str(idx)[:60], wrap=True)
        for j, col in enumerate(sub.columns[:ncols - 1], 2):
            nc(ws, r, j, row[col], SUM)
        r += 1
    if r > 5:
        cs3(ws, f"B5:{GL(ncols)}{r-1}")
    ws.freeze_panes = "B5"
    return sname


def build_report(cfg: Config) -> str:
    print(f"\n[1/8] Читаю файл: {cfg.input_path}")
    df, meta = parse_network_excel(cfg.input_path)
    if "автор_дока" not in df.columns:
        from src.author import enrich_author_columns
        df = enrich_author_columns(df)
    print(f"      Лист: {meta.sheet_name} | Магазинов: {len(meta.stores)} | "
          f"Документов: {meta.doc_count} | Позиций: {meta.sku_count}")
    if "автор_дока" in df.columns:
        rev_n = int(df.loc[df["автор_дока"] == "ревизор", "документ"].nunique())
        op_n = int(df.loc[df["автор_дока"] == "оператор", "документ"].nunique())
        unk_n = int(df.loc[df["автор_дока"] == "неизвестно", "документ"].nunique())
        print(f"      Авторы документов: ревизор={rev_n}, оператор={op_n}, неизвестно={unk_n}")
    if meta.stores:
        print("      Пример магазинов:", ", ".join(meta.stores[:8]))

    print(f"[2/8] Фильтр периода ({cfg.period_days} дней)...")
    end = cfg.end_date or meta.date_max
    df = filter_by_period(df, cfg.period_days, end)
    if df.empty:
        raise ValueError("После фильтрации по периоду не осталось данных.")
    p_start = df.attrs.get("period_start", meta.date_min)
    p_end = df.attrs.get("period_end", meta.date_max)
    period_str = f"{p_start.strftime('%d.%m.%Y')} — {p_end.strftime('%d.%m.%Y')}"
    print(f"      Период: {period_str} | Позиций в периоде: {len(df):,}")

    print("[3/8] Обогащение данных (эталонная иерархия номенклатуры)...")
    master = get_master_hierarchy()
    catalog = load_catalog(cfg.catalog_path)
    df = enrich_dataframe(df, catalog, hierarchy=master)
    unmatched_cnt = int((~df["hierarchy_matched"]).sum()) if "hierarchy_matched" in df.columns else 0
    unmatched_sku = unmatched_summary(df)
    print(
        f"      Справочник: {master.source_file} | SKU в эталоне: {len(master.items):,} | "
        f"не найдено строк: {unmatched_cnt:,} ({len(unmatched_sku):,} уник. наим.)"
    )
    if unmatched_cnt:
        print("      Позиции вне справочника попадут на лист «Вне справочника» (без выдуманных категорий).")

    print("[4/8] Расчёт перекрытий и метрик...")
    overlap_op = build_operational_overlaps(df)
    overlap_op = annotate_overlap_authors(overlap_op, df)
    overlap_cross = build_analytical_cross_store(df) if cfg.enable_cross_store else pd.DataFrame()
    if not overlap_cross.empty:
        overlap_cross = annotate_overlap_authors(overlap_cross, df)

    store_metrics = calc_store_metrics(df, overlap_op)
    doc_metrics = calc_document_metrics(df)
    sku_cross = calc_sku_cross(df)
    chronic = sku_cross[sku_cross["chronic"]]
    summary = calc_network_summary(df, overlap_op)
    pivot = build_sku_store_pivot(df)

    print("[5/8] Аномалии книжных сумм и разрез ревизор/оператор...")
    anom_df = detect_book_sum_anomalies(df)
    anom_sum = anomalies_summary(anom_df, total_rows=len(df))
    store_metrics = enrich_store_metrics_with_anomalies(store_metrics, anom_df)

    author_role = calc_author_role_stats(df, overlap_op)
    author_store = calc_author_store_stats(df, overlap_op)
    author_chains = detect_author_chains(df)
    author_sum = author_network_summary(author_role, author_store, author_chains)
    store_metrics = enrich_store_metrics_with_authors(store_metrics, author_store)
    print(
        f"      Аномалий: {anom_sum['total']} (критич. {anom_sum['critical']}) | "
        f"док. ревизора: {author_sum['rev_docs']} | док. оператора: {author_sum['op_docs']} | "
        f"цепочек: {author_sum['chains_count']}"
    )

    print("[6/8] Оприходование излишков (закрытие смены)...")
    cap_matched = pd.DataFrame()
    cap_summary = None
    if cfg.capitalization_path:
        cap_raw = parse_capitalization_excel(cfg.capitalization_path)
        cap_matched = match_capitalization_to_inventory(cap_raw, df)
        cap_summary = capitalization_summary(cap_matched)
        store_metrics = enrich_store_metrics_with_capitalization(store_metrics, cap_matched)
        print(
            f"      Складов: {cap_summary['warehouses']} | строк: {cap_summary['cap_rows']} | "
            f"сумма: {cap_summary['cap_sum']:,.0f} | matched: "
            f"{cap_summary['matched'] + cap_summary['matched_shortage']} | "
            f"↔недостача: {cap_summary['matched_shortage']} | "
            f"unmatched: {cap_summary['unmatched_name'] + cap_summary['unmatched_store'] + cap_summary['uncertain_store']}"
        )
    else:
        print("      Файл оприходования не передан — лист сверки будет пустым.")

    dd = df[df["недостача_кол"] > 0].copy()
    ds = df[df["излишек_кол"] > 0].copy()
    pa = set(overlap_op["недостача_арт"].unique()) if not overlap_op.empty else set()
    dfn = dd[~dd["артикул"].isin(pa)].copy()
    conclusions = generate_conclusions(
        summary, store_metrics, chronic,
        cap_summary=cap_summary,
        anom_summary=anom_sum,
        author_summary=author_sum,
    )

    print("[7/8] Формирование Excel...")
    wb = Workbook()
    wb.remove(wb.active)
    used_names: set = set()
    sheets_info: List[Tuple[str, str, str]] = []

    sh_dashboard(
        wb, summary, store_metrics, period_str,
        cap_summary=cap_summary,
        anom_summary=anom_sum,
        author_summary=author_sum,
    )
    sheets_info.append(("Сводка", "Ключевые показатели, аномалии, ревизоры/операторы", "Сводка!A2"))

    sh_conclusions(wb, conclusions)
    sheets_info.append(("Выводы", "Рекомендации и итоги", "Выводы!A2"))

    sh_store_ranking(wb, store_metrics)
    sheets_info.append(("Рейтинг магазинов", "Сводный рейтинг классов A/B/C + вклад ролей", "Рейтинг магазинов!A2"))

    sh_anomalies(wb, anom_df, anom_sum)
    sheets_info.append(("Аномалии", "Книжные нормативные/фактические суммы", "Аномалии!A2"))

    sh_authors_vs_operators(wb, author_role, author_store, author_sum)
    sheets_info.append(("Ревизоры vs Операторы", "Вклад ревизоров и операторов", "Ревизоры vs Операторы!A2"))

    sh_author_chains(wb, author_chains, author_sum)
    sheets_info.append(("Цепочки пересортов", "Цепочки ревизор↔оператор по SKU", "Цепочки пересортов!A2"))

    if cap_summary is not None:
        sh_capitalization(wb, cap_matched, cap_summary)
        sheets_info.append(
            ("Оприходование излишков", "Сверка оприходования закрытия смены", "Оприходование излишков!A2")
        )

    sh_modern_analytics_sheet(
        wb, "ТОП НЕДОСТАЧ ПО СЕТИ", "Топ недостач",
        _for_sheet(sku_cross[sku_cross["недостачи"] > 0].head(TOP_N_NETWORK)),
        {"наименование": "Наименование", "category_group": "Категория (лист)", "hierarchy_path": "Иерархия",
         "магазинов": "Магазинов", "документов": "Документов", "недостачи": "Сумма", "излишки": "Излишки",
         "сальдо": "Сальдо"},
        used_names, accent_fill=HR, bar_col_idx=6)
    sheets_info.append(("Топ недостач", "Топ-50 недостач по сети", "Топ недостач!A2"))

    if not overlap_op.empty:
        peresor_top = _for_sheet(overlap_op.sort_values("перекрытие_сум", ascending=False).head(TOP_N_NETWORK))
    else:
        peresor_top = pd.DataFrame()
    sh_modern_analytics_sheet(
        wb, "ТОП ПЕРЕСОРТОВ / ПЕРЕКРЫТИЙ", "Топ пересортов", peresor_top,
        {"уровень": "Уровень", "излишек_товар": "Излишек", "недостача_товар": "Недостача",
         "category_group": "Категория (лист)", "перекрытие_сум": "Сумма", "score": "Балл",
         "fallback_flag": "Вне справочника"},
        used_names, accent_fill=PF, bar_col_idx=5)
    sheets_info.append(("Топ пересортов", "Топ перекрытий однородными парами", "Топ пересортов!A2"))

    sh_modern_analytics_sheet(
        wb, "ХРОНИЧЕСКИЕ ПРОБЛЕМНЫЕ ПОЗИЦИИ", "Хронические позиции",
        _for_sheet(chronic.head(TOP_N_NETWORK)),
        {"наименование": "Наименование", "category_group": "Категория (лист)", "hierarchy_path": "Иерархия",
         "магазинов": "Магазинов", "документов": "Документов", "недостачи": "Недостачи", "излишки": "Излишки"},
        used_names, accent_fill=HR, bar_col_idx=6)
    sheets_info.append(("Хронические позиции", "Повторяющиеся проблемы", "Хронические позиции!A2"))

    sh_pivot_heatmap(wb, pivot, used_names)
    sheets_info.append(("Позиции x Магазины", "Тепловая карта недостач", "Позиции x Магазины!A2"))

    sh_overlap(wb, _for_sheet(overlap_op), "(ОПЕРАЦИОННОЕ — внутри магазина)", "Перекрытие")
    sheets_info.append(("Перекрытие", "Перекрытие между документами", "Перекрытие!A2"))
    if not overlap_cross.empty:
        sh_overlap(wb, _for_sheet(overlap_cross), "(АНАЛИТИКА — между магазинами)", "Перекрытие сеть")
        sheets_info.append(("Перекрытие сеть", "Аналитика между магазинами", "Перекрытие сеть!A2"))

    sh_measures(wb, dd, overlap_op)
    sheets_info.append(("Мероприятия", "План действий", "Мероприятия!A2"))

    shortage_map = {
        "магазин": "Магазин", "документ": "Документ", "автор_дока": "Автор",
        "наименование": "Наименование", "артикул": "Артикул",
        "недостача_кол": "Количество", "недостача_сумма": "Сумма",
        "category_group": "Категория (лист)", "hierarchy_path": "Иерархия",
    }
    sh_modern_analytics_sheet(
        wb, "ПОЛНЫЙ ПЕРЕЧЕНЬ НЕДОСТАЧ", "Недостачи",
        _for_sheet(dd.sort_values("недостача_сумма", ascending=False)),
        shortage_map, used_names, accent_fill=HR, bar_col_idx=7)
    sheets_info.append(("Недостачи", "Все недостачи", "Недостачи!A2"))

    surplus_map = {
        "магазин": "Магазин", "документ": "Документ", "автор_дока": "Автор",
        "наименование": "Наименование", "артикул": "Артикул",
        "излишек_кол": "Количество", "излишек_сумма": "Сумма",
        "category_group": "Категория (лист)", "hierarchy_path": "Иерархия",
    }
    sh_modern_analytics_sheet(
        wb, "ПОЛНЫЙ ПЕРЕЧЕНЬ ИЗЛИШКОВ", "Излишки",
        _for_sheet(ds.sort_values("излишек_сумма", ascending=False)),
        surplus_map, used_names, accent_fill=HG, bar_col_idx=7)
    sheets_info.append(("Излишки", "Все излишки", "Излишки!A2"))

    sh_modern_analytics_sheet(
        wb, "ЧИСТЫЕ НЕДОСТАЧИ", "Чистые недостачи",
        _for_sheet(dfn.sort_values("недостача_сумма", ascending=False)),
        {"магазин": "Магазин", "документ": "Документ", "автор_дока": "Автор",
         "наименование": "Наименование", "недостача_кол": "Количество",
         "недостача_сумма": "Сумма", "category_group": "Категория (лист)", "hierarchy_path": "Иерархия"},
        used_names, accent_fill=HR, bar_col_idx=6)
    sheets_info.append(("Чистые недостачи", "Недостачи без однородного перекрытия", "Чистые недостачи!A2"))

    sh_modern_analytics_sheet(
        wb, "ПОЗИЦИИ — СКВОЗНОЙ АНАЛИЗ", "Анализ позиций", _for_sheet(sku_cross.head(500)),
        {"наименование": "Наименование", "category_group": "Категория (лист)", "hierarchy_path": "Иерархия",
         "магазинов": "Магазинов", "документов": "Документов", "излишки": "Излишки", "недостачи": "Недостачи",
         "сальдо": "Сальдо", "chronic": "Хроничность"},
        used_names, accent_fill=HM, bar_col_idx=7)
    sheets_info.append(("Анализ позиций", "Сквозной анализ по позициям", "Анализ позиций!A2"))

    sh_modern_analytics_sheet(
        wb, "ВНЕ СПРАВОЧНИКА", "Вне справочника", unmatched_sku.head(2000),
        {"наименование": "Наименование", "sku_key": "Ключ", "строк": "Строк в выгрузке"},
        used_names, accent_fill=PF, bar_col_idx=3)
    sheets_info.append(("Вне справочника", "SKU без match в эталонной иерархии", "Вне справочника!A2"))

    sh_modern_analytics_sheet(
        wb, "РЕЙТИНГ ДОКУМЕНТОВ", "Документы", doc_metrics.head(500),
        {"магазин": "Магазин", "документ": "Документ", "автор": "Автор", "класс": "Класс",
         "недостачи": "Недостачи", "излишки": "Излишки", "перекрытие": "Перекрытие",
         "чистые_недостачи": "Чист. недост.", "severity": "Критичность"},
        used_names, accent_fill=HM, bar_col_idx=9)
    sheets_info.append(("Документы", "Рейтинг документов", "Документы!A2"))

    for store in sorted(df["магазин"].unique()):
        sname = sh_store_detail(wb, store, df[df["магазин"] == store], doc_metrics, used_names)
        sheets_info.append((sname, f"Детализация: {store}", f"{sname}!A2"))

    ws_meta = wb.create_sheet("Метаданные")
    tc(ws_meta, 2, 1, f"Исходный файл: {meta.file_name}")
    tc(ws_meta, 3, 1, f"Период анализа: {period_str}")
    tc(ws_meta, 4, 1, f"Эталон иерархии: {master.source_file} ({len(master.items):,} SKU)")
    tc(ws_meta, 5, 1, f"Строк вне справочника: {unmatched_cnt:,} | уник. наименований: {len(unmatched_sku):,}")
    tc(ws_meta, 6, 1, f"Операционное перекрытие: {summary['overlap']:,.2f} руб.")
    tc(ws_meta, 7, 1, f"Чистые недостачи: {summary['clean_shortage']:,.2f} руб.")
    tc(ws_meta, 8, 1, "Категории берутся только из эталонного шаблона иерархии (без keyword-эвристик).")
    tc(ws_meta, 9, 1,
       f"Аномалии книжных сумм: {anom_sum['total']} | критических: {anom_sum['critical']} | "
       f"тип1={anom_sum['type1']} тип2={anom_sum['type2']} тип3={anom_sum['type3']}")
    tc(ws_meta, 10, 1,
       "Автор документа: ровно 08:00:00 → ревизор; любое иное распознанное время → оператор. "
       "Проверка цен в розничном модуле не выполняется.")
    tc(ws_meta, 11, 1,
       f"Документов ревизора: {author_sum['rev_docs']} | оператора: {author_sum['op_docs']} | "
       f"неизвестно: {author_sum.get('unk_docs', 0)} | "
       f"цепочек пересортов: {author_sum['chains_count']} ({author_sum['chains_sum']:,.2f} руб.)")
    if cap_summary is not None:
        tc(ws_meta, 12, 1,
           f"Оприходование: файл загружен | сумма {cap_summary['cap_sum']:,.2f} | "
           f"matched {cap_summary['matched'] + cap_summary['matched_shortage']} | "
           f"↔недостача {cap_summary['matched_shortage']}")
    else:
        tc(ws_meta, 12, 1, "Оприходование: файл не загружен")
    tc(ws_meta, 13, 1,
       f"Розничный модуль Release 2 (AI-агент Cursor), v{__import__('src').__version__}: "
       "аномалии + оприходование + ревизор/оператор + эталон иерархии без производства; "
       "базовые расчёты overlap/рейтинга сохранены; проверка цен отключена. "
       "Первый релиз — монолит v3.0 (анализ_сеть_последняя_с_группами).")
    sheets_info.append(("Метаданные", "Служебная информация", "Метаданные!A2"))

    sh_contents(wb, sheets_info)

    print(f"[8/8] Сохранение: {cfg.output_path}")
    wb.save(cfg.output_path)
    print("\nГОТОВО!")
    return cfg.output_path


def export_report_bytes(cfg: Config) -> bytes:
    """Build report and return XLSX bytes (for Streamlit download)."""
    # Use a temp file so openpyxl save path works cross-platform
    if not cfg.output_path:
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        cfg.output_path = tmp.name
        tmp.close()
    path = build_report(cfg)
    return Path(path).read_bytes()
