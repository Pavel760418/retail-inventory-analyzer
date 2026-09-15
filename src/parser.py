# -*- coding: utf-8 -*-
"""Parse network inventory Excel exports (TDSheet hierarchy)."""
from __future__ import annotations

import datetime
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from openpyxl import load_workbook

from config.settings import EXCLUDE_STORE_KEYWORDS
from src.author import classify_author, parse_time_flexible
from src.models import ParseMeta

DOC_RE = re.compile(
    r"^Инвентаризация\s+(\S+)\s+от\s+(\d{2}\.\d{2}\.\d{4})(?:\s+(\d{1,2}:\d{2}(?::\d{2})?))?",
    re.IGNORECASE,
)

REQUIRED_MARKERS = (
    "документ.подразделение", "количество книжное", "количество фактическое",
    "сумма излишек", "сумма недостача",
)


def validate_network_structure(df_raw: pd.DataFrame) -> None:
    first = " ".join(str(v).lower() for v in df_raw.iloc[:5, :9].fillna("").values.flatten())
    if any(k in first for k in ("сводный анализ", "анализ пересортицы", "план мероприятий")):
        raise ValueError(
            "Похоже, выбран уже готовый аналитический файл, а не исходная выгрузка инвентаризаций."
        )
    missing = [m for m in REQUIRED_MARKERS if m not in first]
    if len(missing) >= 3:
        raise ValueError(
            "Файл не похож на сетевую выгрузку TDSheet. "
            f"Не найдены ключевые заголовки: {', '.join(missing)}"
        )


def _to_float(v) -> float:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _is_missing_cell(v) -> bool:
    """True when Excel cell is empty (None / NaN / blank) — distinct from numeric zero."""
    if v is None:
        return True
    if isinstance(v, float) and np.isnan(v):
        return True
    if isinstance(v, str) and not str(v).strip():
        return True
    return False


def _read_outline_levels(filepath: str, sheet_name: str) -> Dict[int, int]:
    """
    Возвращает карту row_index(1-based) -> outlineLevel из XLSX.
    Нужна для корректного выделения иерархии Подразделение -> Документ -> Номенклатура.
    """
    wb = load_workbook(filepath, read_only=False, data_only=True)
    ws = wb[sheet_name]
    levels: Dict[int, int] = {}
    for row_idx, dim in ws.row_dimensions.items():
        try:
            lv = int(getattr(dim, "outlineLevel", 0) or 0)
        except Exception:
            lv = 0
        levels[int(row_idx)] = lv
    wb.close()
    return levels


def parse_network_excel(filepath: str, sheet_name: Optional[str] = None) -> Tuple[pd.DataFrame, ParseMeta]:
    path = Path(filepath)
    xl = pd.ExcelFile(filepath)
    sheet = sheet_name or (xl.sheet_names[0] if xl.sheet_names else None)
    if not sheet:
        raise ValueError("В файле нет листов.")
    df_raw = pd.read_excel(filepath, sheet_name=sheet, header=None)
    validate_network_structure(df_raw)
    outline_levels = _read_outline_levels(filepath, sheet)

    meta = ParseMeta(file_name=path.name, sheet_name=sheet)
    rows: List[Dict[str, Any]] = []
    current_store = ""
    current_doc = ""
    current_doc_num = ""
    current_doc_dt: Optional[datetime.datetime] = None
    current_doc_author: str = "неизвестно"
    current_doc_time: Optional[datetime.time] = None
    network_total: Optional[Dict[str, Any]] = None

    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        if i < 3:
            continue
        a_val = row.get(0)
        a = str(a_val).strip() if pd.notna(a_val) else ""
        if not a:
            continue

        vals = {k: _to_float(row.get(k)) for k in range(1, 9)}
        # Missing flags for book sums (cols 4/5) — do NOT change numeric calc above.
        miss_book_norm = _is_missing_cell(row.get(4))
        miss_book_fact = _is_missing_cell(row.get(5))
        # i в pandas 0-based, в Excel row 1-based
        outline = int(outline_levels.get(i + 1, 0))

        if i == 3 and not a.lower().startswith("инвентаризация"):
            network_total = {"магазин": a, **{f"c{k}": vals[k] for k in range(1, 9)}}
            if "итого" not in a.lower():
                current_store = a
                if a not in meta.stores:
                    meta.stores.append(a)
            continue

        m = DOC_RE.match(a)
        if m:
            current_doc = a
            current_doc_num = m.group(1)
            d = datetime.datetime.strptime(m.group(2), "%d.%m.%Y").date()
            if m.group(3):
                t = parse_time_flexible(m.group(3))
                if t is None:
                    # fallback: try zero-padded
                    try:
                        t = datetime.datetime.strptime(m.group(3), "%H:%M:%S").time()
                    except ValueError:
                        try:
                            t = datetime.datetime.strptime(m.group(3), "%H:%M").time()
                        except ValueError:
                            t = None
                if t is not None:
                    current_doc_dt = datetime.datetime.combine(d, t)
                    current_doc_time = t
                    current_doc_author = classify_author(time_only=t)
                else:
                    current_doc_dt = datetime.datetime.combine(d, datetime.time())
                    current_doc_time = None
                    current_doc_author = "неизвестно"
            else:
                current_doc_dt = datetime.datetime.combine(d, datetime.time())
                current_doc_time = None
                current_doc_author = "неизвестно"
            meta.doc_count += 1
            if meta.date_min is None or d < meta.date_min:
                meta.date_min = d
            if meta.date_max is None or d > meta.date_max:
                meta.date_max = d
            continue

        if outline == 0 and not m:
            low = a.lower()
            if any(x in low for x in EXCLUDE_STORE_KEYWORDS):
                continue
            current_store = a
            if a not in meta.stores:
                meta.stores.append(a)
            current_doc = ""
            current_doc_num = ""
            current_doc_dt = None
            current_doc_author = "неизвестно"
            current_doc_time = None
            continue

        if not current_store or not current_doc:
            continue
        if outline not in (1, 2):
            continue

        qty_diff = vals[3]
        surplus_qty = max(qty_diff, 0)
        shortage_qty = abs(min(qty_diff, 0))
        surplus_sum = max(vals[7], 0)
        shortage_sum = abs(vals[8]) if vals[8] != 0 else 0
        if surplus_sum == 0 and surplus_qty > 0 and vals[6] > 0:
            surplus_sum = vals[6]
        if shortage_sum == 0 and shortage_qty > 0 and vals[6] < 0:
            shortage_sum = abs(vals[6])

        rows.append({
            "магазин": current_store,
            "документ": current_doc,
            "номер_док": current_doc_num,
            "дата_док": current_doc_dt,
            "автор_дока": current_doc_author,
            "время_дока": current_doc_time,
            "наименование": a,
            "кол_учетное": vals[1],
            "кол_факт": vals[2],
            "кол_разница": qty_diff,
            "сумма_учетная": vals[4],
            "сумма_факт": vals[5],
            "сумма_разница": vals[6],
            "излишек_кол": surplus_qty,
            "излишек_сумма": surplus_sum,
            "недостача_кол": shortage_qty,
            "недостача_сумма": shortage_sum,
            # R4: empty book cells (None), separate from numeric zero used in metrics
            "книжная_норм_отсутствует": bool(miss_book_norm),
            "книжная_факт_отсутствует": bool(miss_book_fact),
        })
        meta.sku_count += 1

    if not rows:
        raise ValueError("Не удалось распознать строки позиций. Проверьте структуру файла.")

    df = pd.DataFrame(rows)
    meta.network_total_row = network_total
    return df, meta


def filter_by_period(df: pd.DataFrame, period_days: int, end_date: Optional[datetime.date]) -> pd.DataFrame:
    if df.empty or "дата_док" not in df.columns:
        return df
    dates = pd.to_datetime(df["дата_док"]).dt.date
    if end_date is None:
        end_date = dates.max()
    start_date = end_date - datetime.timedelta(days=period_days - 1)
    mask = (dates >= start_date) & (dates <= end_date)
    out = df.loc[mask].copy()
    out.attrs["period_start"] = start_date
    out.attrs["period_end"] = end_date
    return out
