# -*- coding: utf-8 -*-
"""Detect inventory book-sum anomalies (Release 4).

Uses empty-cell flags from the parser (книжная_норм_отсутствует /
книжная_факт_отсутствует). Does NOT alter surplus/shortage calculations.

Types (mutually exclusive primary label):
  Аномалия 1 — нет книжной нормативной, есть фактическая; риск нулевой
               себестоимости / пересчёт уходит в излишки.
  Аномалия 2 — нет ни нормативной, ни фактической; искажает итог
               (нет ни недостач, ни излишков по сумме). Критическая.
  Аномалия 3 — только фактическая (нормативной нет); сумма внесена
               вручную в пересчёт; нужна сверка (в т.ч. по другим складам).
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Set

import pandas as pd

TYPE_1 = "Аномалия 1"
TYPE_2 = "Аномалия 2"
TYPE_3 = "Аномалия 3"

TYPE_1_CODE = "missing_normative_zero_cost"
TYPE_2_CODE = "missing_both_book_sums"
TYPE_3_CODE = "missing_normative_manual_fact"

WHAT_HAPPENED = {
    TYPE_1: (
        "Не указана сумма книжная нормативная при наличии фактической. "
        "Пересчёт может встать в излишки; фактически это может означать нулевую себестоимость."
    ),
    TYPE_2: (
        "Не указаны ни книжная нормативная, ни книжная фактическая сумма. "
        "По позиции не возникает ни недостач, ни излишков по сумме — итог инвентаризации искажается."
    ),
    TYPE_3: (
        "Есть только сумма книжная фактическая, нормативной нет. "
        "Сумма в пересчёт, скорее всего, внесена вручную — требуется перепроверка "
        "(в т.ч. по другим подразделениям/складам)."
    ),
}

SHORT_MEANING = {
    TYPE_1: "Нет книжной нормативной → риск нулевой себестоимости / излишков из пересчёта.",
    TYPE_2: "Нет обеих книжных сумм → нет расхождений по сумме, итог искажён.",
    TYPE_3: "Только фактическая книжная → ручной пересчёт, нужна сверка.",
}


def _ensure_flags(df: pd.DataFrame) -> pd.DataFrame:
    """Backward-compatible: derive flags from numeric zeros if parser flags absent."""
    out = df.copy()
    if "книжная_норм_отсутствует" not in out.columns:
        # Fallback: treat exact 0 with no qty book as uncertain — prefer explicit flags
        out["книжная_норм_отсутствует"] = False
    if "книжная_факт_отсутствует" not in out.columns:
        out["книжная_факт_отсутствует"] = False
    return out


def _names_by_store(df: pd.DataFrame) -> Dict[str, Set[str]]:
    mapping: Dict[str, Set[str]] = {}
    if df.empty:
        return mapping
    work = df.copy()
    work["name_exact"] = work["наименование"].astype(str).str.strip()
    for store, grp in work.groupby("магазин"):
        mapping[str(store)] = set(grp["name_exact"].tolist())
    return mapping


def detect_book_sum_anomalies(df: pd.DataFrame) -> pd.DataFrame:
    """Return one row per anomalous inventory line with type, criticality, comment."""
    if df is None or df.empty:
        return pd.DataFrame()

    work = _ensure_flags(df)
    names_map = _names_by_store(work)
    rows: List[dict] = []

    for _, row in work.iterrows():
        miss_n = bool(row.get("книжная_норм_отсутствует"))
        miss_f = bool(row.get("книжная_факт_отсутствует"))
        if not miss_n and not miss_f:
            continue
        if not miss_n:
            # Only fact missing without norm missing — not in the R4 triad
            continue

        store = str(row.get("магазин", "") or "")
        name = str(row.get("наименование", "") or "").strip()
        book_n = float(row.get("сумма_учетная", 0) or 0)
        book_f = float(row.get("сумма_факт", 0) or 0)
        surplus = float(row.get("излишек_сумма", 0) or 0)
        shortage = float(row.get("недостача_сумма", 0) or 0)

        other_stores = [
            s for s, names in names_map.items()
            if s != store and name in names
        ]

        if miss_n and miss_f:
            atype = TYPE_2
            code = TYPE_2_CODE
            critical = True
            comment = (
                "Критично: обе книжные суммы пустые — позиция не формирует "
                "недостач/излишков по сумме."
            )
        elif miss_n and not miss_f:
            # Prefer Аномалия 3 when same name exists elsewhere (сверка);
            # otherwise Аномалия 1 (zero-cost / surplus risk).
            if other_stores:
                atype = TYPE_3
                code = TYPE_3_CODE
                critical = False
                comment = (
                    "Требует сверки: только фактическая книжная сумма; "
                    f"то же наименование есть в: {', '.join(other_stores[:5])}"
                    + ("…" if len(other_stores) > 5 else "")
                )
            else:
                atype = TYPE_1
                code = TYPE_1_CODE
                critical = False
                comment = (
                    "Аномалия: нет книжной нормативной при наличии фактической; "
                    "пересчёт может уйти в излишки (риск нулевой себестоимости)."
                )
                if surplus > 0:
                    comment += f" Излишек по строке: {surplus:,.2f} руб."
        else:
            continue

        rows.append({
            "наименование": name,
            "магазин": store,
            "документ": str(row.get("документ", "") or ""),
            "тип_аномалии": atype,
            "код_аномалии": code,
            "книжная_нормативная_сумма": None if miss_n else book_n,
            "книжная_фактическая_сумма": None if miss_f else book_f,
            "книжная_норм_отсутствует": miss_n,
            "книжная_факт_отсутствует": miss_f,
            "излишек_сумма": surplus,
            "недостача_сумма": shortage,
            "что_произошло": WHAT_HAPPENED[atype],
            "критичность": "ДА" if critical else "НЕТ",
            "критическая": critical,
            "комментарий": comment,
            "другие_магазины": "; ".join(other_stores),
            "магазинов_с_тем_же_именем": len(other_stores) + (1 if name else 0),
        })

    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    # Critical first, then by surplus impact
    out["_ord"] = out["тип_аномалии"].map({TYPE_2: 0, TYPE_1: 1, TYPE_3: 2}).fillna(9)
    out = out.sort_values(
        ["_ord", "критическая", "излишек_сумма"],
        ascending=[True, False, False],
    ).drop(columns=["_ord"])
    return out.reset_index(drop=True)


def anomalies_summary(anom_df: pd.DataFrame, total_rows: int = 0) -> Dict[str, Any]:
    """KPI dict for Сводка / Выводы / Streamlit."""
    empty = {
        "total": 0,
        "critical": 0,
        "non_critical": 0,
        "type1": 0,
        "type2": 0,
        "type3": 0,
        "stores_affected": 0,
        "share_pct": 0.0,
        "surplus_impact": 0.0,
        "meanings": SHORT_MEANING,
    }
    if anom_df is None or anom_df.empty:
        return empty
    total = int(len(anom_df))
    t = anom_df["тип_аномалии"]
    critical = int((anom_df["критичность"] == "ДА").sum()) if "критичность" in anom_df.columns else 0
    base = total_rows if total_rows > 0 else total
    return {
        "total": total,
        "critical": critical,
        "non_critical": total - critical,
        "type1": int((t == TYPE_1).sum()),
        "type2": int((t == TYPE_2).sum()),
        "type3": int((t == TYPE_3).sum()),
        "stores_affected": int(anom_df["магазин"].nunique()) if "магазин" in anom_df.columns else 0,
        "share_pct": (total / base * 100.0) if base else 0.0,
        "surplus_impact": float(anom_df.get("излишек_сумма", pd.Series(dtype=float)).fillna(0).sum()),
        "meanings": SHORT_MEANING,
    }


def enrich_store_metrics_with_anomalies(
    store_metrics: pd.DataFrame,
    anom_df: pd.DataFrame,
) -> pd.DataFrame:
    sm = store_metrics.copy()
    if sm.empty:
        return sm
    sm["аномалий"] = 0
    sm["аномалий_критич"] = 0
    if anom_df is None or anom_df.empty:
        return sm
    agg = anom_df.groupby("магазин").agg(
        аномалий=("тип_аномалии", "count"),
        аномалий_критич=("критичность", lambda s: int((s == "ДА").sum())),
    )
    sm = sm.set_index("магазин")
    for col in agg.columns:
        sm[col] = agg[col]
    sm[list(agg.columns)] = sm[list(agg.columns)].fillna(0)
    return sm.reset_index()
