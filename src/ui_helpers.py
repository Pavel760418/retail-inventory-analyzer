# -*- coding: utf-8 -*-
"""Streamlit UI helpers: KPI cards, executive insights, status table, chart frames."""
from __future__ import annotations

from typing import Any, Dict, List, Tuple

import pandas as pd

from src.analysis_service import AnalysisResult


def kpi_cards(result: AnalysisResult) -> List[Dict[str, Any]]:
    """Build KPI card specs from existing summary / anomaly / author metrics."""
    s = result.summary
    lines = int(s.get("sku_lines", 0) or 0)
    disc = int(s.get("sku_disc", 0) or 0)
    shortage = float(s.get("shortage", 0) or 0)
    surplus = float(s.get("surplus", 0) or 0)
    total_abs = shortage + surplus
    share = (disc / lines * 100.0) if lines else 0.0
    critical = int(result.anom_sum.get("critical", 0) or 0)
    # Critical positions also include chronic SKUs and class-C stores
    chronic_n = int(len(result.chronic)) if result.chronic is not None else 0
    crit_positions = critical + chronic_n

    return [
        {
            "label": "Позиций проверено",
            "value": f"{lines:,}",
            "delta": f"с расхождениями: {disc:,}",
            "help": "Число строк номенклатуры в периоде после парсинга TDSheet.",
        },
        {
            "label": "Общая сумма расхождений, ₽",
            "value": f"{total_abs:,.0f}",
            "delta": f"сальдо: {float(s.get('net', 0) or 0):,.0f}",
            "help": "Сумма недостач + сумма излишков (абсолютный масштаб расхождений).",
        },
        {
            "label": "Недостача, ₽",
            "value": f"{shortage:,.0f}",
            "delta": f"чистые: {float(s.get('clean_shortage', 0) or 0):,.0f}",
            "help": "Сумма недостач по модулю; «чистые» — после однородного перекрытия.",
        },
        {
            "label": "Излишки, ₽",
            "value": f"{surplus:,.0f}",
            "delta": f"магазинов: {int(s.get('stores', 0) or 0)}",
            "help": "Сумма излишков по всем магазинам периода.",
        },
        {
            "label": "Доля проблемных позиций, %",
            "value": f"{share:.1f}",
            "delta": f"{disc:,} из {lines:,}",
            "help": "Доля строк с излишком или недостачей относительно всех проверенных строк.",
        },
        {
            "label": "Критичные позиции",
            "value": f"{crit_positions:,}",
            "delta": f"аном.тип2: {critical}; хронич.: {chronic_n}",
            "help": "Критические аномалии книжных сумм (тип 2) + хронические проблемные SKU.",
        },
    ]


def executive_insights(result: AnalysisResult) -> List[str]:
    """3–5 fact-based insights; skip claims if fields absent."""
    lines: List[str] = []
    s = result.summary
    shortage = float(s.get("shortage", 0) or 0)
    surplus = float(s.get("surplus", 0) or 0)
    clean = float(s.get("clean_shortage", 0) or 0)
    overlap = float(s.get("overlap", 0) or 0)

    lines.append(
        f"За период {result.period_str} проверено {int(s.get('sku_lines', 0)):,} позиций "
        f"в {int(s.get('stores', 0))} магазинах по {int(s.get('docs', 0))} документам."
    )
    total = shortage + surplus
    if total > 0:
        lines.append(
            f"Финансовый масштаб расхождений: недостача {shortage:,.0f} ₽ "
            f"({shortage / total * 100:.1f}% от суммы |недостача+излишки|), "
            f"излишки {surplus:,.0f} ₽ ({surplus / total * 100:.1f}%)."
        )
    if shortage > 0:
        lines.append(
            f"Однородным пересортом перекрыто {overlap:,.0f} ₽ "
            f"({float(s.get('recovery_pct', 0) or 0):.1f}% недостач); "
            f"чистые недостачи — {clean:,.0f} ₽."
        )

    sm = result.store_metrics
    if sm is not None and not sm.empty and "чистые_недостачи" in sm.columns:
        worst = sm.sort_values("чистые_недостачи", ascending=False).head(1)
        if not worst.empty:
            row = worst.iloc[0]
            wsum = float(row["чистые_недостачи"])
            share = (wsum / clean * 100.0) if clean > 0 else 0.0
            lines.append(
                f"Наибольший риск по чистым недостачам — «{row['магазин']}»: "
                f"{wsum:,.0f} ₽ ({share:.1f}% чистых недостач сети)."
            )

    a = result.author_sum or {}
    if a.get("network_dominant"):
        lines.append(
            f"По вкладу в результат сети доминирует роль «{a['network_dominant']}» "
            f"({a.get('network_sign', '')}); цепочек ревизор↔оператор: "
            f"{int(a.get('chains_count', 0))} на сумму {float(a.get('chains_sum', 0) or 0):,.0f} ₽."
        )

    if result.anom_sum and int(result.anom_sum.get("total", 0)):
        lines.append(
            f"Аномалий книжных сумм: {int(result.anom_sum['total'])} "
            f"(критических тип 2: {int(result.anom_sum.get('critical', 0))})."
        )
    return lines[:5]


def status_risk_table(result: AnalysisResult) -> pd.DataFrame:
    """Status table from factual buckets (shortage / surplus / clean / anomalies / OK)."""
    df = result.df
    s = result.summary
    shortage_n = int((df["недостача_сумма"] > 0).sum()) if "недостача_сумма" in df.columns else 0
    surplus_n = int((df["излишек_сумма"] > 0).sum()) if "излишек_сумма" in df.columns else 0
    ok_n = int(len(df) - int(s.get("sku_disc", 0) or 0))
    crit_sum = 0.0
    if result.anom_df is not None and not result.anom_df.empty and "критичность" in result.anom_df.columns:
        crit = result.anom_df[result.anom_df["критичность"] == "ДА"]
        if "излишек_сумма" in crit.columns:
            crit_sum = float(crit["излишек_сумма"].fillna(0).sum())
    rows = [
        {
            "Статус": "🔴 Критично (аномалии тип 2)",
            "Количество позиций": int(result.anom_sum.get("critical", 0) or 0),
            "Сумма, ₽": crit_sum,
            "Действие": "Заполнить книжные суммы в 1С, пересчитать документ",
        },
        {
            "Статус": "🟠 Недостача",
            "Количество позиций": shortage_n,
            "Сумма, ₽": float(s.get("shortage", 0) or 0),
            "Действие": "Сверить пересорт / приходы / лист «Мероприятия»",
        },
        {
            "Статус": "🟡 Излишек",
            "Количество позиций": surplus_n,
            "Сумма, ₽": float(s.get("surplus", 0) or 0),
            "Действие": "Проверить однородность и оприходование смены",
        },
        {
            "Статус": "🟠 Чистые недостачи",
            "Количество позиций": "—",
            "Сумма, ₽": float(s.get("clean_shortage", 0) or 0),
            "Действие": "Приоритет контроля: не объяснены однородным пересортом",
        },
        {
            "Статус": "🟢 Без отклонений",
            "Количество позиций": max(ok_n, 0),
            "Сумма, ₽": 0.0,
            "Действие": "Контроль не требуется",
        },
    ]
    out = pd.DataFrame(rows)
    total = float(s.get("shortage", 0) or 0) + float(s.get("surplus", 0) or 0)
    out["Доля от расхождений"] = [
        (float(x) / total * 100.0) if total > 0 and isinstance(x, (int, float)) else 0.0
        for x in out["Сумма, ₽"]
    ]
    return out


def top_shortage_chart_df(result: AnalysisResult, n: int = 10) -> pd.DataFrame:
    sc = result.sku_cross
    if sc is None or sc.empty or "недостачи" not in sc.columns:
        return pd.DataFrame()
    top = sc[sc["недостачи"] > 0].sort_values("недостачи", ascending=False).head(n)
    if top.empty:
        return pd.DataFrame()
    return top[["наименование", "недостачи"]].rename(
        columns={"наименование": "Позиция", "недостачи": "Недостача, ₽"}
    )


def discrepancy_structure_df(result: AnalysisResult) -> pd.DataFrame:
    s = result.summary
    return pd.DataFrame({
        "Тип": ["Недостача", "Излишки", "Перекрытие", "Чистые недостачи"],
        "Сумма, ₽": [
            float(s.get("shortage", 0) or 0),
            float(s.get("surplus", 0) or 0),
            float(s.get("overlap", 0) or 0),
            float(s.get("clean_shortage", 0) or 0),
        ],
    })


def store_risk_chart_df(result: AnalysisResult, n: int = 12) -> pd.DataFrame:
    sm = result.store_metrics
    if sm is None or sm.empty or "чистые_недостачи" not in sm.columns:
        return pd.DataFrame()
    top = sm.sort_values("чистые_недостачи", ascending=False).head(n)
    return top[["магазин", "чистые_недостачи"]].rename(
        columns={"магазин": "Магазин", "чистые_недостачи": "Чистые недостачи, ₽"}
    )


def detail_table(result: AnalysisResult) -> pd.DataFrame:
    """Line-level detail sorted by financial risk."""
    df = result.df.copy()
    if df.empty:
        return df
    df["риск_сумма"] = df.get("недостача_сумма", 0).fillna(0) + df.get("излишек_сумма", 0).fillna(0)
    cols = [
        c for c in [
            "магазин", "документ", "наименование", "автор_дока",
            "излишек_кол", "излишек_сумма", "недостача_кол", "недостача_сумма",
            "сумма_учетная", "сумма_факт", "category_label", "hierarchy_path",
            "hierarchy_matched", "риск_сумма",
        ] if c in df.columns
    ]
    return df[cols].sort_values("риск_сумма", ascending=False)
