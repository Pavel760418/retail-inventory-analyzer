# -*- coding: utf-8 -*-
"""Store/document/SKU metrics, network summary, conclusions."""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from config.settings import CHRONIC_MIN_DOCS, CHRONIC_MIN_STORES, NON_RETAIL_STORES
from src.overlap import find_peresor

def percentile_rank(series: pd.Series, invert: bool = False) -> pd.Series:
    if series.empty:
        return series
    r = series.rank(pct=True) * 100
    return 100 - r if invert else r


def calc_store_metrics(df: pd.DataFrame, overlap_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    overlap_by_store = overlap_df.groupby("магазин_нед")["перекрытие_сум"].sum().to_dict() if not overlap_df.empty else {}
    for store, grp in df.groupby("магазин"):
        sur = grp["излишек_сумма"].sum()
        sh = grp["недостача_сумма"].sum()
        book = grp["сумма_учетная"].sum()
        overlap = overlap_by_store.get(store, 0)
        clean = max(sh - overlap, 0)
        docs = grp["документ"].nunique()
        sku_disc = ((grp["излишек_сумма"] > 0) | (grp["недостача_сумма"] > 0)).sum()
        shrink = (clean / book * 100) if book > 0 else 0
        recovery = (overlap / sh * 100) if sh > 0 else 0
        rows.append({
            "магазин": store,
            "излишки": sur,
            "недостачи": sh,
            "сальдо": sur - sh,
            "перекрытие": overlap,
            "чистые_недостачи": clean,
            "shrinkage_pct": shrink,
            "recovery_pct": recovery,
            "документов": docs,
            "sku_с_расх": sku_disc,
            "book_sum": book,
            "is_retail": store not in NON_RETAIL_STORES,
        })
    sm = pd.DataFrame(rows)
    if sm.empty:
        return sm
    sm["score_shrink"] = percentile_rank(sm["shrinkage_pct"], invert=True)
    sm["score_recovery"] = percentile_rank(sm["recovery_pct"])
    sm["score_net"] = percentile_rank(sm["чистые_недостачи"], invert=True)
    sm["store_score"] = (
        sm["score_shrink"] * 0.35 + sm["score_recovery"] * 0.25 +
        sm["score_net"] * 0.25 + percentile_rank(sm["sku_с_расх"], invert=True) * 0.15
    ).round(1)
    sm["класс"] = pd.cut(
        sm["store_score"], bins=[-1, 50, 80, 101], labels=["C", "B", "A"]
    ).astype(str)
    sm = sm.sort_values("store_score", ascending=False)
    sm["место"] = range(1, len(sm) + 1)
    return sm


def calc_document_metrics(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (store, doc), grp in df.groupby(["магазин", "документ"]):
        sur = grp["излишек_сумма"].sum()
        sh = grp["недостача_сумма"].sum()
        ovl = find_peresor(grp, level="intra_doc", source_store=store, target_store=store,
                           source_doc=doc, target_doc=doc)
        ov_sum = ovl["перекрытие_сум"].sum() if not ovl.empty else 0
        clean = max(sh - ov_sum, 0)
        dt = grp["дата_док"].iloc[0]
        author = grp["автор_дока"].iloc[0] if "автор_дока" in grp.columns else "неизвестно"
        severity = 0.5 * sh + 0.3 * clean + 0.2 * (grp["недостача_кол"] > 0).sum() * 100
        if sh >= 50000:
            cls = "Критический"
        elif sh >= 10000:
            cls = "Высокий"
        elif sh > 0:
            cls = "Средний"
        elif sur > sh:
            cls = "Излишки"
        else:
            cls = "Нейтральный"
        rows.append({
            "магазин": store, "документ": doc, "дата": dt,
            "автор": author,
            "излишки": sur, "недостачи": sh, "сальдо": sur - sh,
            "перекрытие": ov_sum, "чистые_недостачи": clean,
            "severity": severity, "класс": cls,
            "sku_изл": (grp["излишек_сумма"] > 0).sum(),
            "sku_нед": (grp["недостача_сумма"] > 0).sum(),
        })
    dm = pd.DataFrame(rows).sort_values("severity", ascending=False)
    dm["рейтинг"] = range(1, len(dm) + 1)
    return dm


def calc_sku_cross(df: pd.DataFrame) -> pd.DataFrame:
    agg_map = {
        "наименование": ("наименование", "first"),
        "артикул": ("артикул", "first"),
        "category_group": ("category_group", "first"),
        "магазинов": ("магазин", "nunique"),
        "документов": ("документ", "nunique"),
        "излишки": ("излишек_сумма", "sum"),
        "недостачи": ("недостача_сумма", "sum"),
        "изл_кол": ("излишек_кол", "sum"),
        "нед_кол": ("недостача_кол", "sum"),
    }
    if "category_label" in df.columns:
        agg_map["category_label"] = ("category_label", "first")
    if "hierarchy_path" in df.columns:
        agg_map["hierarchy_path"] = ("hierarchy_path", "first")
    if "hierarchy_matched" in df.columns:
        agg_map["hierarchy_matched"] = ("hierarchy_matched", "first")

    g = df.groupby("sku_key").agg(**{k: v for k, v in agg_map.items()}).reset_index()
    # Display-friendly category for Excel sheets
    if "category_label" in g.columns:
        g["category_group"] = g["category_label"]
    g["сальдо"] = g["излишки"] - g["недостачи"]
    g["chronic"] = (
        (g["магазинов"] >= CHRONIC_MIN_STORES) | (g["документов"] >= CHRONIC_MIN_DOCS)
    ) & (g["недостачи"] > 0)
    return g.sort_values("недостачи", ascending=False)


def build_sku_store_pivot(df: pd.DataFrame) -> pd.DataFrame:
    p = df.pivot_table(
        index="наименование", columns="магазин", values="недостача_сумма",
        aggfunc="sum", fill_value=0,
    )
    return p


def calc_network_summary(df: pd.DataFrame, overlap_df: pd.DataFrame) -> Dict[str, Any]:
    sur = df["излишек_сумма"].sum()
    sh = df["недостача_сумма"].sum()
    ov = overlap_df["перекрытие_сум"].sum() if not overlap_df.empty else 0
    clean = max(sh - ov, 0)
    book = df["сумма_учетная"].sum()
    # Излишки, которые не участвовали в однородном перекрытии (требуют отдельной проверки)
    if not overlap_df.empty:
        used_sur_arts = set(overlap_df["излишек_арт"].astype(str))
        non_hom_sur = df[(df["излишек_сумма"] > 0) & (~df["артикул"].astype(str).isin(used_sur_arts))]["излишек_сумма"].sum()
    else:
        non_hom_sur = sur
    return {
        "stores": df["магазин"].nunique(),
        "docs": df["документ"].nunique(),
        "sku_lines": len(df),
        "sku_disc": ((df["излишек_сумма"] > 0) | (df["недостача_сумма"] > 0)).sum(),
        "surplus": sur, "shortage": sh, "net": sur - sh,
        "overlap": ov, "clean_shortage": clean,
        "non_homogeneous_surplus": non_hom_sur,
        "shrinkage_pct": (clean / book * 100) if book > 0 else 0,
        "recovery_pct": (ov / sh * 100) if sh > 0 else 0,
    }


def generate_conclusions(
    summary: Dict,
    store_metrics: pd.DataFrame,
    chronic: pd.DataFrame,
    cap_summary: Optional[Dict[str, Any]] = None,
    anom_summary: Optional[Dict[str, Any]] = None,
    author_summary: Optional[Dict[str, Any]] = None,
) -> List[Tuple[str, str]]:
    best = store_metrics.head(3)["магазин"].tolist() if not store_metrics.empty else []
    worst = store_metrics.tail(3)["магазин"].tolist() if not store_metrics.empty else []
    lines = [
        ("Масштаб", f"Магазинов: {summary['stores']} | Документов: {summary['docs']} | "
         f"Позиций с расхождениями: {summary['sku_disc']:,}"),
        ("Финансы", f"Излишки: {summary['surplus']:,.0f} руб. | Недостачи: {summary['shortage']:,.0f} руб. | "
         f"Сальдо: {summary['net']:,.0f} руб."),
        ("Пересорт", f"Перекрыто однородным пересортом: {summary['overlap']:,.0f} руб. "
         f"({summary['recovery_pct']:.1f}% от недостач) | "
         f"Чистые недостачи: {summary['clean_shortage']:,.0f} руб."),
        ("Чистые недостачи", (
            "Это недостачи, которые нельзя объяснить пересортом внутри одной товарной группы. "
            "Например: мясо не перекрывается напитками, овощи — молочной группой. "
            "Такие позиции требуют перепроверки, сверки приходных документов и анализа прошлых пересчётов."
        )),
        ("Неоднородные излишки", (
            f"Излишки без однородного перекрытия: {summary.get('non_homogeneous_surplus', 0):,.0f} руб. "
            "Их нельзя автоматически зачесть против чужих категорий — нужна ручная проверка."
        )),
        ("Коэффициент потерь", f"Чистые недостачи к учётной сумме: {summary['shrinkage_pct']:.2f}%"),
        ("Рейтинг", f"Лучшие: {', '.join(best) if best else '—'} | "
         f"Требуют внимания: {', '.join(worst) if worst else '—'}"),
        ("Хронические позиции", f"Повторяющихся проблемных товаров: {len(chronic):,}"),
        ("Действия", "Детальный план — лист «Мероприятия». По магазинам — отдельные листы. "
         "Аномалии книжных сумм — лист «Аномалии». "
         "Сверка оприходования смены — лист «Оприходование излишков». "
         "Вклад ревизоров и операторов — листы «Ревизоры vs Операторы» и «Цепочки пересортов»."),
    ]
    insert_at = 3
    if author_summary:
        from src.author import format_author_conclusion
        lines.insert(insert_at, format_author_conclusion(author_summary))
        insert_at += 1
    if anom_summary and int(anom_summary.get("total", 0) or 0) > 0:
        lines.insert(
            insert_at,
            (
                "Аномалии книжных сумм",
                (
                    f"Всего: {int(anom_summary.get('total', 0))} "
                    f"({anom_summary.get('share_pct', 0):.1f}% строк) | "
                    f"Критических (тип 2): {int(anom_summary.get('critical', 0))} | "
                    f"Тип 1 (нет норм., риск нулевой с/с): {int(anom_summary.get('type1', 0))} | "
                    f"Тип 3 (только факт., нужна сверка): {int(anom_summary.get('type3', 0))}. "
                    "Пустые книжные суммы искажают итог — см. лист «Аномалии»."
                ),
            ),
        )
        insert_at += 1
    if cap_summary:
        lines.insert(
            insert_at,
            (
                "Оприходование излишков (закрытие смены)",
                (
                    f"Складов в файле: {int(cap_summary.get('warehouses', 0))} | "
                    f"Строк: {int(cap_summary.get('cap_rows', 0))} | "
                    f"Сумма: {cap_summary.get('cap_sum', 0):,.0f} руб. | "
                    f"Сопоставлено: {int(cap_summary.get('matched', 0)) + int(cap_summary.get('matched_shortage', 0))} | "
                    f"Связь с недостачей: {int(cap_summary.get('matched_shortage', 0))} | "
                    f"Не найдено по имени: {int(cap_summary.get('unmatched_name', 0))} | "
                    f"Склад не сопоставлен/неоднозначен: "
                    f"{int(cap_summary.get('unmatched_store', 0)) + int(cap_summary.get('uncertain_store', 0))}. "
                    "Сопоставление только точное и только внутри своего склада."
                ),
            ),
        )
    return lines
