# -*- coding: utf-8 -*-
"""Пересорт / overlap matching (intra-doc, cross-doc, cross-store)."""
from __future__ import annotations

from typing import Tuple

import pandas as pd

from config.settings import MIN_OVERLAP_SCORE, PAIR_HEAD_LIMIT, PAIR_VOLUME_LIMIT
from src.text_normalize import (
    clean_tokens,
    comparable_units,
    extract_weight,
    price_per_unit,
    product_key,
)

def similarity_score_row(src: pd.Series, dst: pd.Series) -> float:
    u1 = src.get("единица", "")
    u2 = dst.get("единица", "")
    if not comparable_units(u1, u2):
        return -999
    t1 = set(clean_tokens(src["наименование"]))
    t2 = set(clean_tokens(dst["наименование"]))
    inter = len(t1 & t2)
    if inter == 0:
        return -999
    score = inter * 10
    k1 = product_key(src["наименование"], 3)
    k2 = product_key(dst["наименование"], 3)
    if k1 and k2 and k1 == k2:
        score += 20
    w1, wu1 = extract_weight(src["наименование"])
    w2, wu2 = extract_weight(dst["наименование"])
    if w1 and w2 and wu1 == wu2:
        diff = abs(w1 - w2) / max(w1, w2)
        if diff <= 0.10:
            score += 20
        elif diff <= 0.25:
            score += 10
        elif diff <= 0.50:
            score += 3
        else:
            score -= 20
    p1 = price_per_unit(float(src.get("излишек_кол", 0)), float(src.get("излишек_сумма", 0)))
    p2 = price_per_unit(float(dst.get("недостача_кол", 0)), float(dst.get("недостача_сумма", 0)))
    if p1 > 0 and p2 > 0:
        pdiff = abs(p1 - p2) / max(p1, p2)
        if pdiff <= 0.10:
            score += 20
        elif pdiff <= 0.20:
            score += 12
        elif pdiff <= 0.35:
            score += 5
        else:
            score -= 25
    return score


def homogeneous_pair(src: pd.Series, dst: pd.Series) -> Tuple[bool, str]:
    g1 = str(src.get("category_group", ""))
    g2 = str(dst.get("category_group", ""))
    if g1 and g2 and g1 == g2:
        return True, "same_category"
    if str(src.get("артикул", "")) and str(src.get("артикул", "")) == str(dst.get("артикул", "")):
        return True, "same_article"
    if product_key(src["наименование"]) == product_key(dst["наименование"]):
        fb = bool(src.get("category_fallback")) or bool(dst.get("category_fallback"))
        return True, "fallback_product_key" if fb else "product_key"
    return False, "rejected_category"


def find_peresor(df: pd.DataFrame, level: str = "intra_doc",
                 source_doc: str = "", target_doc: str = "",
                 source_store: str = "", target_store: str = "") -> pd.DataFrame:
    cols = [
        "уровень", "магазин_изл", "док_изл", "магазин_нед", "док_нед", "группа",
        "излишек_товар", "излишек_арт", "излишек_кол", "излишек_сумма",
        "недостача_товар", "недостача_арт", "недостача_кол", "недостача_сумма",
        "перекрытие_кол", "перекрытие_сум", "единица", "score", "category_group",
        "homogeneity", "fallback_flag",
        "цена_излишка", "цена_недостачи", "разница_цены_%",
    ]
    ds = df[df["излишек_кол"] > 0].copy()
    dd = df[df["недостача_кол"] > 0].copy()
    if ds.empty or dd.empty:
        return pd.DataFrame(columns=cols)

    ds = ds.reset_index(drop=True)
    dd = dd.reset_index(drop=True)
    ds["avail_qty"] = ds["излишек_кол"].astype(float)
    ds["avail_sum"] = ds["излишек_сумма"].astype(float)
    ds["price"] = ds.apply(lambda r: price_per_unit(r["излишек_кол"], r["излишек_сумма"]), axis=1)
    dd["need_qty"] = dd["недостача_кол"].astype(float)
    dd["need_sum"] = dd["недостача_сумма"].astype(float)
    dd["price"] = dd.apply(lambda r: price_per_unit(r["недостача_кол"], r["недостача_сумма"]), axis=1)

    # Блокирующий отбор по категории уменьшает квадратичность.
    ds = ds.sort_values("излишек_сумма", ascending=False)
    dd = dd.sort_values("недостача_сумма", ascending=False)
    # Guardrail: если объём слишком большой, ограничиваем пул наименее значимые хвосты.
    pair_volume = len(ds) * len(dd)
    if pair_volume > PAIR_VOLUME_LIMIT:
        ds = ds.head(PAIR_HEAD_LIMIT).copy()
        dd = dd.head(PAIR_HEAD_LIMIT).copy()

    candidates = []
    by_cat_ds = {k: g for k, g in ds.groupby(ds["category_group"].fillna(""), dropna=False)}
    by_cat_dd = {k: g for k, g in dd.groupby(dd["category_group"].fillna(""), dropna=False)}
    common_cats = set(by_cat_ds.keys()) & set(by_cat_dd.keys())

    for cat in common_cats:
        gds = by_cat_ds[cat]
        gdd = by_cat_dd[cat]
        for si, sr in gds.iterrows():
            for di, dr in gdd.iterrows():
                if str(sr.get("артикул", "")) == str(dr.get("артикул", "")) and str(sr.get("артикул", "")):
                    same_sku = True
                else:
                    same_sku = False
                    if str(sr.get("артикул", "")) == str(dr.get("артикул", "")) and sr["наименование"] != dr["наименование"]:
                        continue
                ok, hom = homogeneous_pair(sr, dr)
                if not ok:
                    continue
                score = 100.0 if same_sku else similarity_score_row(sr, dr)
                if score < MIN_OVERLAP_SCORE:
                    continue
                p1, p2 = sr["price"], dr["price"]
                pdiff = abs(p1 - p2) / max(p1, p2) if p1 > 0 and p2 > 0 else 999
                candidates.append((score, pdiff, si, di, hom))

    candidates.sort(key=lambda x: (-x[0], x[1]))
    pairs = []
    for score, pdiff, si, di, hom in candidates:
        if ds.at[si, "avail_qty"] <= 0 or dd.at[di, "need_qty"] <= 0:
            continue
        cover_qty = min(ds.at[si, "avail_qty"], dd.at[di, "need_qty"])
        if cover_qty <= 0:
            continue
        src_price = ds.at[si, "price"]
        dst_price = dd.at[di, "price"]
        cover_sum = min(
            ds.at[si, "avail_sum"],
            round(cover_qty * dst_price, 2) if dst_price > 0 else dd.at[di, "need_sum"],
        )
        if cover_sum <= 0:
            continue
        ds.at[si, "avail_qty"] = round(ds.at[si, "avail_qty"] - cover_qty, 6)
        ds.at[si, "avail_sum"] = round(ds.at[si, "avail_sum"] - cover_sum, 2)
        dd.at[di, "need_qty"] = round(dd.at[di, "need_qty"] - cover_qty, 6)
        dd.at[di, "need_sum"] = round(dd.at[di, "need_sum"] - cover_sum, 2)
        sr, dr = ds.loc[si], dd.loc[di]
        pairs.append({
            "уровень": level,
            "магазин_изл": source_store or sr.get("магазин", ""),
            "док_изл": source_doc or sr.get("документ", ""),
            "магазин_нед": target_store or dr.get("магазин", ""),
            "док_нед": target_doc or dr.get("документ", ""),
            "группа": product_key(sr["наименование"], 3) or product_key(dr["наименование"], 3),
            "излишек_товар": sr["наименование"],
            "излишек_арт": sr.get("артикул", ""),
            "излишек_кол": float(sr["излишек_кол"]),
            "излишек_сумма": float(sr["излишек_сумма"]),
            "недостача_товар": dr["наименование"],
            "недостача_арт": dr.get("артикул", ""),
            "недостача_кол": float(dr["недостача_кол"]),
            "недостача_сумма": float(dr["недостача_сумма"]),
            "перекрытие_кол": round(cover_qty, 3),
            "перекрытие_сум": round(cover_sum, 2),
            "единица": sr.get("единица", ""),
            "score": score,
            "category_group": sr.get("category_group", ""),
            "homogeneity": hom,
            "fallback_flag": "ДА" if hom.startswith("fallback") or sr.get("category_fallback") or dr.get("category_fallback") else "НЕТ",
            "цена_излишка": round(src_price, 2) if src_price else 0,
            "цена_недостачи": round(dst_price, 2) if dst_price else 0,
            "разница_цены_%": round(abs(src_price - dst_price) / max(src_price, dst_price), 4) if src_price and dst_price else None,
        })
    return pd.DataFrame(pairs, columns=cols)


def build_all_overlaps(df: pd.DataFrame, enable_cross_store: bool = True) -> pd.DataFrame:
    parts = []
    for (store, doc), grp in df.groupby(["магазин", "документ"]):
        parts.append(find_peresor(grp, level="intra_doc", source_store=store, target_store=store,
                                  source_doc=doc, target_doc=doc))
    for store, grp in df.groupby("магазин"):
        docs = grp["документ"].unique()
        if len(docs) < 2:
            continue
        parts.append(find_peresor(grp, level="cross_doc", source_store=store, target_store=store))
    if enable_cross_store:
        parts.append(find_peresor(df, level="cross_store_analytical"))
    frames = [p for p in parts if p is not None and not p.empty]
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True)
    return out.sort_values(["перекрытие_сум", "score"], ascending=[False, False])


def build_operational_overlaps(df: pd.DataFrame) -> pd.DataFrame:
    """Перекрытие внутри магазина (между документами), без двойного подсчёта intra+cross."""
    parts = []
    for store, grp in df.groupby("магазин"):
        parts.append(find_peresor(grp, level="cross_doc", source_store=store, target_store=store))
    frames = [p for p in parts if not p.empty]
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True).sort_values(["перекрытие_сум", "score"], ascending=[False, False])


def build_analytical_cross_store(df: pd.DataFrame) -> pd.DataFrame:
    return find_peresor(df, level="cross_store_analytical")
