# -*- coding: utf-8 -*-
"""Nomenclature enrichment: master hierarchy is the source of truth.

Heuristic keyword classification is disabled for runtime category assignment.
Optional external CSV catalog may still override article/unit fields only when provided.
"""
from __future__ import annotations

import os
from typing import Optional

import pandas as pd

from src.master_hierarchy import (
    UNMATCHED_LABEL,
    MasterHierarchy,
    get_master_hierarchy,
    map_by_master_hierarchy,
)
from src.text_normalize import infer_unit, norm_txt, sku_key


def load_catalog(path: Optional[str]) -> pd.DataFrame:
    """Optional legacy CSV catalog (article/unit hints). Hierarchy comes from embedded master."""
    if not path or not os.path.exists(path):
        return pd.DataFrame(columns=["наименование", "артикул", "единица", "category_group"])
    df = pd.read_csv(path, dtype=str)
    df.columns = [norm_txt(c) for c in df.columns]
    rename = {}
    for c in df.columns:
        if "наимен" in c:
            rename[c] = "наименование"
        elif "арт" in c:
            rename[c] = "артикул"
        elif "един" in c:
            rename[c] = "единица"
        elif "кateg" in c or "групп" in c or "category" in c:
            rename[c] = "category_group"
    df = df.rename(columns=rename)
    if "наименование" in df.columns:
        df["sku_key"] = df["наименование"].apply(sku_key)
    return df


def enrich_dataframe(
    df: pd.DataFrame,
    catalog: Optional[pd.DataFrame] = None,
    hierarchy: Optional[MasterHierarchy] = None,
) -> pd.DataFrame:
    """Enrich inventory rows with etalon hierarchy from embedded master reference.

    Category / hierarchy are taken ONLY from the master template reference.
    Unknown SKUs are marked unmatched — no heuristic category invention.
    """
    out = df.copy()
    if "sku_key" not in out.columns:
        out["sku_key"] = out["наименование"].apply(sku_key)
    if "единица" not in out.columns:
        out["единица"] = out["наименование"].apply(infer_unit)
    if "артикул" not in out.columns:
        out["артикул"] = ""

    href = hierarchy or get_master_hierarchy()
    max_depth = href.max_depth

    cat_groups, cat_labels, cat_methods, cat_fallbacks = [], [], [], []
    matched_flags, path_strs, leaf_groups = [], [], []
    level_cols = {f"level_{i}": [] for i in range(max_depth)}

    catalog_by_key = {}
    if catalog is not None and not catalog.empty and "sku_key" in catalog.columns:
        for _, crow in catalog.iterrows():
            k = crow.get("sku_key")
            if k and k not in catalog_by_key:
                catalog_by_key[k] = crow

    for _, row in out.iterrows():
        mapped = map_by_master_hierarchy(row["наименование"], href)
        cat_groups.append(mapped["category_group"])
        cat_labels.append(mapped["category_label"])
        cat_methods.append(mapped["category_method"])
        cat_fallbacks.append(bool(mapped["category_fallback"]))
        matched_flags.append(bool(mapped["hierarchy_matched"]))
        path_strs.append(mapped["hierarchy_path"])
        leaf_groups.append(mapped["leaf_group"])
        levels = mapped.get("levels") or {}
        for i in range(max_depth):
            level_cols[f"level_{i}"].append(levels.get(f"level_{i}", ""))

    out["category_group"] = cat_groups
    out["category_label"] = cat_labels
    out["category_method"] = cat_methods
    out["category_fallback"] = cat_fallbacks
    out["hierarchy_matched"] = matched_flags
    out["hierarchy_path"] = path_strs
    out["leaf_group"] = leaf_groups
    for col, values in level_cols.items():
        out[col] = values

    def _resolve_article(row):
        art = str(row.get("артикул", "")).strip()
        if art:
            return art
        hit = catalog_by_key.get(row["sku_key"])
        if hit is not None:
            return str(hit.get("артикул", row["sku_key"][:20]))
        return row["sku_key"][:20]

    out["артикул"] = out.apply(_resolve_article, axis=1)

    # Optional unit hint from external catalog only
    if catalog_by_key:
        def _resolve_unit(row):
            u = str(row.get("единица", "")).strip()
            hit = catalog_by_key.get(row["sku_key"])
            if hit is not None and str(hit.get("единица", "")).strip():
                return str(hit.get("единица")).strip()
            return u or infer_unit(row["наименование"])

        out["единица"] = out.apply(_resolve_unit, axis=1)

    return out


def unmatched_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Unique unmatched SKU names for review."""
    if df.empty or "hierarchy_matched" not in df.columns:
        return pd.DataFrame(columns=["наименование", "sku_key", "строк"])
    um = df[df["hierarchy_matched"] == False]  # noqa: E712
    if um.empty:
        return pd.DataFrame(columns=["наименование", "sku_key", "строк"])
    g = (
        um.groupby(["наименование", "sku_key"], dropna=False)
        .size()
        .reset_index(name="строк")
        .sort_values("строк", ascending=False)
    )
    return g


def display_category(row) -> str:
    """Human-readable category for Excel/UI (never shows __unmatched__ keys)."""
    label = row.get("category_label")
    if label:
        return str(label)
    cg = str(row.get("category_group", "") or "")
    if cg.startswith(UNMATCHED_LABEL) or cg.startswith("__unmatched__"):
        return UNMATCHED_LABEL
    return cg
