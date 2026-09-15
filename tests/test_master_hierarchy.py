# -*- coding: utf-8 -*-
"""Tests for embedded master hierarchy lookup (etalon)."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.catalog import enrich_dataframe, unmatched_summary
from src.master_hierarchy import (
    UNMATCHED_LABEL,
    UNMATCHED_PREFIX,
    clear_hierarchy_cache,
    load_master_hierarchy,
    map_by_master_hierarchy,
)


FIXTURE = ROOT / "data" / "fixtures" / "master_hierarchy_sample.json"
PROD_JSON = ROOT / "data" / "master_hierarchy.json"


@pytest.fixture
def sample_hierarchy():
    clear_hierarchy_cache()
    return load_master_hierarchy(FIXTURE)


def test_fixture_lookup_matched(sample_hierarchy):
    m = map_by_master_hierarchy("Молоко тестовое 1л", sample_hierarchy)
    assert m["hierarchy_matched"] is True
    assert m["category_method"] == "master_hierarchy"
    assert m["category_fallback"] is False
    assert m["leaf_group"] == "Молоко пастеризованное"
    assert "Молочная продукция" in m["hierarchy_path"]
    assert m["category_group"] == "Молоко пастеризованное"


def test_fixture_lookup_unmatched_no_heuristic(sample_hierarchy):
    """Unknown SKU must NOT get keyword/heuristic category."""
    m = map_by_master_hierarchy("Совершенно неизвестный товар XYZ", sample_hierarchy)
    assert m["hierarchy_matched"] is False
    assert m["category_method"] == "master_unmatched"
    assert m["category_fallback"] is True
    assert m["category_label"] == UNMATCHED_LABEL
    assert str(m["category_group"]).startswith(UNMATCHED_PREFIX)
    # Must not invent «Напитки» / «Молочные» etc.
    assert m["category_group"] not in {"Напитки", "Молочные продукты", "Бакалея"}


def test_enrich_uses_master_not_keyword(sample_hierarchy):
    df = pd.DataFrame(
        {
            "наименование": [
                "Молоко тестовое 1л",
                "Кефир тестовый 1л",
                "Сок выдуманный неизвестный 1л",
            ],
            "излишек_кол": [1, 0, 0],
            "недостача_кол": [0, 1, 1],
            "излишек_сумма": [10, 0, 0],
            "недостача_сумма": [0, 20, 30],
        }
    )
    out = enrich_dataframe(df, catalog=None, hierarchy=sample_hierarchy)
    assert bool(out.loc[0, "hierarchy_matched"]) is True
    assert out.loc[0, "category_label"] == "Молоко пастеризованное"
    assert bool(out.loc[1, "hierarchy_matched"]) is True
    assert bool(out.loc[2, "hierarchy_matched"]) is False
    assert out.loc[2, "category_label"] == UNMATCHED_LABEL
    um = unmatched_summary(out)
    assert len(um) == 1
    assert "Сок выдуманный" in um.iloc[0]["наименование"]


def test_same_leaf_group_for_related_dairy(sample_hierarchy):
    """Homogeneity for overlap should use template leaf groups, not keyword buckets."""
    a = map_by_master_hierarchy("Молоко тестовое 1л", sample_hierarchy)
    b = map_by_master_hierarchy("Кефир тестовый 1л", sample_hierarchy)
    # Different leaf groups in fixture — must not collapse into one keyword «Молочные»
    assert a["category_group"] != b["category_group"]
    assert a["levels"]["level_1"] == b["levels"]["level_1"] == "Молочная продукция"


@pytest.mark.skipif(not PROD_JSON.exists(), reason="production hierarchy JSON not packaged")
def test_production_hierarchy_loads():
    clear_hierarchy_cache()
    href = load_master_hierarchy(PROD_JSON)
    assert len(href.items) > 1000
    assert href.max_depth >= 3
    # Spot-check a known product from template extraction
    sample_name = href.items[0].name
    hit = href.lookup(sample_name)
    assert hit is not None
    assert hit.leaf_group
    assert hit.hierarchy_path
