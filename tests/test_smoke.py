# -*- coding: utf-8 -*-
"""Smoke and unit tests for network inventory analyzer."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.categories import classify_category
from src.metrics import calc_network_summary, percentile_rank
from src.overlap import homogeneous_pair, similarity_score_row
from src.parser import filter_by_period, validate_network_structure
from src.scenarios import assign_scenario
from src.text_normalize import (
    comparable_units,
    infer_unit,
    norm_txt,
    product_key,
    sku_key,
)


def test_norm_txt_basic():
    assert "масло" in norm_txt("Масло Ёлочка!")
    assert norm_txt("Тест") == "тест"


def test_product_and_sku_key():
    assert product_key("Молоко 2.5% 1л") 
    assert sku_key("Молоко") == "молоко"


def test_infer_unit():
    assert infer_unit("Сахар 1 кг") == "кг"
    assert infer_unit("Сок 1 л") == "л"
    assert infer_unit("Печенье упаковка") in ("шт", "кг", "л")


def test_comparable_units():
    assert comparable_units("кг", "г")
    assert comparable_units("л", "мл")
    assert not comparable_units("шт", "кг")


def test_classify_category_keyword_deprecated():
    """Legacy classifier still works but is deprecated and unused in enrich."""
    import warnings

    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        group, method, fb = classify_category("Молоко пастеризованное 2.5%")
        assert group == "Молочные продукты"
        assert method == "keyword"
        assert fb is False
        assert any(issubclass(x.category, DeprecationWarning) for x in w)


def test_homogeneous_pair_same_category():
    src = pd.Series({"category_group": "Молочные продукты", "артикул": "A1", "наименование": "Молоко"})
    dst = pd.Series({"category_group": "Молочные продукты", "артикул": "A2", "наименование": "Кефир"})
    ok, reason = homogeneous_pair(src, dst)
    assert ok is True
    assert reason == "same_category"


def test_homogeneous_pair_rejected():
    src = pd.Series({"category_group": "Мясо и птица", "артикул": "M1", "наименование": "Курица"})
    dst = pd.Series({"category_group": "Напитки", "артикул": "N1", "наименование": "Кола"})
    ok, reason = homogeneous_pair(src, dst)
    assert ok is False
    assert reason == "rejected_category"


def test_percentile_rank():
    s = pd.Series([1.0, 2.0, 3.0])
    r = percentile_rank(s)
    assert len(r) == 3
    assert r.max() <= 100


def test_network_summary_empty_overlap():
    df = pd.DataFrame(
        {
            "излишек_сумма": [100.0],
            "недостача_сумма": [50.0],
            "сумма_учетная": [1000.0],
            "магазин": ["A"],
            "документ": ["D1"],
            "артикул": ["X"],
        }
    )
    summary = calc_network_summary(df, pd.DataFrame())
    assert summary["surplus"] == 100.0
    assert summary["shortage"] == 50.0
    assert summary["overlap"] == 0
    assert summary["clean_shortage"] == 50.0


def test_assign_scenario_small():
    row = pd.Series(
        {
            "наименование": "Печенье простое",
            "единица": "шт",
            "недостача_сумма": 100.0,
            "недостача_кол": 1.0,
            "артикул": "P1",
        }
    )
    sc, *_ = assign_scenario(row, set(), pd.DataFrame())
    assert sc == "СЦ-6"


def test_validate_rejects_report_like_file():
    raw = pd.DataFrame([["Сводный анализ инвентаризаций"] + [None] * 8])
    with pytest.raises(ValueError, match="аналитический файл"):
        validate_network_structure(raw)


def test_filter_by_period():
    import datetime as dt

    df = pd.DataFrame(
        {
            "дата_док": [
                dt.datetime(2026, 7, 1),
                dt.datetime(2026, 7, 20),
                dt.datetime(2026, 6, 1),
            ],
            "val": [1, 2, 3],
        }
    )
    out = filter_by_period(df, period_days=30, end_date=dt.date(2026, 7, 20))
    assert len(out) == 2


def test_import_pipeline_smoke():
    """Smoke: critical packages import without error."""
    from src.pipeline import Config, build_report, run_analysis  # noqa: F401
    from config.settings import load_settings

    s = load_settings()
    assert s.period_days >= 1
