# -*- coding: utf-8 -*-
"""Release 4 tests: capitalization parse/match, store isolation."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.capitalization import (
    STATUS_MATCHED,
    STATUS_MATCHED_SHORTAGE,
    STATUS_UNCERTAIN_STORE,
    STATUS_UNMATCHED_NAME,
    STATUS_UNMATCHED_STORE,
    capitalization_summary,
    enrich_store_metrics_with_capitalization,
    map_warehouse_to_store,
    match_capitalization_to_inventory,
    parse_capitalization_excel,
)
from src.metrics import calc_store_metrics, generate_conclusions


def _inv_frame():
    return pd.DataFrame(
        {
            "магазин": ["Гаджиева, 198", "Гаджиева, 198", "Автодом", "Автодом", "БКК"],
            "документ": ["D1", "D1", "D2", "D2", "D3"],
            "наименование": [
                "Пирожок картофель/грибы",
                "Булочка с Маком",
                "Соус соевый 30г",
                "Чужой товар Автодом",
                "Пицца 4 сыра",
            ],
            "излишек_кол": [0, 2, 1, 0, 0],
            "излишек_сумма": [0, 50, 120, 0, 0],
            "недостача_кол": [3, 0, 0, 1, 2],
            "недостача_сумма": [200, 0, 0, 40, 500],
            "сумма_учетная": [1000, 1000, 1000, 1000, 1000],
            "sku_key": [
                "пирожок картофель грибы",
                "булочка с маком",
                "соус соевый 30г",
                "чужой товар автодом",
                "пицца 4 сыра",
            ],
        }
    )


def _cap_frame():
    return pd.DataFrame(
        {
            "склад": [
                "Склад 104 (Гаджиева, 198)",
                "Склад 104 (Гаджиева, 198)",
                "Склад Автодом",
                "Склад Автодом",
                "Склад НеизвестныйXYZ",
            ],
            "наименование": [
                "Пирожок картофель/грибы",  # shortage in same store
                "Булочка с Маком",  # surplus only
                "Соус соевый 30г",  # matched surplus
                "Несуществующий товар XYZ",  # unmatched name
                "Что угодно",  # unmatched store
            ],
            "оп_количество": [8, 27, 76, 1, 1],
            "оп_сумма": [311.7, 228.96, 2280.0, 10.0, 5.0],
            "sku_key": [
                "пирожок картофель грибы",
                "булочка с маком",
                "соус соевый 30г",
                "несуществующий товар xyz",
                "что угодно",
            ],
            "name_exact": [
                "Пирожок картофель/грибы",
                "Булочка с Маком",
                "Соус соевый 30г",
                "Несуществующий товар XYZ",
                "Что угодно",
            ],
        }
    )


def test_map_warehouse_unique_by_address():
    store, status = map_warehouse_to_store(
        "Склад 104 (Гаджиева, 198)",
        ["Гаджиева, 198", "Автодом", "БКК"],
    )
    assert status == "mapped"
    assert store == "Гаджиева, 198"


def test_map_warehouse_by_name_token():
    store, status = map_warehouse_to_store("Склад Автодом", ["Гаджиева, 198", "Автодом", "БКК"])
    assert status == "mapped"
    assert store == "Автодом"


def test_map_warehouse_unmatched():
    store, status = map_warehouse_to_store("Склад НеизвестныйXYZ", ["Автодом", "БКК"])
    assert store is None
    assert status == STATUS_UNMATCHED_STORE


def test_no_cross_store_name_match():
    """Same SKU name must not match across different stores."""
    inv = _inv_frame()
    # Put identical name only on БКК; capitalization comes from Автодом warehouse
    cap = pd.DataFrame(
        {
            "склад": ["Склад Автодом"],
            "наименование": ["Пицца 4 сыра"],  # exists only on БКК in inventory
            "оп_количество": [1],
            "оп_сумма": [100],
            "sku_key": ["пицца 4 сыра"],
            "name_exact": ["Пицца 4 сыра"],
        }
    )
    out = match_capitalization_to_inventory(cap, inv)
    assert len(out) == 1
    assert out.iloc[0]["статус"] == STATUS_UNMATCHED_NAME
    assert out.iloc[0]["магазин"] == "Автодом"


def test_exact_match_and_shortage_flag():
    out = match_capitalization_to_inventory(_cap_frame(), _inv_frame())
    by_name = {r["наименование"]: r for _, r in out.iterrows()}
    assert by_name["Пирожок картофель/грибы"]["статус"] == STATUS_MATCHED_SHORTAGE
    assert by_name["Пирожок картофель/грибы"]["связь_с_недостачей"] == "ДА"
    assert by_name["Булочка с Маком"]["статус"] == STATUS_MATCHED
    assert by_name["Соус соевый 30г"]["статус"] == STATUS_MATCHED
    assert by_name["Несуществующий товар XYZ"]["статус"] == STATUS_UNMATCHED_NAME
    assert by_name["Что угодно"]["статус"] == STATUS_UNMATCHED_STORE


def test_capitalization_summary_and_store_enrich():
    out = match_capitalization_to_inventory(_cap_frame(), _inv_frame())
    summary = capitalization_summary(out)
    assert summary["cap_rows"] == 5
    assert summary["matched_shortage"] >= 1
    assert summary["unmatched_name"] >= 1

    # minimal store metrics
    sm = pd.DataFrame(
        {
            "магазин": ["Гаджиева, 198", "Автодом", "БКК"],
            "класс": ["A", "B", "C"],
            "store_score": [90, 70, 40],
            "излишки": [1, 1, 1],
            "недостачи": [1, 1, 1],
            "сальдо": [0, 0, 0],
            "перекрытие": [0, 0, 0],
            "чистые_недостачи": [1, 1, 1],
            "shrinkage_pct": [1, 1, 1],
            "recovery_pct": [0, 0, 0],
            "документов": [1, 1, 1],
            "место": [1, 2, 3],
        }
    )
    enriched = enrich_store_metrics_with_capitalization(sm, out)
    assert "оприходование_сумма" in enriched.columns
    gad = enriched[enriched["магазин"] == "Гаджиева, 198"].iloc[0]
    assert float(gad["оприходование_сумма"]) > 0


def test_conclusions_include_capitalization():
    lines = generate_conclusions(
        {
            "stores": 1,
            "docs": 1,
            "sku_disc": 1,
            "surplus": 1,
            "shortage": 1,
            "net": 0,
            "overlap": 0,
            "recovery_pct": 0,
            "clean_shortage": 1,
            "non_homogeneous_surplus": 0,
            "shrinkage_pct": 0,
        },
        pd.DataFrame({"магазин": ["A"]}),
        pd.DataFrame(),
        cap_summary={"warehouses": 2, "cap_rows": 10, "cap_sum": 1000, "matched": 5,
                     "matched_shortage": 2, "unmatched_name": 2, "unmatched_store": 1, "uncertain_store": 0},
    )
    titles = [t for t, _ in lines]
    assert "Оприходование излишков (закрытие смены)" in titles


def test_detect_excel_container_xls_disguised_as_xlsx():
    import os
    from src.capitalization import _detect_excel_container

    root_raw = os.environ.get("RETAIL_DATA_DIR", "").strip()
    root = Path(root_raw) if root_raw else None
    if root is None or not root.is_dir():
        pytest.skip("RETAIL_DATA_DIR not set — skip optional legacy xls probe")
    ole = next(
        (
            p
            for p in root.glob("*.xls*")
            if not p.name.startswith("~$") and p.read_bytes()[:2] != b"PK"
        ),
        None,
    )
    if ole is None:
        pytest.skip("no legacy .xls capitalization sample in data dir")
    assert _detect_excel_container(ole) == "xls"
    df = parse_capitalization_excel(str(ole))
    assert not df.empty
    assert df["склад"].astype(str).str.len().gt(0).any()


def test_parse_real_capitalization_file_if_present():
    # Optional local smoke only — path via env, never committed
    import os
    raw = os.environ.get("RETAIL_CAP_SAMPLE_XLSX", "").strip()
    path = Path(raw) if raw else None
    if not path or not path.exists():
        pytest.skip("real capitalization file not available")
    df = parse_capitalization_excel(str(path))
    assert not df.empty
    assert {"склад", "наименование", "оп_количество", "оп_сумма"}.issubset(df.columns)
    assert df["склад"].nunique() >= 2
    assert (df["склад"].astype(str).str.len() > 0).all()
