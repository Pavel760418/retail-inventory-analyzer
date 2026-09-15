# -*- coding: utf-8 -*-
"""Integration-ish tests for analysis_service KPI vs Excel (synthetic where possible)."""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.analysis_service import AnalysisResult, QualityReport
from src.models import ParseMeta
from src.ui_helpers import executive_insights, kpi_cards, status_risk_table


def _fake_result() -> AnalysisResult:
    df = pd.DataFrame({
        "магазин": ["Маг1", "Маг1", "Маг2"],
        "документ": ["D1", "D1", "D2"],
        "наименование": ["A", "B", "A"],
        "излишек_сумма": [100.0, 0.0, 0.0],
        "недостача_сумма": [0.0, 200.0, 50.0],
        "излишек_кол": [1, 0, 0],
        "недостача_кол": [0, 2, 1],
        "сумма_учетная": [500, 500, 500],
        "сумма_факт": [600, 300, 450],
        "автор_дока": ["ревизор", "ревизор", "оператор"],
        "hierarchy_matched": [True, True, True],
    })
    summary = {
        "stores": 2, "docs": 2, "sku_lines": 3, "sku_disc": 3,
        "surplus": 100.0, "shortage": 250.0, "net": -150.0,
        "overlap": 0.0, "clean_shortage": 250.0,
        "non_homogeneous_surplus": 100.0,
        "shrinkage_pct": 25.0, "recovery_pct": 0.0,
    }
    store_metrics = pd.DataFrame({
        "магазин": ["Маг1", "Маг2"],
        "чистые_недостачи": [200.0, 50.0],
        "класс": ["C", "B"],
        "store_score": [40, 60],
        "место": [2, 1],
    })
    sku_cross = pd.DataFrame({
        "наименование": ["B", "A"],
        "недостачи": [200.0, 50.0],
        "chronic": [True, False],
    })
    return AnalysisResult(
        df=df,
        meta=ParseMeta(file_name="demo.xlsx", sheet_name="TDSheet"),
        summary=summary,
        store_metrics=store_metrics,
        sku_cross=sku_cross,
        chronic=sku_cross[sku_cross["chronic"]],
        overlap_op=pd.DataFrame(),
        anom_df=pd.DataFrame(),
        anom_sum={"total": 0, "critical": 0, "type1": 0, "type2": 0, "type3": 0},
        author_role=pd.DataFrame(),
        author_store=pd.DataFrame(),
        author_chains=pd.DataFrame(),
        author_sum={"network_dominant": "ревизор", "network_sign": "недостачи", "chains_count": 0, "chains_sum": 0},
        cap_matched=pd.DataFrame(),
        cap_sum={"matched": 0, "matched_shortage": 0, "cap_sum": 0},
        conclusions=[("Масштаб", "ok")],
        quality=QualityReport(after_period_rows=3, passed=True),
        excel_bytes=b"PK\x03\x04demo",
        excel_sheets=5,
        excel_seconds=0.1,
        period_str="01.09.2026 — 30.09.2026",
        source_fingerprint="demo",
    )


def test_kpi_cards_match_summary():
    r = _fake_result()
    cards = kpi_cards(r)
    labels = [c["label"] for c in cards]
    assert "Позиций проверено" in labels
    assert "Недостача, ₽" in labels
    assert "Излишки, ₽" in labels
    shortage_card = next(c for c in cards if c["label"] == "Недостача, ₽")
    assert "250" in shortage_card["value"].replace("\xa0", "")


def test_executive_insights_factual():
    lines = executive_insights(_fake_result())
    assert len(lines) >= 3
    assert any("250" in x or "недостач" in x.lower() for x in lines)


def test_status_table_has_rows():
    tbl = status_risk_table(_fake_result())
    assert len(tbl) >= 4
    assert "Статус" in tbl.columns
