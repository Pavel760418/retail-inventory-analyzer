# -*- coding: utf-8 -*-
"""Validators and UI metrics tests with synthetic fixtures (no real inventory)."""
from __future__ import annotations

import datetime
import sys
import tempfile
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.metrics import calc_network_summary
from src.ui_helpers import kpi_cards
from src.validators import validate_capitalization_bytes, validate_uploaded_inventory_bytes


def _write_minimal_inventory(path: Path) -> None:
    """Synthetic TDSheet-like workbook (DEMO)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "TDSheet"
    headers = [
        "Документ.Подразделение",
        "Количество книжное",
        "Количество фактическое",
        "Количество разница",
        "Сумма книжная (нормативная)",
        "Сумма фактическая (нормативная)",
        "Сумма разница",
        "Сумма излишек",
        "Сумма недостача",
    ]
    ws.append(headers)
    ws.append(["Документ"] + [None] * 8)
    ws.append(["Номенклатура"] + [None] * 8)
    # network / store header row (row 4 in 1-based → index 3)
    ws.append(["Демо-магазин", 10, 9, -1, 1000, 900, -100, 0, -100])
    ws.append(["Инвентаризация ДМ0001 от 01.09.2026 08:00:00", 10, 9, -1, 1000, 900, -100, 0, -100])
    # outline levels for parser
    ws.row_dimensions[5].outlineLevel = 0  # store already set above as row4; doc is row5
    ws.row_dimensions[4].outlineLevel = 0
    ws.row_dimensions[6].outlineLevel = 1
    ws.append(["Товар демо А", 5, 3, -2, 500, 300, -200, 0, -200])
    ws.row_dimensions[6].outlineLevel = 1
    ws.append(["Товар демо Б", 5, 6, 1, 500, 600, 100, 100, 0])
    ws.row_dimensions[7].outlineLevel = 1
    wb.save(path)
    wb.close()


def _write_bad_inventory(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["Сводный анализ инвентаризаций", None])
    wb.save(path)
    wb.close()


def test_validator_rejects_empty():
    r = validate_uploaded_inventory_bytes(b"", "empty.xlsx")
    assert r.ok is False


def test_validator_rejects_report_like(tmp_path: Path):
    p = tmp_path / "bad.xlsx"
    _write_bad_inventory(p)
    r = validate_uploaded_inventory_bytes(p.read_bytes(), "bad.xlsx")
    assert r.ok is False
    assert any("аналитическ" in e.lower() or "заголов" in e.lower() or "маркер" in e.lower() or "сводный" in e.lower() for e in r.errors)


def test_validator_accepts_synthetic(tmp_path: Path):
    p = tmp_path / "ok.xlsx"
    _write_minimal_inventory(p)
    r = validate_uploaded_inventory_bytes(p.read_bytes(), "ok.xlsx")
    assert r.ok is True


def test_cap_validator_missing_columns(tmp_path: Path):
    p = tmp_path / "cap.xlsx"
    wb = Workbook()
    wb.active.append(["A", "B"])
    wb.save(p)
    r = validate_capitalization_bytes(p.read_bytes(), "cap.xlsx")
    assert r.ok is False


def test_network_summary_kpi_consistency():
    df = pd.DataFrame({
        "излишек_сумма": [100.0, 0.0],
        "недостача_сумма": [0.0, 50.0],
        "сумма_учетная": [1000.0, 1000.0],
        "магазин": ["A", "A"],
        "документ": ["D1", "D1"],
        "артикул": ["X", "Y"],
    })
    summary = calc_network_summary(df, pd.DataFrame())
    assert summary["surplus"] == 100.0
    assert summary["shortage"] == 50.0
    assert summary["sku_disc"] == 2
    assert abs((summary["surplus"] + summary["shortage"]) - 150.0) < 1e-9


def test_no_absolute_user_paths_in_runtime():
    """Runtime modules must not hardcode the local Admin Documents path."""
    banned = r"C:\Users\Администратор\Documents"
    for rel in ("app.py", "main.py", "src/analysis_service.py", "src/validators.py", "src/ui_helpers.py"):
        text = (ROOT / rel).read_text(encoding="utf-8")
        assert banned not in text, f"Absolute path found in {rel}"
