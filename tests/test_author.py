# -*- coding: utf-8 -*-
"""Tests for revisor/operator author classification and chains."""
from __future__ import annotations

import datetime

import pandas as pd

from src.author import (
    AUTHOR_OPERATOR,
    AUTHOR_REVISOR,
    AUTHOR_UNKNOWN,
    classify_author,
    detect_author_chains,
    parse_time_flexible,
)


def test_classify_revisor_exact_080000():
    assert classify_author(time_only=datetime.time(8, 0, 0)) == AUTHOR_REVISOR
    assert classify_author(datetime.datetime(2026, 7, 6, 8, 0, 0)) == AUTHOR_REVISOR


def test_classify_operator_other_times():
    assert classify_author(time_only=datetime.time(21, 0, 0)) == AUTHOR_OPERATOR
    assert classify_author(time_only=datetime.time(10, 53, 34)) == AUTHOR_OPERATOR
    assert classify_author(time_only=datetime.time(8, 0, 1)) == AUTHOR_OPERATOR
    assert classify_author(time_only=datetime.time(7, 59, 59)) == AUTHOR_OPERATOR


def test_parse_flexible_time_formats():
    assert parse_time_flexible("8:00:00") == datetime.time(8, 0, 0)
    assert parse_time_flexible("08:00:00") == datetime.time(8, 0, 0)
    assert parse_time_flexible("8:00") == datetime.time(8, 0, 0)
    assert parse_time_flexible("21:00:00") == datetime.time(21, 0, 0)
    assert parse_time_flexible("10:53:34") == datetime.time(10, 53, 34)
    assert classify_author(time_only=parse_time_flexible("8:00:00")) == AUTHOR_REVISOR


def test_unknown_when_no_time():
    assert classify_author(None) == AUTHOR_UNKNOWN
    assert classify_author("") == AUTHOR_UNKNOWN
    assert classify_author(time_only=None) == AUTHOR_UNKNOWN


def test_detect_author_chains_operator_then_revisor():
    df = pd.DataFrame([
        {
            "магазин": "Автодом",
            "документ": "Инвентаризация АВ1 от 01.07.2026 21:00:00",
            "дата_док": datetime.datetime(2026, 7, 1, 21, 0, 0),
            "автор_дока": AUTHOR_OPERATOR,
            "sku_key": "сыр моцарелла",
            "наименование": "Сыр Моцарелла",
            "category_group": "Сыры",
            "излишек_сумма": 0,
            "недостача_сумма": 1000,
            "излишек_кол": 0,
            "недостача_кол": 2,
        },
        {
            "магазин": "Автодом",
            "документ": "Инвентаризация АВ2 от 06.07.2026 8:00:00",
            "дата_док": datetime.datetime(2026, 7, 6, 8, 0, 0),
            "автор_дока": AUTHOR_REVISOR,
            "sku_key": "сыр моцарелла",
            "наименование": "Сыр Моцарелла",
            "category_group": "Сыры",
            "излишек_сумма": 800,
            "недостача_сумма": 0,
            "излишек_кол": 1,
            "недостача_кол": 0,
        },
    ])
    chains = detect_author_chains(df)
    assert len(chains) == 1
    assert chains.iloc[0]["инициатор"] == AUTHOR_OPERATOR
    assert chains.iloc[0]["сумма_цепочки"] == 800
    assert "оператор недостача" in chains.iloc[0]["тип_цепочки"]
