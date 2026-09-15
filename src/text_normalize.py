# -*- coding: utf-8 -*-
"""Text normalization, product keys, units, production/weight helpers."""
from __future__ import annotations

import re
from typing import Any, List, Optional, Tuple

STOP_WORDS = {
    "масло", "сливочное", "сливочно", "растительный", "растительно", "жировая", "жиры",
    "пищевые", "продукт", "товар", "традиционное", "традиц", "крестьянское", "крестьян",
    "сладкосливочн", "топленое", "топленая", "топленый", "шоколадное", "соленое",
    "для", "и", "в", "на", "произ", "п/ф", "пф",
}

WEIGHT_KEYWORDS = [
    "яблок", "банан", "груш", "апельсин", "фрукт", "цитрус", "картоф", "лук", "огур",
    "помидор", "томат", "капуст", "морков", "зелень", "укроп", "петруш", "салат",
    "мясо", "говядин", "свинин", "куриц", "бедро", "филе", "рыба", "селедк", "горбуш",
    "минтай", "лосос", "сыр", "колбас", "сосиск", "ветчин", "фарш", "котлет", "весовой",
    "развес", "на развес",
]

PROD_KEYWORDS = [
    "п/ф", "пф ", "пф.", "полуфабрикат", "полуфабр", "сырье", "сырьё", "заготовка",
    "для производства", "для кухни", "кухня", "цех", "фарш", "маринад", "бульон", "ттк", "акп",
]
PROD_STRONG_WORDS = [
    "фарш", "котлетная масса", "полуфабрикат", "полуфабр", "п/ф", "сырье", "сырьё",
    "заготовка", "маринад", "бульон",
]

UNIT_PATTERNS = [
    (re.compile(r"\b(кг|г|гр|г\.)\b", re.I), "кг"),
    (re.compile(r"\b(л|мл|ml)\b", re.I), "л"),
    (re.compile(r"\b(шт|уп|пач|пак|ст/б|бут)\b", re.I), "шт"),
]

def norm_txt(s: Any) -> str:
    s = str(s).lower().replace("ё", "е")
    s = re.sub(r"[^a-zа-я0-9%.,\s-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_tokens(name: str) -> List[str]:
    n = norm_txt(name)
    n = re.sub(r"\d+(?:[.,]\d+)?\s*(кг|г|гр|мл|л|шт|уп|пачк|пак|ст/б)", " ", n)
    toks = re.findall(r"[a-zа-я0-9%]+", n)
    return [t for t in toks if t not in STOP_WORDS and len(t) > 2]


def product_key(name: str, n: int = 3) -> str:
    toks = clean_tokens(name)
    return " ".join(toks[:n])


def sku_key(name: str) -> str:
    return norm_txt(name)


def infer_unit(name: str) -> str:
    n = norm_txt(name)
    for pat, unit in UNIT_PATTERNS:
        if pat.search(n):
            return unit
    return "шт"


def extract_weight(name: str) -> Tuple[Optional[float], Optional[str]]:
    n = norm_txt(name)
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(кг|г|гр|мл|л)", n)
    if not m:
        return None, None
    val = float(m.group(1).replace(",", "."))
    unit = m.group(2)
    if unit in ("г", "гр"):
        return val, "г"
    if unit == "кг":
        return val * 1000, "г"
    if unit == "мл":
        return val, "мл"
    if unit == "л":
        return val * 1000, "мл"
    return val, unit


def comparable_units(u1: str, u2: str) -> bool:
    u1, u2 = norm_txt(u1), norm_txt(u2)
    if u1 == u2:
        return True
    groups = [{"шт"}, {"кг", "гр", "г"}, {"л", "мл"}]
    return any(u1 in g and u2 in g for g in groups)


def price_per_unit(qty: float, sm: float) -> float:
    if qty <= 0:
        return 0.0
    return sm / qty


def is_production(name: str, unit: str = "") -> bool:
    n = norm_txt(name)
    u = norm_txt(unit)
    strong = any(kw in n for kw in PROD_STRONG_WORDS)
    keyword = any(kw in n for kw in PROD_KEYWORDS)
    return strong or (keyword and u in ("кг", "гр", "л", "мл"))


def is_weight_item(name: str, unit: str) -> bool:
    return norm_txt(unit) == "кг" or any(kw in name.lower() for kw in WEIGHT_KEYWORDS)


def is_pack_peresor(name1: str, name2: str) -> bool:
    def base(n):
        n = n.lower()
        n = re.sub(r"\d+[,.]?\d*\s*(г|гр|мл|л|кг|шт|уп|пачк)", "", n)
        return re.sub(r"\s+", " ", n).strip()
    return base(name1) == base(name2) and name1.lower() != name2.lower()
