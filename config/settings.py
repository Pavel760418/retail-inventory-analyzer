# -*- coding: utf-8 -*-
"""Runtime settings for network inventory analyzer.

Absolute local paths are never required. Defaults are relative / env-based.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Sequence, Set

# Analysis defaults (business parameters — preserved from v3.0)
PERIOD_DAYS = 30
MIN_OVERLAP_SCORE = 20
CHRONIC_MIN_STORES = 2
CHRONIC_MIN_DOCS = 3
TOP_N_NETWORK = 50
TOP_N_STORE = 10
EXCLUDE_STORE_KEYWORDS = ("итого",)
NON_RETAIL_STORES = frozenset({"РЦ", "Фабрика-кухня"})

# Scenario thresholds (rubles)
SCENARIO_LARGE_SHORTAGE = 50_000
SCENARIO_MEDIUM_SHORTAGE = 10_000
SCENARIO_WEIGHT_LOSS = 1_000

# Performance guardrails for pair matching
PAIR_VOLUME_LIMIT = 2_500_000
PAIR_HEAD_LIMIT = 1_200

# Overlap export limit
OVERLAP_EXPORT_ROWS = 500

# Author classification: exact wall-clock time of inventory document
AUTHOR_REVISOR_TIME = (8, 0, 0)  # 08:00:00 = ревизор


@dataclass
class Settings:
    """Application settings loaded from environment / defaults."""

    period_days: int = PERIOD_DAYS
    min_overlap_score: int = MIN_OVERLAP_SCORE
    chronic_min_stores: int = CHRONIC_MIN_STORES
    chronic_min_docs: int = CHRONIC_MIN_DOCS
    top_n_network: int = TOP_N_NETWORK
    top_n_store: int = TOP_N_STORE
    enable_cross_store: bool = False
    default_search_dir: Optional[str] = None
    catalog_filename: str = "nomenclature.csv"
    exclude_store_keywords: Sequence[str] = field(
        default_factory=lambda: EXCLUDE_STORE_KEYWORDS
    )
    non_retail_stores: Set[str] = field(
        default_factory=lambda: set(NON_RETAIL_STORES)
    )

    @property
    def search_dir(self) -> Path:
        raw = self.default_search_dir or os.environ.get("INVENTORY_DATA_DIR", ".")
        return Path(raw).expanduser().resolve()


def load_settings() -> Settings:
    """Load settings from environment variables (optional .env via python-dotenv)."""
    try:
        from dotenv import load_dotenv

        load_dotenv()
    except ImportError:
        pass

    return Settings(
        period_days=int(os.environ.get("PERIOD_DAYS", PERIOD_DAYS)),
        min_overlap_score=int(os.environ.get("MIN_OVERLAP_SCORE", MIN_OVERLAP_SCORE)),
        enable_cross_store=os.environ.get("ENABLE_CROSS_STORE", "0").strip()
        in ("1", "true", "True", "yes", "YES"),
        default_search_dir=os.environ.get("INVENTORY_DATA_DIR") or None,
    )
