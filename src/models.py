# -*- coding: utf-8 -*-
"""Domain dataclasses."""
from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from config.settings import PERIOD_DAYS


@dataclass
class Config:
    """Run configuration for a single analysis job."""

    input_path: str
    output_path: str
    period_days: int = PERIOD_DAYS
    end_date: Optional[datetime.date] = None
    catalog_path: Optional[str] = None
    enable_cross_store: bool = False
    capitalization_path: Optional[str] = None
    hierarchy_path: Optional[str] = None


@dataclass
class ParseMeta:
    """Metadata collected while parsing a network inventory export."""

    file_name: str
    sheet_name: str
    network_total_row: Optional[Dict[str, Any]] = None
    stores: List[str] = field(default_factory=list)
    doc_count: int = 0
    sku_count: int = 0
    date_min: Optional[datetime.date] = None
    date_max: Optional[datetime.date] = None
