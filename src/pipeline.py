# -*- coding: utf-8 -*-
"""High-level analysis pipeline facade."""
from __future__ import annotations

from src.excel.export import build_report, export_report_bytes
from src.models import Config

__all__ = ["Config", "build_report", "export_report_bytes", "run_analysis"]


def run_analysis(cfg: Config) -> str:
    """Run full analysis and return output file path."""
    return build_report(cfg)
