# -*- coding: utf-8 -*-
"""Analysis orchestration for Streamlit — single-pass UI metrics; Excel on demand."""
from __future__ import annotations

import datetime
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd

from src.anomalies import anomalies_summary, detect_book_sum_anomalies
from src.author import (
    annotate_overlap_authors,
    author_network_summary,
    calc_author_role_stats,
    calc_author_store_stats,
    detect_author_chains,
    enrich_author_columns,
    enrich_store_metrics_with_authors,
)
from src.capitalization import (
    capitalization_summary,
    enrich_store_metrics_with_capitalization,
    match_capitalization_to_inventory,
    parse_capitalization_excel,
)
from src.catalog import enrich_dataframe, load_catalog, unmatched_summary
from src.excel.export import build_report
from src.metrics import (
    calc_network_summary,
    calc_sku_cross,
    calc_store_metrics,
    generate_conclusions,
)
from src.models import Config, ParseMeta
from src.overlap import build_analytical_cross_store, build_operational_overlaps
from src.parser import filter_by_period, parse_network_excel
from src.text_normalize import sku_key as make_sku_key


ProgressCb = Optional[Callable[[float, str], None]]


@dataclass
class QualityReport:
    loaded_rows: int = 0
    after_period_rows: int = 0
    excluded_by_period: int = 0
    stores: int = 0
    docs: int = 0
    hierarchy_unmatched_rows: int = 0
    hierarchy_unmatched_names: int = 0
    anomaly_total: int = 0
    anomaly_critical: int = 0
    empty_name_rows: int = 0
    duplicate_keys: int = 0
    passed: bool = True
    issues: List[Dict[str, str]] = field(default_factory=list)


@dataclass
class AnalysisResult:
    df: pd.DataFrame
    meta: ParseMeta
    summary: Dict[str, Any]
    store_metrics: pd.DataFrame
    sku_cross: pd.DataFrame
    chronic: pd.DataFrame
    overlap_op: pd.DataFrame
    anom_df: pd.DataFrame
    anom_sum: Dict[str, Any]
    author_role: pd.DataFrame
    author_store: pd.DataFrame
    author_chains: pd.DataFrame
    author_sum: Dict[str, Any]
    cap_matched: pd.DataFrame
    cap_sum: Dict[str, Any]
    conclusions: List[Tuple[str, str]]
    quality: QualityReport
    excel_bytes: bytes
    excel_sheets: int
    excel_seconds: float
    period_str: str
    source_fingerprint: str
    excel_pending: bool = False
    # Kept only in session for lazy Excel (not required for KPI display)
    _inv_bytes: bytes = field(default=b"", repr=False)
    _cap_bytes: bytes = field(default=b"", repr=False)
    _inv_name: str = ""
    _cap_name: str = ""
    _period_days: int = 30
    _end_date: Optional[datetime.date] = None
    _enable_cross_store: bool = False


def _fingerprint(*names: str) -> str:
    return "|".join(str(n) for n in names)


def _progress(cb: ProgressCb, value: float, text: str) -> None:
    if cb:
        cb(value, text)


def _build_quality(df_raw_len: int, df: pd.DataFrame, um: pd.DataFrame, anom_sum: Dict) -> QualityReport:
    q = QualityReport()
    q.loaded_rows = int(df_raw_len)
    q.after_period_rows = int(len(df))
    q.excluded_by_period = max(df_raw_len - len(df), 0)
    q.stores = int(df["магазин"].nunique()) if not df.empty else 0
    q.docs = int(df["документ"].nunique()) if not df.empty else 0
    q.hierarchy_unmatched_rows = int((~df["hierarchy_matched"]).sum()) if "hierarchy_matched" in df.columns else 0
    q.hierarchy_unmatched_names = int(len(um)) if um is not None else 0
    q.anomaly_total = int(anom_sum.get("total", 0))
    q.anomaly_critical = int(anom_sum.get("critical", 0))
    q.empty_name_rows = int(df["наименование"].isna().sum()) if "наименование" in df.columns else 0
    if {"магазин", "документ", "наименование"}.issubset(df.columns):
        q.duplicate_keys = int(df.duplicated(subset=["магазин", "документ", "наименование"]).sum())

    if q.empty_name_rows:
        q.issues.append({
            "причина": "Пустые наименования",
            "количество": str(q.empty_name_rows),
            "исправление": "Проверьте строки номенклатуры в исходной выгрузке 1С.",
        })
    if q.hierarchy_unmatched_rows:
        q.issues.append({
            "причина": "SKU вне эталонной иерархии",
            "количество": f"{q.hierarchy_unmatched_rows} строк / {q.hierarchy_unmatched_names} имён",
            "исправление": "Обновите шаблон иерархии или сверьте наименование с справочником.",
        })
    if q.anomaly_critical:
        q.issues.append({
            "причина": "Критические аномалии книжных сумм (тип 2)",
            "количество": str(q.anomaly_critical),
            "исправление": "Заполните книжную нормативную и фактическую суммы в 1С.",
        })
    q.passed = q.empty_name_rows == 0
    return q


def run_ui_analysis(
    inventory_path: str,
    capitalization_path: str,
    *,
    period_days: int = 30,
    end_date: Optional[datetime.date] = None,
    enable_cross_store: bool = False,
    source_names: Optional[Tuple[str, str]] = None,
    inv_bytes: bytes = b"",
    cap_bytes: bytes = b"",
    progress_cb: ProgressCb = None,
) -> AnalysisResult:
    """Single-pass metrics for Streamlit UI. Does NOT call build_report (avoids 2× work)."""
    _progress(progress_cb, 0.04, "AI-агент открывает файл инвентаризации…")
    df, meta = parse_network_excel(inventory_path)
    raw_len = int(meta.sku_count or len(df))
    _progress(progress_cb, 0.10, f"AI-агент загрузил {raw_len:,} строк, размечает авторов документов…")
    if "автор_дока" not in df.columns:
        df = enrich_author_columns(df)
    if "sku_key" not in df.columns:
        df["sku_key"] = df["наименование"].map(make_sku_key)

    _progress(progress_cb, 0.16, "AI-агент фильтрует документы по выбранному периоду…")
    end = end_date or meta.date_max
    df = filter_by_period(df, period_days, end)
    if df.empty:
        raise ValueError(
            "После фильтрации по периоду не осталось данных. "
            "Увеличьте период или проверьте даты документов."
        )

    p_start = df.attrs.get("period_start", meta.date_min)
    p_end = df.attrs.get("period_end", meta.date_max)
    period_str = f"{p_start.strftime('%d.%m.%Y')} — {p_end.strftime('%d.%m.%Y')}"

    _progress(progress_cb, 0.24, "AI-агент сопоставляет номенклатуру с эталоном иерархии…")
    df = enrich_dataframe(df, load_catalog(None))
    um = unmatched_summary(df)

    _progress(progress_cb, 0.34, "AI-агент ищет однородные пересорты (перекрытия)…")
    overlap_op = build_operational_overlaps(df)
    _progress(progress_cb, 0.42, "AI-агент привязывает пересорты к ролям ревизор/оператор…")
    overlap_op = annotate_overlap_authors(overlap_op, df)
    # Cross-store is expensive (pairwise) — only if explicitly enabled
    if enable_cross_store:
        _progress(progress_cb, 0.48, "AI-агент считает аналитику между магазинами (медленно)…")
        _ = build_analytical_cross_store(df)

    _progress(progress_cb, 0.54, "AI-агент считает рейтинг магазинов и топы SKU…")
    store_metrics = calc_store_metrics(df, overlap_op)
    sku_cross = calc_sku_cross(df)
    chronic = sku_cross[sku_cross["chronic"]]
    summary = calc_network_summary(df, overlap_op)

    _progress(progress_cb, 0.62, "AI-агент проверяет аномалии книжных сумм…")
    anom_df = detect_book_sum_anomalies(df)
    anom_sum = anomalies_summary(anom_df, total_rows=len(df))
    _progress(progress_cb, 0.70, "AI-агент сравнивает вклад ревизоров и операторов…")
    author_role = calc_author_role_stats(df, overlap_op)
    author_store = calc_author_store_stats(df, overlap_op)
    _progress(progress_cb, 0.76, "AI-агент строит цепочки пересортов ревизор↔оператор…")
    author_chains = detect_author_chains(df)
    author_sum = author_network_summary(author_role, author_store, author_chains)
    store_metrics = enrich_store_metrics_with_authors(store_metrics, author_store)

    _progress(progress_cb, 0.84, "AI-агент сверяет оприходование излишков закрытия смены…")
    cap_raw = parse_capitalization_excel(capitalization_path)
    cap_matched = match_capitalization_to_inventory(cap_raw, df)
    cap_sum = capitalization_summary(cap_matched)
    store_metrics = enrich_store_metrics_with_capitalization(store_metrics, cap_matched)

    _progress(progress_cb, 0.90, "AI-агент формулирует выводы для руководителя…")
    conclusions = generate_conclusions(
        summary, store_metrics, chronic,
        cap_summary=cap_sum,
        anom_summary=anom_sum,
        author_summary=author_sum,
    )
    quality = _build_quality(raw_len, df, um, anom_sum)

    names = source_names or (meta.file_name, Path(capitalization_path).name)
    fp = _fingerprint(*names, period_str, str(period_days), str(bool(enable_cross_store)))

    _progress(progress_cb, 0.96, "AI-агент готовит BI-дашборд с метриками и топами…")
    return AnalysisResult(
        df=df,
        meta=meta,
        summary=summary,
        store_metrics=store_metrics,
        sku_cross=sku_cross,
        chronic=chronic,
        overlap_op=overlap_op,
        anom_df=anom_df,
        anom_sum=anom_sum,
        author_role=author_role,
        author_store=author_store,
        author_chains=author_chains,
        author_sum=author_sum,
        cap_matched=cap_matched,
        cap_sum=cap_sum,
        conclusions=conclusions,
        quality=quality,
        excel_bytes=b"",
        excel_sheets=0,
        excel_seconds=0.0,
        period_str=period_str,
        source_fingerprint=fp,
        excel_pending=True,
        _inv_bytes=inv_bytes,
        _cap_bytes=cap_bytes,
        _inv_name=names[0] if names else "inventory.xlsx",
        _cap_name=names[1] if len(names) > 1 else "cap.xlsx",
        _period_days=period_days,
        _end_date=end_date,
        _enable_cross_store=enable_cross_store,
    )


def ensure_excel_bytes(result: AnalysisResult, progress_cb: ProgressCb = None) -> AnalysisResult:
    """Build Excel once (same build_report as CLI). Called lazily from Excel tab."""
    if result.excel_bytes and not result.excel_pending:
        return result
    if not result._inv_bytes or not result._cap_bytes:
        raise ValueError("Нет исходных файлов в сессии для формирования Excel. Запустите анализ снова.")

    _progress(progress_cb, 0.1, "AI-агент формирует Excel-отчёт…")
    inv_path = None
    cap_path = None
    out_path = None
    t0 = time.perf_counter()
    try:
        inv_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=Path(result._inv_name).suffix or ".xlsx")
        inv_tmp.write(result._inv_bytes)
        inv_tmp.flush()
        inv_tmp.close()
        inv_path = inv_tmp.name

        cap_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=Path(result._cap_name).suffix or ".xlsx")
        cap_tmp.write(result._cap_bytes)
        cap_tmp.flush()
        cap_tmp.close()
        cap_path = cap_tmp.name

        out_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        out_path = out_tmp.name
        out_tmp.close()

        cfg = Config(
            input_path=inv_path,
            output_path=out_path,
            period_days=result._period_days,
            end_date=result._end_date,
            enable_cross_store=result._enable_cross_store,
            capitalization_path=cap_path,
        )
        _progress(progress_cb, 0.35, "AI-агент собирает листы отчёта (один проход)…")
        build_report(cfg)
        excel_bytes = Path(out_path).read_bytes()
        from openpyxl import load_workbook
        wb = load_workbook(out_path, read_only=True)
        n_sheets = len(wb.sheetnames)
        wb.close()
    finally:
        for p in (inv_path, cap_path, out_path):
            if p:
                try:
                    Path(p).unlink(missing_ok=True)
                except OSError:
                    pass

    elapsed = time.perf_counter() - t0
    result.excel_bytes = excel_bytes
    result.excel_sheets = n_sheets
    result.excel_seconds = elapsed
    result.excel_pending = False
    # Drop heavy copies after excel built to free session memory
    result._inv_bytes = b""
    result._cap_bytes = b""
    _progress(progress_cb, 1.0, "Excel готов")
    return result


def run_full_analysis(
    inventory_path: str,
    capitalization_path: str,
    *,
    period_days: int = 30,
    end_date: Optional[datetime.date] = None,
    enable_cross_store: bool = False,
    source_names: Optional[Tuple[str, str]] = None,
) -> AnalysisResult:
    """Backward-compatible: UI metrics + Excel in one call (slower). Prefer run_ui_analysis."""
    inv_bytes = Path(inventory_path).read_bytes()
    cap_bytes = Path(capitalization_path).read_bytes()
    result = run_ui_analysis(
        inventory_path,
        capitalization_path,
        period_days=period_days,
        end_date=end_date,
        enable_cross_store=enable_cross_store,
        source_names=source_names,
        inv_bytes=inv_bytes,
        cap_bytes=cap_bytes,
    )
    return ensure_excel_bytes(result)
