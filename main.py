# -*- coding: utf-8 -*-
"""CLI entrypoint — розничный анализатор (без производства)."""
from __future__ import annotations

import datetime
import logging
import os
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.settings import load_settings
from src import __version__
from src.cli import choose_source_file
from src.models import Config
from src.pipeline import run_analysis

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger(__name__)


def _rebuild_hierarchy(hierarchy_xlsx: str) -> None:
    """Пересобрать data/master_hierarchy.json из шаблона 1С."""
    from scripts.build_master_hierarchy import DEFAULT_CSV, DEFAULT_OUT, extract
    import csv
    import json

    payload = extract(Path(hierarchy_xlsx))
    DEFAULT_OUT.parent.mkdir(parents=True, exist_ok=True)
    DEFAULT_OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    depth = payload["max_hierarchy_depth"]
    fields = ["name", "leaf_group", "hierarchy_path"] + [f"level_{i}" for i in range(depth)] + ["outline_level"]
    with DEFAULT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for it in payload["items"]:
            row = dict(it)
            row["hierarchy_path"] = " / ".join(it["hierarchy_path"])
            w.writerow(row)
    try:
        from src.master_hierarchy import get_master_hierarchy
        get_master_hierarchy.cache_clear()
    except Exception:
        pass
    print(f"      Эталон обновлён: {payload['unique_names_stored']:,} SKU → {DEFAULT_OUT}")


def main() -> None:
    print("=" * 60)
    print(f"  АНАЛИЗАТОР ИНВЕНТАРИЗАЦИЙ РОЗНИЦЫ  v{__version__}")
    print("  (без собственного производства, без проверки цен)")
    print("=" * 60)

    settings = load_settings()
    search_dir = settings.search_dir
    if not search_dir.is_dir():
        search_dir = Path(".").resolve()
    print(f"\nРабочая папка (INVENTORY_DATA_DIR или текущая):\n{search_dir}")

    if not search_dir.is_dir():
        print(f"\nПапка не найдена: {search_dir}")
        try:
            input("\nEnter для выхода...")
        except EOFError:
            pass
        return

    print("\n[1/3] Выберите файл ИНВЕНТАРИЗАЦИЙ (розница без производства):")
    src = choose_source_file(str(search_dir))
    if not src or not os.path.exists(src):
        print("Файл инвентаризации не выбран.")
        try:
            input("\nEnter для выхода...")
        except EOFError:
            pass
        return

    print("\n[2/3] Выберите файл ОПРИХОДОВАНИЯ ИЗЛИШКОВ (закрытие смены, торговый зал):")
    print("      Нужны колонки: «Закрытие смены количество», «Закрытие смены сумма».")
    cap = choose_source_file(str(search_dir))
    if not cap or not os.path.exists(cap):
        print("Файл оприходования не выбран — отчёт без сверки оприходования.")
        cap = None

    print("\n[3/3] Выберите файл ИЕРАРХИИ НОМЕНКЛАТУРЫ (Шаблон_иерархия_безпроизводства.xlsx):")
    print("      При выборе файл будет зафиксирован в data/master_hierarchy.json.")
    hier = choose_source_file(str(search_dir))
    if hier and os.path.exists(hier):
        try:
            print("      Сборка эталона иерархии...")
            _rebuild_hierarchy(hier)
        except Exception as exc:
            print(f"ОШИБКА сборки иерархии: {exc}")
            traceback.print_exc()
            try:
                input("\nEnter для выхода...")
            except EOFError:
                pass
            return
    else:
        embedded = ROOT / "data" / "master_hierarchy.json"
        if embedded.exists():
            print(f"Файл иерархии не выбран — используем уже зафиксированный эталон:\n{embedded}")
        else:
            print("Эталон иерархии отсутствует. Выберите Шаблон_иерархия_безпроизводства.xlsx.")
            try:
                input("\nEnter для выхода...")
            except EOFError:
                pass
            return

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(os.path.dirname(src))
    out = str(out_dir / f"Анализ_розница_{Path(src).stem}_{ts}.xlsx")
    catalog = os.path.join(os.path.dirname(src), settings.catalog_filename)

    cfg = Config(
        input_path=src,
        output_path=out,
        period_days=settings.period_days,
        catalog_path=catalog if os.path.exists(catalog) else None,
        enable_cross_store=settings.enable_cross_store,
        capitalization_path=cap,
        hierarchy_path=hier if hier and os.path.exists(hier) else None,
    )

    try:
        result = run_analysis(cfg)
        print(f"\nИтоговый файл:\n{result}")
    except Exception as exc:
        log.error("ОШИБКА: %s", exc)
        traceback.print_exc()

    try:
        input("\nEnter для выхода...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()
