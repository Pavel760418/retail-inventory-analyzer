# -*- coding: utf-8 -*-
"""Embedded master nomenclature hierarchy (etalon from Шаблон_иерархия.xlsx).

Runtime loads data/master_hierarchy.json shipped with the repository.
No per-run upload of hierarchy is required.
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from src.text_normalize import sku_key

log = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_HIERARCHY_PATH = PROJECT_ROOT / "data" / "master_hierarchy.json"

UNMATCHED_LABEL = "Не найдено в справочнике"
UNMATCHED_PREFIX = "__unmatched__:"


@dataclass(frozen=True)
class HierarchyItem:
    """One SKU row from the master hierarchy template."""

    name: str
    leaf_group: str
    hierarchy_path: Tuple[str, ...]
    outline_level: int = 0
    levels: Tuple[str, ...] = field(default_factory=tuple)

    @property
    def path_str(self) -> str:
        return " / ".join(self.hierarchy_path)

    @property
    def category_group(self) -> str:
        """Group used for overlap homogeneity (immediate parent in template)."""
        return self.leaf_group or (self.hierarchy_path[-1] if self.hierarchy_path else UNMATCHED_LABEL)


@dataclass
class MasterHierarchy:
    """In-memory etalon reference."""

    items: List[HierarchyItem]
    by_sku_key: Dict[str, HierarchyItem]
    max_depth: int
    source_file: str = ""
    meta: Dict[str, Any] = field(default_factory=dict)

    def lookup(self, name: str) -> Optional[HierarchyItem]:
        key = sku_key(name)
        if not key:
            return None
        return self.by_sku_key.get(key)

    def map_name(self, name: str) -> Dict[str, Any]:
        """Map inventory SKU name to hierarchy fields for enrichment."""
        hit = self.lookup(name)
        if hit is None:
            key = sku_key(name)
            return {
                "category_group": f"{UNMATCHED_PREFIX}{key}" if key else UNMATCHED_LABEL,
                "category_label": UNMATCHED_LABEL,
                "category_method": "master_unmatched",
                "category_fallback": True,
                "hierarchy_matched": False,
                "hierarchy_path": "",
                "leaf_group": "",
                "levels": {},
            }
        levels = {f"level_{i}": v for i, v in enumerate(hit.hierarchy_path)}
        return {
            "category_group": hit.category_group,
            "category_label": hit.category_group,
            "category_method": "master_hierarchy",
            "category_fallback": False,
            "hierarchy_matched": True,
            "hierarchy_path": hit.path_str,
            "leaf_group": hit.leaf_group,
            "levels": levels,
        }


def _parse_item(raw: Dict[str, Any]) -> HierarchyItem:
    path = raw.get("hierarchy_path") or []
    if isinstance(path, str):
        path = [p.strip() for p in path.split("/") if p.strip()]
    path_t = tuple(str(x) for x in path)
    levels = []
    i = 0
    while f"level_{i}" in raw and raw.get(f"level_{i}") not in (None, ""):
        levels.append(str(raw[f"level_{i}"]))
        i += 1
    if not path_t and levels:
        path_t = tuple(levels)
    return HierarchyItem(
        name=str(raw.get("name", "")).strip(),
        leaf_group=str(raw.get("leaf_group", "") or (path_t[-1] if path_t else "")),
        hierarchy_path=path_t,
        outline_level=int(raw.get("outline_level", 0) or 0),
        levels=tuple(levels) if levels else path_t,
    )


def load_master_hierarchy(path: Optional[Path | str] = None) -> MasterHierarchy:
    """Load etalon hierarchy from JSON. Defaults to packaged data/master_hierarchy.json."""
    p = Path(path) if path else DEFAULT_HIERARCHY_PATH
    if not p.exists():
        raise FileNotFoundError(
            f"Эталонный справочник иерархии не найден: {p}. "
            "Ожидается data/master_hierarchy.json в корне проекта."
        )
    data = json.loads(p.read_text(encoding="utf-8"))
    items_raw = data.get("items") or []
    items: List[HierarchyItem] = []
    by_key: Dict[str, HierarchyItem] = {}
    for raw in items_raw:
        item = _parse_item(raw)
        if not item.name:
            continue
        items.append(item)
        key = sku_key(item.name)
        if key and key not in by_key:
            by_key[key] = item
    max_depth = int(data.get("max_hierarchy_depth") or 0)
    if not max_depth and items:
        max_depth = max((len(it.hierarchy_path) for it in items), default=0)
    log.info(
        "Master hierarchy loaded: %s items, %s keys, depth=%s, source=%s",
        len(items),
        len(by_key),
        max_depth,
        data.get("source_file", p.name),
    )
    return MasterHierarchy(
        items=items,
        by_sku_key=by_key,
        max_depth=max_depth,
        source_file=str(data.get("source_file", p.name)),
        meta={k: v for k, v in data.items() if k != "items"},
    )


@lru_cache(maxsize=4)
def get_master_hierarchy(path: str = "") -> MasterHierarchy:
    """Cached loader for runtime (CLI / Streamlit)."""
    return load_master_hierarchy(path or None)


def map_by_master_hierarchy(name: str, hierarchy: Optional[MasterHierarchy] = None) -> Dict[str, Any]:
    """Public lookup helper used by enrichment."""
    href = hierarchy or get_master_hierarchy()
    return href.map_name(name)


def clear_hierarchy_cache() -> None:
    get_master_hierarchy.cache_clear()
