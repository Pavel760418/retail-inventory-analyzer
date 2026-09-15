# -*- coding: utf-8 -*-
"""CLI helpers: discover source files and interactive selection."""
from __future__ import annotations

import datetime
import logging
from pathlib import Path
from typing import List

log = logging.getLogger(__name__)

def get_source_files(search_dir: str) -> List[Path]:
    d = Path(search_dir)
    if not d.exists():
        return []
    exclude_prefixes = ("анализ_", "~$")
    exclude_parts = ("автосохран", "autosave", "копия", "copy")
    files = []
    for f in d.glob("*.xlsx"):
        name = f.name.lower()
        if any(name.startswith(p) for p in exclude_prefixes):
            continue
        if any(p in name for p in exclude_parts):
            continue
        files.append(f)
    return sorted(files, key=lambda x: x.stat().st_mtime, reverse=True)


def choose_source_file_tk(search_dir: str) -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Выберите исходный Excel-файл инвентаризаций",
            initialdir=search_dir,
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        root.destroy()
        return path or ""
    except Exception as exc:
        log.warning("tkinter недоступен (%s), использую консольный выбор.", exc)
        return ""


def choose_source_file_console(search_dir: str) -> str:
    files = get_source_files(search_dir)
    if not files:
        print(f"\nВ папке не найдено исходных Excel-файлов:\n{search_dir}")
        return ""
    print("\nВыберите файл для анализа:\n")
    for i, f in enumerate(files, 1):
        mtime = datetime.datetime.fromtimestamp(f.stat().st_mtime).strftime("%d.%m.%Y %H:%M")
        print(f"  {i}. {f.name}   ({round(f.stat().st_size/1024, 1)} KB, {mtime})")
    while True:
        choice = input("\nВведите номер файла: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            return str(files[int(choice) - 1])
        print("Некорректный номер.")


def choose_source_file(search_dir: str) -> str:
    path = choose_source_file_tk(search_dir)
    if path:
        return path
    return choose_source_file_console(search_dir)
