# -*- coding: utf-8 -*-
"""User-facing validation for inventory uploads (wraps existing parser checks)."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import pandas as pd

from src.parser import REQUIRED_MARKERS, validate_network_structure


@dataclass
class ValidationResult:
    ok: bool
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    sheet_name: Optional[str] = None
    rows_preview: int = 0

    @property
    def user_message(self) -> str:
        if self.ok:
            return "Структура файла распознана."
        return "Файл не прошёл проверку структуры.\n" + "\n".join(f"• {e}" for e in self.errors)


EXPECTED_FORMAT_HINT = (
    "Ожидается исходная выгрузка инвентаризаций (TDSheet), не готовый «Анализ_…».\n"
    "В шапке должны быть: Документ.Подразделение, Количество книжное, "
    "Количество фактическое, Сумма излишек, Сумма недостача.\n"
    "Строки: магазин → «Инвентаризация № от ДД.ММ.ГГГГ [время]» → номенклатура."
)


def validate_uploaded_inventory_bytes(data: bytes, filename: str = "upload.xlsx") -> ValidationResult:
    """Validate inventory workbook from in-memory bytes without changing parse logic."""
    import tempfile
    from pathlib import Path

    if not data:
        return ValidationResult(ok=False, errors=["Файл пустой. Загрузите исходный Excel инвентаризаций."])

    suffix = Path(filename).suffix.lower() or ".xlsx"
    if suffix not in (".xlsx", ".xls"):
        return ValidationResult(
            ok=False,
            errors=[f"Формат «{suffix}» не поддерживается. Используйте .xlsx (рекомендуется)."],
        )

    tmp_path = None
    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(data)
        tmp.flush()
        tmp.close()
        tmp_path = tmp.name

        xl = pd.ExcelFile(tmp_path)
        if not xl.sheet_names:
            return ValidationResult(ok=False, errors=["В файле нет листов."])
        sheet = xl.sheet_names[0]
        df_raw = pd.read_excel(tmp_path, sheet_name=sheet, header=None)
        if df_raw.empty:
            return ValidationResult(ok=False, errors=["Лист пустой."])

        try:
            validate_network_structure(df_raw)
        except ValueError as exc:
            msg = str(exc)
            errors = [msg, EXPECTED_FORMAT_HINT]
            first = " ".join(str(v).lower() for v in df_raw.iloc[:5, :9].fillna("").values.flatten())
            missing = [m for m in REQUIRED_MARKERS if m not in first]
            if missing:
                errors.append("Не найдены маркеры: " + ", ".join(missing))
            return ValidationResult(ok=False, errors=errors, sheet_name=sheet, rows_preview=len(df_raw))

        warnings: List[str] = []
        if len(df_raw) < 10:
            warnings.append("В файле мало строк — проверьте, что выбрана полная выгрузка.")
        return ValidationResult(
            ok=True,
            warnings=warnings,
            sheet_name=sheet,
            rows_preview=len(df_raw),
        )
    except Exception as exc:
        return ValidationResult(
            ok=False,
            errors=[
                f"Не удалось прочитать Excel: {exc}",
                EXPECTED_FORMAT_HINT,
            ],
        )
    finally:
        if tmp_path:
            try:
                Path(tmp_path).unlink(missing_ok=True)
            except OSError:
                pass


def validate_capitalization_bytes(data: bytes, filename: str = "cap.xlsx") -> ValidationResult:
    """Light check that capitalization file has expected shift-close columns."""
    import tempfile
    from pathlib import Path

    if not data:
        return ValidationResult(ok=False, errors=["Файл оприходования пустой."])
    tmp_path = None
    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=Path(filename).suffix or ".xlsx")
        tmp.write(data)
        tmp.flush()
        tmp.close()
        tmp_path = tmp.name
        df = pd.read_excel(tmp_path, header=None, nrows=30)
        blob = " ".join(str(v).lower() for v in df.fillna("").values.flatten())
        need = ["закрытие смены количество", "закрытие смены сумма"]
        missing = [n for n in need if n not in blob]
        if missing:
            return ValidationResult(
                ok=False,
                errors=[
                    "Не найдены колонки оприходования: " + ", ".join(missing),
                    "Нужен файл с колонками «Закрытие смены количество» и «Закрытие смены сумма».",
                ],
            )
        return ValidationResult(ok=True, sheet_name="TDSheet?", rows_preview=len(df))
    except Exception as exc:
        return ValidationResult(ok=False, errors=[f"Не удалось прочитать файл оприходования: {exc}"])
    finally:
        if tmp_path:
            try:
                Path(tmp_path).unlink(missing_ok=True)
            except OSError:
                pass
