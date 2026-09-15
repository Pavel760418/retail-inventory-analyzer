# -*- coding: utf-8 -*-
"""Action-plan scenario assignment for shortage rows."""
from __future__ import annotations

from config.settings import (
    SCENARIO_LARGE_SHORTAGE,
    SCENARIO_MEDIUM_SHORTAGE,
    SCENARIO_WEIGHT_LOSS,
)
from src.text_normalize import is_pack_peresor, is_production, is_weight_item

def assign_scenario(row, peresor_arts, dfp):
    name = row["наименование"]
    unit = row.get("единица", "")
    amount = float(row.get("недостача_сумма", 0))
    qty = float(row.get("недостача_кол", 0))
    art = row.get("артикул", "")
    is_prod = is_production(name, unit)
    is_w = is_weight_item(name, unit)
    in_p = art in peresor_arts

    if is_prod:
        return ("СЦ-1", "Производственное сырьё / ПФ",
                "1. Проверить ТТК.\n2. Провести АКП.\n3. Сверить расход с нормами.\n4. Проверить списания.", 2)
    if in_p and dfp is not None and not dfp.empty:
        pr = dfp[dfp["недостача_арт"] == art]
        if not pr.empty:
            best = pr.sort_values(["score", "перекрытие_сум"], ascending=[False, False]).iloc[0]
            pn = best["излишек_товар"]
            if is_weight_item(name, unit) and is_weight_item(pn, best.get("единица", unit)):
                return ("СЦ-2А", "Весовой пересорт",
                        f'Пара: "{pn}".\n1. Контроль взвешивания.\n2. Акт пересортицы.', 3)
            if is_pack_peresor(name, pn):
                return ("СЦ-2Б", "Упаковочный пересорт",
                        f'Пара: "{pn}".\n1. Раздельное хранение.\n2. Сверка штрихкода.', 3)
            return ("СЦ-2В", "Пересорт — разновидность",
                    f'Пара: "{pn}".\n1. Идентификация позиции.\n2. Акт пересортицы.', 3)
    if amount >= SCENARIO_LARGE_SHORTAGE:
        return ("СЦ-3", "Крупная недостача (>= 50 000 руб.)",
                "1. Служебное расследование.\n2. Первичная документация.\n3. Видеозаписи.", 1)
    if amount >= SCENARIO_MEDIUM_SHORTAGE:
        return ("СЦ-4", "Средняя недостача (10–50 тыс.)",
                "1. Объяснительная.\n2. Контрольный пересчёт.\n3. Сверка движений.", 2)
    if is_w and amount >= SCENARIO_WEIGHT_LOSS:
        return ("СЦ-5", "Потери весового товара",
                "1. Условия хранения.\n2. Естественная убыль.\n3. Акты списания.", 3)
    if amount < SCENARIO_WEIGHT_LOSS:
        return ("СЦ-6", "Малая недостача", "1. Точность учёта.\n2. Оценить повторяемость.", 5)
    return ("СЦ-7", "Недостача — требует проверки",
            "1. Объяснения МОЛ.\n2. Сверка движений.\n3. Решение комиссии.", 4)
