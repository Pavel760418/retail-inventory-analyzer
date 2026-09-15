# Архитектура

## Обзор

```text
[Excel TDSheet]──► parser ──► enrich ──► overlap ──► metrics/scenarios
                      │                      │
                      │                      ├── anomalies (R4)
                      │                      ├── price_compare (R4)
                      │                      └── capitalization (R4)
                      └──► excel/export ──► XLSX отчёт

CLI (main.py) / Streamlit (app.py) ───┘
```

## Слои

1. **Entrypoints**
   - `main.py` — интерактивный CLI (tkinter → console fallback)
   - `app.py` — Streamlit UI (upload ×2 → preview → download)

2. **Config / Master data**
   - `config/settings.py` — пороги, флаги, env
   - `data/master_hierarchy.json` — эталон иерархии (source of truth)
   - `src/master_hierarchy.py` — load + lookup

3. **Domain**
   - `src/models.py` — `Config`, `ParseMeta`
   - `src/parser.py` — валидация и разбор иерархии выгрузки (+ флаги пустых книжных сумм)
   - `src/catalog.py` — enrich через master hierarchy (не keyword)
   - `src/categories.py` — legacy heuristics (deprecated, не используется в runtime)
   - `src/anomalies.py` — Release 4: аномалии книжной нормативной/фактической суммы
   - `src/price_compare.py` — Release 4: цены между магазинами (exact name)
   - `src/capitalization.py` — Release 4: оприходование излишков + exact match по складам
   - `src/overlap.py` — matching излишек↔недостача по `category_group` (= leaf_group эталона)
   - `src/metrics.py` / `src/scenarios.py`

4. **Presentation / Export**
   - `src/excel/styles.py` — стили openpyxl
   - `src/excel/export.py` — листы отчёта + `build_report`

5. **Quality**
   - `tests/test_smoke.py`, `tests/test_capitalization.py`, `tests/test_anomalies_prices.py`
   - `.github/workflows/ci.yml`

## Поток данных `build_report`

1. `parse_network_excel`
2. `filter_by_period`
3. `load_catalog` + `enrich_dataframe`
4. `build_operational_overlaps` (+ optional cross-store)
5. Метрики магазина/документа/SKU
6. `detect_book_sum_anomalies` + `compare_prices_across_stores`
7. `parse_capitalization_excel` + `match_capitalization_to_inventory` (если файл передан)
8. Сборка workbook (Сводка, Выводы, Рейтинг, Аномалии, Проверка цен,
   Оприходование, Топы, Перекрытие, Мероприятия, детали магазинов, Метаданные, Содержание)
9. `wb.save(output_path)`

## Границы ответственности

| Модуль | Не должен |
|--------|-----------|
| `parser` | Считать пересорт / писать Excel |
| `overlap` | Знать стили openpyxl |
| `anomalies` / `price_compare` | Менять surplus/shortage формулы |
| `excel/*` | Менять правила matching |
| `app.py` | Дублировать формулы метрик (только оркестрация + UI) |

## Расширение

- Новый сценарий → `src/scenarios.py` + цвет в `SCENARIO_COLORS`
- Новая категория → эталон иерархии (не keyword)
- Новый лист отчёта → функция `sh_*` в `src/excel/export.py` + вызов в `build_report`
