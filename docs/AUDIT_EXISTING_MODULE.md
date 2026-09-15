# Аудит существующего модуля — retail_inventory_analyzer

**Дата аудита:** 2026-09-15  
**Путь:** `C:\Users\Администратор\Documents\ЗЯ\ПЮ\Инвет\retail_inventory_analyzer`  
**Версия пакета:** `2.0.0` (`src/__init__.py`)

## 1. Фактическая структура

Форк `network_inventory_analyzer` под розницу **без собственного производства**:

| Компонент | Назначение |
|-----------|------------|
| `app.py` | Streamlit UI (базовый, требует доработки UX) |
| `main.py` | CLI: 3 файла (инвент / оприходование / иерархия) |
| `config/settings.py` | Период, пороги, ревизор=08:00:00 |
| `src/parser.py` | Парсинг TDSheet + валидация заголовков |
| `src/catalog.py` + `master_hierarchy.py` | Эталон иерархии из `data/master_hierarchy.json` |
| `src/overlap.py` | Пересорт / перекрытие |
| `src/metrics.py` | KPI магазинов, сети, выводы |
| `src/author.py` | Ревизор / оператор по времени документа |
| `src/anomalies.py` | Аномалии книжных сумм |
| `src/capitalization.py` | Оприходование излишков (закрытие смены) |
| `src/excel/export.py` | Итоговый многолистовый Excel |
| `src/pipeline.py` | Фасад `run_analysis` → `build_report` |
| `tests/` | smoke, capitalization, anomalies, author, hierarchy |

**Проверка цен между магазинами — отключена** (отличие от сетевого модуля).

## 2. Язык и зависимости

- Python **3.10+** (локально проверено 3.10.7)
- pandas, numpy, openpyxl, xlrd==1.2.0, streamlit, python-dotenv, pytest
- Для web-визуализаций добавляется **plotly** (UI-слой, не бизнес-логика)

## 3. Точки входа

| Режим | Команда |
|-------|---------|
| Web | `streamlit run app.py` |
| CLI | `python main.py` |
| Библиотека | `src.pipeline.run_analysis(Config(...))` |

## 4. Входные данные

1. **Инвентаризация** `.xlsx` (лист TDSheet-like):
   - обязательные маркеры: `Документ.Подразделение`, `Количество книжное`, `Количество фактическое`, `Сумма излишек`, `Сумма недостача`;
   - иерархия строк через outlineLevel: магазин → документ → номенклатура;
   - время документа в заголовке: ровно `08:00:00` = **ревизор**, иначе **оператор**.
2. **Оприходование излишков** `.xlsx`: колонки `Закрытие смены количество`, `Закрытие смены сумма`.
3. **Иерархия** (опционально при уже собранном эталоне): `Шаблон_иерархия_безпроизводства.xlsx` → `data/master_hierarchy.json`.

Готовые файлы `Анализ_*.xlsx` отклоняются валидатором.

## 5. Выходные данные (Excel-листы)

Сводка, Выводы, Рейтинг магазинов, Аномалии, Ревизоры vs Операторы, Цепочки пересортов, Оприходование излишков, Топ недостач/пересортов, Хронические, Позиции×Магазины, Перекрытие, Мероприятия, Недостачи, Излишки, Чистые недостачи, Анализ позиций, Вне справочника, Документы, детали магазинов, Метаданные, Содержание.

## 6. Ключевые KPI (фактические из `calc_network_summary` / author / anomalies)

- `stores`, `docs`, `sku_lines`, `sku_disc`
- `surplus`, `shortage`, `net`, `overlap`, `clean_shortage`
- `shrinkage_pct`, `recovery_pct`
- аномалии type1/2/3, critical
- вклад ревизор/оператор, цепочки пересортов
- оприходование: matched / matched_shortage / unmatched

## 7. Поток данных

```text
Upload xlsx → tempfile
  → validate_network_structure / parse_network_excel
  → enrich_author_columns
  → filter_by_period
  → enrich_dataframe (master hierarchy)
  → build_operational_overlaps (+ optional cross-store)
  → calc_store_metrics / calc_sku_cross / calc_network_summary
  → detect_book_sum_anomalies
  → author stats + chains
  → parse_capitalization + match
  → generate_conclusions
  → build_report / export_report_bytes → download
  → cleanup tempfile
```

## 8. Найденные риски

| Риск | Статус |
|------|--------|
| Абсолютный путь `C:\Users\Администратор\Documents\ЗЯ\ПЮ\Инвет` в `app.py` / `main.py` | Устраняется |
| `.env` с локальным путём | Не коммитить; `.env.example` обезличить |
| Скрипты audit_* с абсолютными путями | Только локальные утилиты; не для Cloud |
| `*.docx` release summary в корне | Исключить из Git |
| Квадратичный matching на больших файлах | Guardrail есть; риск timeout на Cloud |
| Запись `master_hierarchy.json` при upload иерархии | Допустимо локально; на Cloud — только session temp |
| Реальные `*.xlsx` рядом с проектом | `.gitignore` уже исключает |

## 9. Что переиспользуется без изменения алгоритмов

- Весь расчётный контур: parser, overlap, metrics, scenarios, author, anomalies, capitalization, excel/export
- Формат и состав Excel-отчёта
- Эталон `data/master_hierarchy.json`
- Существующие pytest (с правкой локальных путей в optional-тестах)

## 10. Минимальные доработки для web-версии

1. UX Streamlit: вкладки, KPI-карточки, выводы, статусы, ≤3 графика Plotly.
2. Адаптер: валидация с понятными сообщениями, quality-отчёт, helpers UI.
3. Убрать жёсткие абсолютные пути из runtime (`app.py`, `main.py`, `.env.example`).
4. Excel в tempfile / bytes + `st.download_button`.
5. Синтетические fixtures + тесты валидации/KPI/Excel.
6. Документация: INPUT_FILE_GUIDE, DEPLOYMENT, README.
7. Усилить `.gitignore`; подготовить private GitHub repo.
8. **Не** переносить код в `src/retail_inventory_analyzer/` — ломает импорты; сохраняем рабочий layout `src/*`.

## 11. Воспроизводимость Excel

Локально: `python main.py` или `run_analysis(Config(...))` формирует тот же `build_report`.  
Web-слой вызывает тот же `build_report` / `export_report_bytes` — расчёты не дублируются.
