# Анализ итогов инвентаризации торгового зала

Streamlit-приложение для бизнес-пользователей поверх существующего модуля
`retail_inventory_analyzer` (розница без собственного производства).

## Возможности

- Загрузка Excel инвентаризации и оприходования излишков
- Контроль структуры файла до расчёта
- KPI, выводы для руководителя, статусы/риски, до 3 графиков
- Вкладки: Главная / Детализация / Контроль качества / Excel-отчёт / Инструкция
- Скачивание полного Excel-отчёта (тот же `build_report`, что и CLI)
- Разрез ревизор (08:00:00) / оператор; аномалии книжных сумм; без проверки цен

## Архитектура

```text
app.py (UI)
  → validators / analysis_service / ui_helpers
  → src/parser, overlap, metrics, author, anomalies, capitalization
  → src/excel/export.build_report
```

Бизнес-логика анализа **не переписана** — UI вызывает существующий пайплайн.

## Установка и запуск

```bash
cd retail_inventory_analyzer
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
streamlit run app.py
```

CLI:

```bash
python main.py
```

## Входной файл

См. [docs/INPUT_FILE_GUIDE.md](docs/INPUT_FILE_GUIDE.md).

## Тесты

```bash
pytest -q
```

## Политика данных

- Реальные `*.xlsx` / выгрузки / `.env` / токены **не коммитятся**
- Загрузки пользователя — только tempfile, удаляются после анализа
- В `sample_data/` — только DEMO/SYNTHETIC

## Деплой

См. [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md).

Кратко: Streamlit Cloud → New app → GitHub repo → `app.py` → Deploy.

## Документация

| Файл | Содержание |
|------|------------|
| [docs/AUDIT_EXISTING_MODULE.md](docs/AUDIT_EXISTING_MODULE.md) | Аудит модуля |
| [docs/INPUT_FILE_GUIDE.md](docs/INPUT_FILE_GUIDE.md) | Входной формат |
| [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md) | Деплой |
| [docs/user_guide.md](docs/user_guide.md) | Пользовательское руководство |
