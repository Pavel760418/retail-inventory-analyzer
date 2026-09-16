# Деплой на Streamlit Community Cloud

## Предварительные условия

1. Репозиторий GitHub: для **публичного** приложения без логина удобнее **public** (текущий статус
   `retail-inventory-analyzer` — public). Private repo возможен, но тогда в Cloud нужно явно
   выставить public sharing (и учитывать лимит: один private app на workspace).
2. В корне есть `app.py` и `requirements.txt`.
3. Нет зависимости от локальных путей `C:\Users\...`.
4. В Git нет `.env`, `.streamlit/secrets.toml`, реальных `*.xlsx` / выгрузок.

## Локальный запуск (проверка перед деплоем)

```bash
cd retail_inventory_analyzer
python -m venv .venv
# Windows:
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Тесты:

```bash
pytest -q
```

## Деплой

1. Откройте https://share.streamlit.io и войдите через GitHub.
2. **New app**.
3. Репозиторий `retail-inventory-analyzer`, ветка `main`.
4. **Main file path:** `app.py`
5. Advanced → Python 3.10 или 3.11.
6. **Deploy**.

Секреты не требуются. Пользовательские Excel не сохраняются между сессиями.

## Обязательно: публичный доступ (Sharing)

После деплоя приложение может остаться **private** (особенно если когда‑то деплоилось
из private repo или вручную выставили «Only specific people»). Тогда внешние пользователи
увидят: *«У вас нет доступа… войдите»*.

Сделайте так:

1. Откройте app под аккаунтом владельца.
2. **Share** (справа сверху) → **Make this app public**  
   **или** Settings → **Sharing** → **Who can view this app** →  
   **This app is public and searchable** → Save.
3. URL обычно **не меняется**.

Пользователям **не** нужны аккаунты Streamlit/GitHub, VPN или приглашения по e-mail.

## Проверка публичного доступа после деплоя

Выполнять в **режиме инкогнито / гостевом профиле** (без Sign in в Streamlit и GitHub):

- [ ] URL открывается без сообщения «нет доступа» / «войдите» / «does not exist»
- [ ] Видна стартовая страница («Анализ итогов инвентаризации…»)
- [ ] Доступна форма загрузки файла
- [ ] На синтетическом файле считаются KPI на «Главная»
- [ ] Скачивается Excel-отчёт
- [ ] Нет traceback, локальных путей `C:\Users\...`, токенов на экране
- [ ] В репозитории по-прежнему нет реальных данных и секретов

Не считать успехом открытие приложения **только** под владельцем.

## После деплоя — функциональный smoke-test

- [ ] Стартовый экран открывается  
- [ ] Загрузка demo/синтетических xlsx  
- [ ] KPI на вкладке «Главная»  
- [ ] Скачивание Excel  
- [ ] Понятные ошибки при неверном файле  

## Ограничения Cloud

- Большие сети: matching может быть долгим (есть guardrail в коде).
- Запись эталона иерархии из upload на Cloud не рекомендуется — используйте встроенный `data/master_hierarchy.json`.

## Инцидент доступа

См. [ACCESS_INCIDENT_REPORT.md](ACCESS_INCIDENT_REPORT.md).
