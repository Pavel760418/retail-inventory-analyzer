# Деплой на Streamlit Community Cloud

## Предварительные условия

1. Репозиторий GitHub опубликован (рекомендуется **private**).
2. В корне есть `app.py` и `requirements.txt`.
3. Нет зависимости от локальных путей `C:\Users\...`.

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

## Деплой (одно действие пользователя)

1. Откройте https://share.streamlit.io и войдите через GitHub.
2. **New app**.
3. Выберите репозиторий `retail-inventory-analyzer` и ветку `main` (или `feature/streamlit-inventory-app`).
4. **Main file path:** `app.py`
5. Advanced → Python 3.10 или 3.11.
6. Нажмите **Deploy**.

Секреты не требуются. Пользовательские Excel не сохраняются между сессиями.

## После деплоя — smoke-test

- [ ] Стартовый экран открывается  
- [ ] Загрузка двух demo/синтетических xlsx (или ваших файлов)  
- [ ] KPI на вкладке «Главная»  
- [ ] Скачивание Excel  
- [ ] Понятные ошибки при неверном файле  

## Ограничения Cloud

- Большие сети: matching может быть долгим (есть guardrail в коде).
- Запись эталона иерархии из upload на Cloud не рекомендуется — используйте встроенный `data/master_hierarchy.json`.
