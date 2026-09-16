# Access Incident Report — Streamlit public access

## Симптом

Внешние пользователи по URL
https://retail-inventory-analyzer-69mv6gs2paayemqxpq2iwn.streamlit.app/
видели отказ вместо приложения:

> У вас нет доступа к этому приложению, или оно не существует.  
> Пожалуйста, войдите, чтобы продолжить.

---

## Точная подтверждённая причина

**Streamlit Community Cloud: приложение было Private**  
(`Who can view this app` ≠ public / viewer auth wall для гостей).

Доказательства до исправления:

- Анонимный браузер: redirect `/-/auth/app` → `/-/login`, UI access wall.
- Скриншот: `docs/access_denied_external_2026-09-16.png`
- `api/v2/apps/disambiguate` для гостя → 404 (Cloud скрывает private app).

**Не причина:** private GitHub (репозиторий уже был **public**).  
**Не причина:** VPN / корпоративная сеть / отсутствие e-mail invite как «единственный» путь — при Public sharing они не нужны.

---

## Этап 0. Исходное состояние (до изменений)

| Поле | Значение |
|------|----------|
| Дата/время диагностики | 2026-09-16 12:14–12:25 (+03:00) |
| GitHub URL | https://github.com/Pavel760418/retail-inventory-analyzer |
| Streamlit URL | https://retail-inventory-analyzer-69mv6gs2paayemqxpq2iwn.streamlit.app/ |
| Ветка деплоя | `main` (SHA `99d37122404dcd41e2b984e415fbc7c37c45d4ba`) |
| Видимость GitHub | **public** |
| Видимость Streamlit (до) | **Private** |
| Владелец | GitHub / Cloud: `pavel760418` |

Аудит Git: нет tracked `.env` / secrets / реальных `*.xlsx`; `.gitignore` покрывает operational data.

---

## Этап 1. Диагностика

| ID | Проверка | Результат |
|----|----------|-----------|
| A | GitHub public | ✅ |
| A2 | Private GitHub как причина | ❌ |
| B | Streamlit sharing Private | ✅ |
| C | App существует | ✅ |
| D | Внешний доступ до фикса | ❌ |
| E | Секреты в Git | ✅ чисто |

---

## Этап 2. Решение

**Вариант 1:** оставить GitHub public (уже безопасный аудит) и включить **public sharing** в Streamlit Cloud.  
Варианты 2–3 не потребовались.

---

## Этап 3. Что изменено

| Изменение | Детали |
|-----------|--------|
| Streamlit Sharing | Владелец включил public access (**Make this app public** / Sharing → public) |
| URL | **Без изменения:** https://retail-inventory-analyzer-69mv6gs2paayemqxpq2iwn.streamlit.app/ |
| Код приложения | Не менялся для доступа |
| Документация | README, `docs/deployment.md`, этот отчёт |

Дата/время исправления (проверка): **2026-09-16 ~12:30–12:35 (+03:00)**.

Почему безопасно: в Git нет реальных инвентаризаций/секретов; публичен runtime (пользователь сам загружает файлы в сессию).

---

## Этап 4. Внешний smoke-test (после фикса)

Контекст: Chromium/Puppeteer **без** Sign in Streamlit/GitHub.

| Проверка | Результат |
|----------|-----------|
| Нет access wall / «войдите» | ✅ |
| `api/v2/app/disambiguate` анонимно | ✅ 200 |
| `api/v2/app/status` → `viewerAuthEnabled` | ✅ **`false`** |
| `user.isAnonymous` | ✅ `true` |
| `/~/+/_stcore/health` | ✅ `ok` |
| Стартовая страница / загрузка 2× Excel | ✅ |
| Запуск анализа на DEMO xlsx | ✅ «Анализ готов», вкладки Главная / … / Excel-отчёт |
| KPI на главной | ✅ (позиции, недостача/излишки, сальдо) |
| Утечки путей/traceback/токенов | ✅ нет (`LEAK_CHECK PASS`) |
| Скриншоты | `docs/public_access_ok_2026-09-16.png`, `docs/public_smoke_functional_2026-09-16.png` |

Старый URL **не менялся** — в рассылках заменять не нужно.

---

## Меры предотвращения

1. После каждого Deploy: Share → **Make this app public** (если не унаследовалось).
2. Чеклист в `docs/deployment.md`: проверка в **инкогнито** без логина.
3. Не считать успехом открытие только под владельцем.
4. Не коммитить реальные `*.xlsx` / `.env` / secrets.

## Сообщение для пользователей

Приложение доступно по ссылке (логин не нужен):

https://retail-inventory-analyzer-69mv6gs2paayemqxpq2iwn.streamlit.app/
