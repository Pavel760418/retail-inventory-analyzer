# Эталонная иерархия номенклатуры

## Источник истины

Иерархия зафиксирована из шаблона `Шаблон_иерархия.xlsx` (лист «Лист1», outline Excel):

| Колонка | Смысл |
|---------|--------|
| B | Узлы иерархии (группы / категории) |
| C | Листовые SKU (номенклатура) |
| outlineLevel | Глубина узла (0…6) |

## Где хранится в проекте

| Файл | Назначение |
|------|------------|
| `data/master_hierarchy.json` | Runtime-эталон (встроен в репозиторий) |
| `data/master_hierarchy.csv` | Человекочитаемая копия для ревью |
| `data/fixtures/master_hierarchy_sample.json` | Обезличенный sample для тестов |
| `src/master_hierarchy.py` | Загрузка + lookup API |
| `scripts/build_master_hierarchy.py` | Пересборка эталона из xlsx (только для maintainers) |

## Как подставляется при анализе

1. При старте / первом enrich загружается `data/master_hierarchy.json`.
2. Для каждой строки входного Excel ищется SKU по нормализованному имени (`sku_key`).
3. Если найдено:
   - `category_group` / `category_label` = ближайший родитель (`leaf_group`);
   - `hierarchy_path` = полный путь `level_0 / level_1 / …`;
   - `category_method` = `master_hierarchy`.
4. Если **не** найдено:
   - категория **не выдумывается** keyword-эвристикой;
   - `category_label` = «Не найдено в справочнике»;
   - строка попадает на лист / список «Вне справочника»;
   - в matching пересорта уникальный служебный ключ не схлопывает разные unknown SKU в одну «ложную» категорию.

## Пересборка эталона

Когда обновится шаблон иерархии:

```bash
python scripts/build_master_hierarchy.py --source "C:\path\to\Шаблон_иерархия.xlsx"
```

Закоммитьте обновлённые `data/master_hierarchy.json` и `.csv`.  
Runtime и Streamlit Cloud **не** требуют локального пути к шаблону.

## Legacy

`src/categories.py` (`classify_category`) оставлен как deprecated и **не** вызывается из `enrich_dataframe`.
