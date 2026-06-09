# create_contractor_odata.py

Створює контрагента в 1С УТ11 (українська конфігурація) через OData REST API.

## Що робить

Один скрипт обробляє будь-який вхідний документ і створює повний ланцюжок в 1С:

```
Вхідний файл → витяг реквізитів → 1С OData
                                    ├── Партнер
                                    ├── Контрагент
                                    ├── Банківський рахунок
                                    └── Угода з постачальником
```

## Підтримувані формати вхідних файлів

| Формат | Як обробляється |
|--------|----------------|
| `.jpg` `.jpeg` `.png` `.pdf` | GPT-4o vision — розпізнає реквізити з фото/скану |
| `.docx` `.doc` | GPT-4o — витягує текст і парсить реквізити |
| `.xlsx` | Перший рядок — заголовки, решта — дані (пакетний режим) |
| `.txt` `.csv` `.tsv` | CSV/TSV з заголовками (пакет) або `Поле: Значення` (один запис) |
| `.json` | Словник з реквізитами |

## Як використовувати

**Перетягни файл на `run.bat`** — більше нічого не потрібно.

Або з командного рядка:
```
C:\Python311-32\python.exe create_contractor_odata.py 1.jpg --db Vlada2
C:\Python311-32\python.exe create_contractor_odata.py batch.xlsx --db Vlada
C:\Python311-32\python.exe create_contractor_odata.py --check-only --db Vlada2
```

## Аргументи

| Аргумент | Опис |
|----------|------|
| `--db Vlada` або `--db Vlada2` | Назва бази 1С (замінює `--base-url`) |
| `--base-url URL` | Повний OData URL |
| `--user` / `--password` | Логін/пароль 1С (за замовчуванням Админ/77556670) |
| `--check-only` | Тільки перевірити підключення до 1С |
| `--sample` | Створити приклади вхідних файлів |
| `--debug` | Детальний лог |

## Налаштування

### 1С підключення
- **Server 1:** `--db Vlada` → `http://localhost/Vlada/odata/standard.odata`
- **Server 2:** `--db Vlada2` → `http://localhost/Vlada2/odata/standard.odata`

### OpenAI API ключ (для зображень і PDF)
```
D:\Project\tools\.env
OPENAI_API_KEY=sk-...
```

## Захист від дублів

Перед створенням скрипт перевіряє ЄДРПОУ в базі 1С.  
Якщо контрагент вже є — нічого не створює, повертає `status: already_exists`.

## Файли на серверах

```
D:\Project\1c_worck\
  create_contractor_odata.py   ← основний скрипт
  run.bat                      ← запуск (перетягни файл)
```

## Що зроблено (2026-06-09)

- Додано підтримку Excel (`.xlsx`) та пакетного TXT/CSV
- Додано GPT-4o vision для JPEG/PNG/PDF/DOCX
- Автодетекція кодування (UTF-8/UTF-16/cp1251)
- Аргумент `--db` для вибору бази 1С
- Деплой на Server 1 (Vlada) і Server 2 (Vlada2)
- ASCII ім'я файлу для сумісності з bat
