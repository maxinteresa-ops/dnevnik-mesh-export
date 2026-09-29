# AGENTS.md — «Выгрузка оценок из МЭШ»

Единственное описание проекта. Кратко.

## Что это

Windows-утилита: собирает все оценки за учебный год из дневника **school.mos.ru (МЭШ)**
и сохраняет в `grades.xlsx`. Репозиторий `maxinteresa-ops/dnevnik-mesh-export`, лицензия
**CC BY-NC-SA 4.0**. Аудитория — родители; сценарий: запустить `.exe`, ничего не настраивать.

## Файлы

| Файл | Назначение |
|---|---|
| `dnevnik-mesh-export.py` | Весь код (~690 строк, один файл, логика в `main()`) |
| `dnevnik-mesh-export.exe` / `.zip` | Сборка PyInstaller (onefile, console) / дистрибутив для релиза |
| `grades.xlsx` | Результат работы (gitignored) |
| `chrome-profile/` | Профиль Chrome с авторизацией МЭШ (gitignored) |
| `.github/ISSUE_TEMPLATE/` | `bug_report.yml`, `question.yml`, `config.yml` (blank issues выключены) |
| `reasonix.toml` | Локальный конфиг агента |

## Как работает

1. **Chrome**: `taskkill /f /im chrome.exe`, затем запуск с `--remote-debugging-port=9222`,
   `--remote-allow-origins=*`, `--user-data-dir=chrome-profile`. Поиск браузера — `where chrome.exe`
   + стандартные пути Program Files / LocalAppData.
2. **Авторизация**: до 150 итераций по 2 с (~5 мин), ждёт cookies `aupd_token` и `active_student`
   на странице `school.mos.ru/.../marks`; при 5 подряд потерях порта 9222 перезапускает Chrome;
   по успеху возвращает фокус консоли через Win32 API.
3. **CDP** (`websocket-client`): `Network.getAllCookies`; строка cookie собирается из
   `session-cookie`, `JSESSIONID`, `student_person_id`, `active_student`, `cluster_id`, `aupd_token` и др.
4. **API** через `urllib` с заголовком `X-Mes-Subsystem: familyweb`:
   `/api/family/web/v1/profile` → дети, `/api/family/web/v1/marks?student_id=&from=&to=` → оценки.
5. **Сбор**: каждый ребёнок × 38 недель от 1 сентября учебного года (месяц < 9 → предыдущий год).
6. **XLSX** (`openpyxl`): лист «Оценки» — плоская таблица (Ребёнок, Дата, Предмет, Тема урока, Оценка)
   с автофильтром `A1:E{n}`; правее (`len(headers)+3`) — пивот «предмет × оценка» по каждому ребёнку.
7. **Финал**: `taskkill` Chrome, статистика в консоль, `input()` для Enter.

## Соглашения

- Зависимости: `openpyxl`, `websocket-client` (+ `pyinstaller` для сборки); проверяются при импорте.
- Только Windows (`taskkill`, `where`, `ctypes.windll`).
- Ошибки не пробрасываются: `log_exception(context)` пишет traceback в `dnevnik-mesh-export_<ts>.log`.
- `SCRIPT_DIR` учитывает `sys.frozen`; вывод и комментарии — на русском, для непрограммиста.

## Сборка

```bash
pip install openpyxl websocket-client pyinstaller
pyinstaller --onefile --console --distpath . --name dnevnik-mesh-export dnevnik-mesh-export.py
```

## Известные слабые места

- Пивот сортирует оценки как числа, нечисловые — после цифр (`x.isdigit()`).
- Запись колонок ограничена 90-й (`chr(64 + idx)`).
- Сбор строго последовательный, без ретраев на уровне запроса; хардкод `WEEK_COUNT = 38`.
- Автотестов нет: проверка — ручной запуск с готовым `chrome-profile/` и просмотр `grades.xlsx`.

## Правила работы

- **Коммитить после каждого изменения.** Любая правка файлов проекта завершается
  отдельным осмысленным git-коммитом — не оставлять изменения незакоммиченными.
- Описания проекта живут в одном месте — в этом файле (`README.md` — только для пользователей GitHub).

## Заметки по рабочей копии

- URL `origin` содержит GitHub-токен в открытом виде — заменить на SSH или перевыпустить токен.
