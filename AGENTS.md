# AGENTS.md — «Выгрузка оценок из МЭШ»

Единственное описание проекта. Кратко.

## Что это

Windows-утилита: собирает все оценки за текущий учебный год из дневника
**school.mos.ru (МЭШ)** и сохраняет в `grades.xlsx`. Репозиторий
`maxinteresa-ops/dnevnik-mesh-export`, лицензия **CC BY-NC-SA 4.0**. Аудитория — родители;
сценарий: запустить `.exe`, ничего не настраивать.

## Файлы

| Файл | Назначение |
|---|---|
| `dnevnik-mesh-export.py` | Весь код (~840 строк, один файл, логика в `main()`; сбор и запись листа — `collect_year()` / `write_year_sheet()`) |
| `dnevnik-mesh-export.exe` / `.zip` | Сборка PyInstaller (onefile, console) / дистрибутив для релиза |
| `grades.xlsx` | Результат работы (gitignored) |
| `chrome-profile/` | Профиль Chrome с авторизацией МЭШ (gitignored) |
| `.github/ISSUE_TEMPLATE/` | `bug_report.yml`, `question.yml`, `config.yml` (blank issues выключены) |
| `ридми/` | Все картинки README + копия его текста (`ридми/README.md`). Картинки лежат в репозитории, а не на CDN GitHub |
| `social-preview.png` | Превью репозитория на GitHub (в README не используется) |

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
   Запрос повторяется до `REQUEST_ATTEMPTS` (3) раз с паузой `REQUEST_RETRY_DELAY` (3 с):
   МЭШ иногда отвечает HTTP 500 или не отвечает вовсе, без повторов неделя терялась молча.
5. **Сбор**: `YEARS_TO_COLLECT` учебных годов, по умолчанию **1** (только текущий — почему,
   см. «Что известно про API МЭШ»). Текущий уч. год определяется автоматом по дате:
   `_year = now.year`, если месяц ≥ 9, иначе `now.year - 1`; диапазоны годов строятся как
   `range(_year, _year - YEARS_TO_COLLECT, -1)`. На каждый год — свой `collect_year()`:
   каждый ребёнок × 38 недель от 1 сентября этого года.
   `get_marks` возвращает `None` при ошибке запроса (неделя считается «не загружена»),
   `[]` — если оценок за неделю нет.
6. **Коэффициент (вес) оценки**: МЭШ показывает цифру рядом с оценкой только когда она больше 1.
   Поле в ответе API — **`weight`** (подтверждено на живом ответе 04.10.2026); на всякий случай
   перебираются также `coefficient`, `mark_weight`, `factor`, `ratio` (`parse_coefficient`).
   Нет поля / мусор / <1 → коэффициент 1, максимум — `MAX_COEFFICIENT` (10).
   Первая полученная оценка целиком пишется в лог (`dump_mark_fields`) — если МЭШ переименует
   поле, это будет видно по дампу. Оценка с коэффициентом N даёт N одинаковых строк в таблице.
7. **XLSX** (`openpyxl`): лист на каждый собранный учебный год, имя = `sheet_name_for_year(year)`
   (`Оценки_2026-2027`); при `YEARS_TO_COLLECT = 1` лист один и он активный.
   На листе — плоская таблица (Ребёнок, Дата, Предмет, Тема урока, Оценка, Коэффициент)
   с автофильтром `A1:{последняя колонка}{n}`; правее (`len(headers)+3`) — пивот
   «предмет × оценка» по каждому ребёнку (коэффициент учтён автоматически, т.к. строки
   продублированы). Год без оценок даёт пустой лист с одной строкой заголовков.
8. **Финал**: `taskkill` Chrome, статистика в консоль — **отдельным блоком по каждому году**
   (оценок, с коэффициентом, недель с данными / без оценок / не загружено, среднее за неделю),
   при одном годе — примечание, почему нет прошлого года, `input()` для Enter.

## Соглашения

- Зависимости: `openpyxl`, `websocket-client` (+ `pyinstaller` для сборки); проверяются при импорте.
- Только Windows (`taskkill`, `where`, `ctypes.windll`).
- Ошибки не пробрасываются: `log_exception(context)` пишет traceback в `dnevnik-mesh-export_<ts>.log`.
- `SCRIPT_DIR` учитывает `sys.frozen`; вывод и комментарии — на русском, для непрограммиста.

## Сборка

Зависимости для сборки: `pip install openpyxl websocket-client pyinstaller`.

Собирать **обязательно со списком `--exclude-module`**: без него PyInstaller затягивает
в exe случайные пакеты из окружения (numpy с OpenBLAS, Pillow, lxml, psutil, pyreadline3,
pywin32, PyYAML, charset_normalizer) — сборка раздувается с ~9,4 МБ до ~31 МБ.
Проверено 05.10.2026: в архиве сборки без исключений лежало ~17 МБ неиспользуемых пакетов.

```powershell
$ex = 'numpy','PIL','lxml','psutil','yaml','pyreadline3','win32','win32com','pythoncom','pywintypes',
      'charset_normalizer','requests','tkinter','unittest','pydoc','doctest','sqlite3','xmlrpc',
      'http.server','socketserver','cgi','multiprocessing','asyncio','concurrent','distutils',
      'setuptools','pkg_resources','pip','IPython','pytest','pandas','matplotlib','scipy','mypy'
$a = '--onefile','--console','--noconfirm','--distpath','.','--name','dnevnik-mesh-export'
foreach ($m in $ex) { $a += @('--exclude-module', $m) }
$a += 'dnevnik-mesh-export.py'
python -m PyInstaller @a
```

Нельзя исключать `email` (его использует `http.client` при HTTPS-запросах) и `ssl`/`_socket`/
`_hashlib` (нужны для HTTPS и CDP). После сборки проверяйте содержимое:
`python -m PyInstaller.utils.cliutils.archive_viewer -r -l dnevnik-mesh-export.exe`
(не должно быть numpy/PIL/lxml, должны быть openpyxl и websocket).

Дистрибутив: `Compress-Archive -Path dnevnik-mesh-export.exe -DestinationPath dnevnik-mesh-export.zip -Force`.

## Что известно про API МЭШ

Проверено на живом аккаунте 04.10.2026 (один ребёнок, 8 класс):

- `GET /api/family/web/v1/marks?student_id=&from=&to=` — **только текущий учебный год**.
  Дата-фильтр работает корректно (за текущий год оценки есть), но за любой прошлый год
  всегда `payload: []` — проверялись 2020-2021, 2024-2025, 2025-2026 и календарные диапазоны.
- `GET /api/ej/core/family/v1/marks?student_profile_id=&pid=` — второй слой API, родителю доступен,
  но **игнорирует и `from`/`to`, и `academic_year_id`**: всегда одни и те же записи текущего года.
- `GET /api/ej/core/family/v1/academic_years` — список годов (2014-2015 … 2033-2034) с полем
  `current_year`; сами годы в системе есть, доступа к оценкам за них это не даёт.
- `academic_year_id` эндпоинты оценок принимают, но не применяют (проверялось с id=13 = 2025-2026:
  вернулись данные текущего года).
- Веб-интерфейс: маршрут `/marks/archive` объявлен в JS дневника, но живая страница
  `school.mos.ru/diary/marks/archive` отдаёт заглушку «Эту страницу ещё не изобрели».
- `GET /api/family/web/v1/final_marks?student_id=` — **итоговые (годовые) оценки за все годы**
  одним ответом (по 9–14 предметов на 2020-2021 … 2025-2026, за текущий год пусто).
  Утилита это не использует — по решению владельца проекта такие листы не нужны.

**Итог:** детальные оценки за прошлые учебные годы МЭШ родителю не отдаёт — ни через сайт,
ни через API. Это ограничение сервиса, а не утилиты.

## Известные слабые места

- Пивот сортирует оценки как числа, нечисловые — после цифр (`x.isdigit()`).
- Запись колонок ограничена 90-й (`chr(64 + idx)`).
- Сбор строго последовательный, хардкод `WEEK_COUNT = 38` (без повторов, если задать
  `REQUEST_ATTEMPTS = 1`). Неделя, которую не удалось загрузить за все попытки, попадает
  в статистику «Недель не загружено», но оценки за неё теряются.
- Автотестов нет: проверка — ручной запуск с готовым `chrome-profile/` и просмотр `grades.xlsx`.

## Правила работы

- **Коммитить после каждого изменения.** Любая правка файлов проекта завершается
  отдельным осмысленным git-коммитом — не оставлять изменения незакоммиченными.
- Описания проекта живут в одном месте — в этом файле (`README.md` — только для пользователей GitHub).

## Заметки по рабочей копии

- URL `origin` — обычный `https://github.com/maxinteresa-ops/dnevnik-mesh-export.git`,
  токен из него убран (05.10.2026). Пуш идёт через `gh` / Git Credential Manager.
- `README.md` в корне — то, что видно на GitHub; картинки берутся из `ридми/` по относительным
  путям (раньше часть лежала на CDN GitHub, теперь всё в репозитории).
- `ридми/README.md` — копия для чтения локально, пути к картинкам в ней без префикса `ридми/`.
  **При правке корневого README обновляйте и копию**, иначе они разойдутся.
- В `.gitignore` внесены копии результатов (`*grades*.xlsx`), черновики скриншотов
  (`first-run-dialog[0-9]*.png`), собранный дистрибутив (`dnevnik-mesh-export.zip`)
  и Excel-локи (`~$*.xlsx`). Рабочая копия при этом чистая — лишнего в `git status` нет.
- Временные материалы (папка `1/` с mhtml, логи проверок, черновики скриншотов) удалены
  05.10.2026; нужное из них описано в этом файле. Копия старого результата
  `2025_grades — копия.xlsx` оставлена намеренно.
