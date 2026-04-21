# Baseline smoke-checklist для текущего `acad-cmd`

Цель: зафиксировать текущее поведение `acad-cmd` до внедрения event bridge (roadmap, пункт 1).

## Preconditions

1. Запущен AutoCAD (рекомендуемо открыть любой DWG).
2. Локальная среда проекта готова:
   - `.venv` создан;
   - зависимости установлены (`pip install .`).
3. Сервер запускается из этого репозитория.

## Быстрый запуск baseline

```bat
.venv\Scripts\python.exe scripts\mcp_baseline_capture.py
```

Скрипт пишет отчёт в `out/baseline/acad_cmd_baseline_YYYYMMDD_HHMMSS.json`.

## Что проверяется

1. `get_status` (до сценария): есть подключение к AutoCAD и валидный status snapshot.
2. `start_logging(mode=logfile)`: поднят logfile stream (`stream_id`, `logfile_path`, `cursor`).
3. `send_command(wait=true)`: команда отправляется и завершается без `needs_input`.
4. `run_lisp(wait=true)`: LISP-выражение выполняется и возвращает marker/result без зависания.
5. `LOGFILEMODE`: через `(getvar "LOGFILEMODE")` и рост лог-файла после команд.
6. `get_new_output_since` и `get_last_output(source=logfile)`: доступен хвост текстового лога.
7. `get_status` (после сценария): соединение остаётся рабочим.

## Критерий "baseline зафиксирован"

1. Отчёт создан со `status = "passed"`.
2. В отчёте присутствуют блоки:
   - `checks.get_status_before`
   - `checks.send_command`
   - `checks.run_lisp`
   - `checks.logfilemode_probe`
   - `checks.get_status_after`
3. Отчёт сохранён и используется как эталон для сравнения после следующих пунктов roadmap.
