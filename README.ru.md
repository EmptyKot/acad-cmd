# acad-cmd (MCP-сервер для AutoCAD)

Локальный MCP-сервер (stdio / JSON-RPC), который подключается к AutoCAD в Windows через COM (pywin32) и предоставляет MCP-инструменты для работы с командной строкой AutoCAD.

Что умеет:

- отправлять команды в командную строку AutoCAD (`SendCommand`);
- открывать DWG и гарантированно переключать активный контекст на открытый чертеж;
- использовать полное логирование истории команд через `LOGFILEMODE` / `LOGFILENAME` (основной источник текстового вывода);
- при необходимости читать `LASTPROMPT` для обратной совместимости;
- опционально использовать in-process `.NET` event bridge (`AcadEventBridge`) для структурированных событий;
- выполнять ожидание завершения команд по схеме event-first с автоматическим fallback на COM idle wait;
- писать audit log (JSONL) по каждому вызову инструмента.

## Требования

- Windows
- AutoCAD (в первую очередь тестировался AutoCAD 2021; другие версии тоже могут работать)
- Python 3.10+

## Совместимость (текущий baseline)

| Компонент | Статус |
| --- | --- |
| Windows 10/11 | Поддерживается |
| AutoCAD 2021 (major 24) | Основная тестовая цель |
| Python 3.10+ | Поддерживается |

## Установка

```bat
py -3.11 -m venv .venv
.venv\Scripts\activate
pip install .
```

Если команда `python` открывает Microsoft Store, используйте `py -3.11`, как в примере выше.

## Запуск (standalone)

Рекомендуется сначала запустить AutoCAD, затем:

```bat
acad-cmd
```

Или использовать вспомогательный скрипт (создает `.venv` и устанавливает пакет при первом запуске):

```bat
start_server.bat
```

Если нужно, чтобы сервер сам запускал AutoCAD, задайте `AUTOCAD_MCP_ACAD_EXE` с полным путем к `acad.exe`.

Если установлено/запущено несколько продуктов Autodesk (например Civil 3D и AutoCAD) и нужно зафиксировать конкретный major AutoCAD, задайте `AUTOCAD_MCP_TARGET_MAJOR`.

Пример: для AutoCAD 2021 major = `24`:

```bat
set AUTOCAD_MCP_TARGET_MAJOR=24
```

Особенности подключения:

- экземпляры AutoCAD обнаруживаются через COM (`GetActiveObject`, затем при необходимости `Dispatch`);
- в некоторых установках обычный UI-экземпляр не виден через `GetActiveObject`;
- в таком случае `Dispatch` может подключиться к существующему экземпляру или создать новый automation-экземпляр;
- чтобы запретить создание нового экземпляра через COM-активацию, установите `AUTOCAD_MCP_ALLOW_NEW_INSTANCE=0`;
- если нужно, чтобы обычный UI-экземпляр корректно обнаруживался, запускайте AutoCAD с automation-параметром (часто `acad.exe /automation`, но это зависит от установки).

Runtime-логи пишутся в `logs/acad-cmd/<session_id>/`.

## Event Bridge (опционально)

`acad-cmd` может использовать in-process плагин AutoCAD (`AcadEventBridge`) и named pipe для структурированных событий.

- завершение команд/LISP может определяться по event stream (обычно быстрее и стабильнее, чем по текстовому выводу);
- lifecycle событий документов доступен в статусе/диагностике;
- если bridge недоступен или деградировал, сервер автоматически откатывается на legacy-механику (COM idle wait и логфайл).

Исходники плагина и smoke-команды: `plugins/AcadEventBridge/README.md`.

## Конфигурация (переменные окружения)

Подключение / выбор версии:

- `AUTOCAD_MCP_TARGET_MAJOR` (опционально): фиксирует major AutoCAD (например, `24` для AutoCAD 2021);
- `AUTOCAD_MCP_ALLOW_NEW_INSTANCE` (по умолчанию: разрешено): установите `0`, чтобы запретить запуск нового `acad.exe` через COM;
- `AUTOCAD_MCP_USE_DISPATCH` (по умолчанию: выключено, кроме случая с `AUTOCAD_MCP_TARGET_MAJOR`): принудительно пробует `Dispatch`;
- `AUTOCAD_MCP_PREFER_CURVER` (по умолчанию: выключено): предпочитать registry `CurVer` ProgID при определении версии AutoCAD;
- `AUTOCAD_MCP_EVENT_BRIDGE_ENABLED` (по умолчанию: `0`): включает интеграцию с event bridge.

Поведение event bridge:

- `AUTOCAD_MCP_EVENT_BRIDGE_HEARTBEAT_TIMEOUT_SEC` (по умолчанию: `6.0`): порог актуальности heartbeat при ожидании;
- `AUTOCAD_MCP_EVENT_BRIDGE_MAX_DROPPED_FOR_WAIT` (по умолчанию: `0`): максимально допустимое число dropped-сообщений перед деградацией в fallback;
- `AUTOCAD_MCP_EVENT_BRIDGE_OBJECT_EVENTS_ENABLED` (по умолчанию: `1`): желаемый режим object events в плагине (`AEB_OBJECT_EVENTS_ON/OFF`);
- `AUTOCAD_MCP_EVENT_BRIDGE_AUTOLOAD` (по умолчанию: `1`): при включенном bridge разрешает best-effort автозагрузку плагина (`NETLOAD`), если pipe недоступен;
- `AUTOCAD_MCP_EVENT_BRIDGE_AUTOLOAD_DLL` (опционально): явный путь к `AcadEventBridge.dll` для автозагрузки; иначе используется поиск в стандартных build-путях.

Запуск AutoCAD:

- `AUTOCAD_MCP_ACAD_EXE` (опционально): полный путь к `acad.exe` для явного запуска AutoCAD;
- `AUTOCAD_MCP_ACAD_ARGS` (опционально): дополнительные аргументы для `acad.exe`;
- `AUTOCAD_MCP_LAUNCH_WAIT_SEC` (по умолчанию: `30`): сколько ждать запуск AutoCAD перед повторной попыткой COM attach.

## Пример конфигурации Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "acad-cmd": {
      "command": "C:/path/to/project/.venv/Scripts/python.exe",
      "args": ["-m", "acad_cmd.server"]
    }
  }
}
```

Примечания:

- `.venv` не коммитится в git и не появится после clone; создайте ее локально (или запустите `start_server.bat`);
- если хотите автозапуск AutoCAD, задайте `AUTOCAD_MCP_ACAD_EXE`.

## Инструменты

Все инструменты возвращают JSON (FastMCP часто оборачивает результат как `{ "result": ... }`).

- `get_status()`
  - возвращает информацию о подключении (DWG label, `ACADVER`, HWND/PID при наличии) и детали default stream;
  - включает поля стабильности/деградации: `busy`, `stale`, `source`, `error_class`, `cmdactive`;
  - включает binding-поля против дрейфа версий: `locked_major`, `bound_progid`;
  - включает диагностику `event_bridge`: доступность/подключение, pipe/plugin metadata, heartbeat/queue state, reason деградации и optional service status.
- `send_command(command, timeout_sec, wait=true, poll_interval_sec=0.1)`
  - отправляет сырую команду в командную строку AutoCAD;
  - при `wait=true` сначала ожидает completion по bridge-событиям (если bridge доступен), иначе fallback на COM idle wait;
  - гарантирует default logfile stream и возвращает блок `log` с новым выводом и обновленным курсором;
  - возвращает диагностику ожидания (`wait_source`, `wait_completion_event`, `wait_completion_seq`, `wait_fallback_used`, `bridge_wait_prepare_issue`).
- `open_drawing(path, timeout_sec, read_only=false)`
  - открывает DWG и гарантированно переключает активный контекст на целевой документ (или возвращает ошибку по timeout);
  - возвращает `dwg_before`, `dwg`, `already_open`, `opened`, `activated`.
- `get_last_output(source=lastprompt|logfile)`
  - источник по умолчанию: `logfile`;
  - `logfile`: возвращает tail текущего default logfile stream (при необходимости автоматически запускает logfile stream);
  - `lastprompt`: читает `LASTPROMPT` (legacy/fallback).
- `start_logging(mode=logfile|lastprompt, logfile_path=null, reset=false)`
  - запускает stream и возвращает `{stream_id, cursor, ...}`;
  - `logfile` включает `LOGFILEMODE` и отслеживает `LOGFILENAME`;
  - `lastprompt` оставлен только для обратной совместимости;
  - если `logfile_path` не задан, сервер старается использовать текущий `LOGFILENAME` AutoCAD, чтобы снизить проблемы с путями.
- `get_new_output_since(stream_id, cursor, max_bytes=65536)`
  - читает добавленные байты логфайла и возвращает `{text, new_cursor, truncated}`.
- `stop_logging(stream_id)`
  - останавливает stream; best-effort отключает `LOGFILEMODE`, когда остановлен последний logfile stream, запущенный сервером.
- `load_lisp_file(path, timeout_sec, wait=true)`
  - отправляет `(load "...")` (путь нормализуется для AutoCAD);
  - при `wait=true` сначала использует LISP completion events через bridge, затем fallback на COM.
- `run_lisp(expr, timeout_sec, wait=true)`
  - выполняет AutoLISP выражение/скрипт через `SendCommand` со start/end marker в истории команд;
  - при `wait=true` сначала использует LISP completion events через bridge, затем fallback на COM.
- `selection(timeout_sec, prompt=null, filter=null, max_objects=null, alert_message=null)`
  - возвращает текущий выбор объектов (PickFirst); если выбор пуст, инициирует интерактивный выбор;
  - если задан `alert_message` и нужен интерактивный выбор, показывает стандартный AutoCAD `alert`;
  - возвращает для каждого объекта только `handle` и `type`.
- `dict_list/dict_keys/dict_xrecord_get/dict_xrecord_set/dict_xrecord_delete/dict_delete(..., timeout_sec)`
  - инструменты работы со словарями требуют явный `timeout_sec`;
  - диапазон timeout для waiting-инструментов: `0.1..1800` секунд.

## Troubleshooting

- Загрузка AutoLISP-файлов может блокироваться настройками безопасности AutoCAD:
  - добавьте папку с `.lsp` в **Trusted Locations**;
  - проверьте поведение `SECURELOAD` (не ослабляйте безопасность глобально в production).
- При COM-ошибке "callee busy" сервер выполняет ретраи с backoff.
- `LOGFILEMODE`/`LOGFILENAME` пишет файл в codepage AutoCAD; декодирование выполняется через preferred encoding Windows с fallback.

## Разработка / smoke-тесты

- `scripts/mcp_smoketest_stdio.py`: поднимает сервер по stdio, листит инструменты, запускает logging, выполняет простое LISP-выражение;
- `scripts/mcp_sanity_acadver.py`: прямой COM sanity-check, что `LOGFILEMODE` растет после `(getvar 'ACADVER)`;
- `scripts/mcp_baseline_capture.py`: baseline-capture (`get_status`, `send_command`, `run_lisp`, `LOGFILEMODE`) с JSON-отчетом в `out/baseline/`.

Запуск unit-тестов:

```bat
python -m unittest discover -s tests -p "test_*.py"
```

Запуск integration-тестов (нужен реальный AutoCAD):

```bat
set ACAD_MCP_RUN_INTEGRATION=1
python -m unittest discover -s tests -p "test_*.py"
```

## Документы репозитория

- гайд по участию: `CONTRIBUTING.md`
- security policy: `SECURITY.md`
- лицензия: `LICENSE`
