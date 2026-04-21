# AcadEventBridge — implementation spec for Codex

## Implementation Status (2026-04-21)

- Step 1: baseline for current `acad-cmd` captured.
- Step 2: Python feature flag `event_bridge_enabled` added (default off).
- Step 3: `plugins/AcadEventBridge` skeleton created and buildable.
- Step 4: named pipe server implemented, `hello` NDJSON message emitted.
- Step 5: periodic `heartbeat` NDJSON implemented.
- Step 6: bounded `EventQueue` + `StateTracker` integrated; heartbeat is built from tracker snapshot and sent through queue-backed writer path.
- Step 7: `DocumentCollection` events implemented and published to stream: `document_created`, `document_activated`, `document_destroyed`.

## 1. Что именно делаем

Нужно расширить текущий `acad-cmd` гибридной схемой:

- `Python MCP` остаётся внешней точкой входа;
- команды в MVP по-прежнему уходят в AutoCAD через `COM / SendCommand`;
- внутри AutoCAD работает `C#/.NET` плагин `AcadEventBridge`;
- плагин публикует структурированные события наружу через `Named Pipe`;
- `LOGFILEMODE` и текущий COM-polling остаются fallback-каналом.

Итог MVP:

1. `send_command(wait=true)` ждёт completion в первую очередь по event stream.
2. Если bridge недоступен или умер — сервер автоматически откатывается на старую схему.
3. Публичные MCP tools не ломаются.

---

## 2. На что опираемся в текущем `acad-cmd`

Текущие точки интеграции в репозитории:

```text
src/acad_cmd/
  autocad_bridge.py   # COM connect/send_command/get_variable/wait_for_idle
  lisp.py             # helpers для Lisp-команд
  output_log.py       # LOGFILEMODE / LOGFILENAME streams
  server.py           # MCP tools и верхний orchestration слой
  session_log.py      # audit log
```

Главное правило MVP: **не переписывать сервер заново**. Добавляем новый канал состояния рядом с текущим COM-слоем.

---

## 3. Целевая архитектура

```text
LLM / Codex
    ↓
MCP client
    ↓
Python server (acad-cmd)
    ├─ COM → AutoCAD (send_command / vars / fallback)
    └─ Named Pipe client → AcadEventBridge.dll

AutoCAD.exe
    └─ AcadEventBridge.dll
         ├─ DocumentCollection events
         ├─ Document events
         ├─ internal bounded queue
         └─ NDJSON writer over named pipe
```

### Разделение ответственности

**Python слой**

- подключается к AutoCAD по COM;
- знает PID/окно нужного экземпляра AutoCAD;
- подключается к named pipe нужного процесса;
- отправляет команды через COM;
- ждёт completion по событиям;
- при отказе bridge уходит на `wait_for_idle()` и, если нужно, на `LOGFILEMODE`.

**.NET слой**

- грузится через `NETLOAD`;
- регистрирует события для текущих и новых документов;
- превращает события AutoCAD в компактные JSON-сообщения;
- пишет их в pipe;
- не выполняет тяжёлой бизнес-логики.

---

## 4. Границы MVP

В MVP делаем только это:

- document lifecycle;
- command lifecycle;
- LISP lifecycle;
- selection change;
- heartbeat + status;
- fallback на старый код.

В MVP **не делаем**:

- ObjectARX;
- полный отказ от COM;
- object events по умолчанию;
- двусторонний RPC поверх pipe;
- строгую корреляцию вложенных команд;
- модификацию БД из event handlers.

Принятое упрощение: **одновременно ждём только одну активную MCP-команду**.

---

## 5. Структура .NET-плагина

Рекомендуемый каталог:

```text
plugins/
  AcadEventBridge/
    AcadEventBridge.csproj
    EntryPoint.cs
    BridgeHost.cs
    PipeServer.cs
    EventRegistrar.cs
    StateTracker.cs
    EventQueue.cs
    EventModels.cs
    DebugCommands.cs
```

### Назначение файлов

**EntryPoint.cs**

- реализует `IExtensionApplication`;
- в `Initialize()` запускает `BridgeHost`;
- в `Terminate()` делает best-effort cleanup.

**BridgeHost.cs**

- связывает `PipeServer`, `EventRegistrar`, `StateTracker`, `EventQueue`;
- создаёт pipe name из PID процесса;
- публикует `bridge_ready`.

**PipeServer.cs**

- поднимает `NamedPipeServerStream`;
- шлёт `hello`, `event`, `heartbeat`;
- не вызывает AutoCAD API напрямую;
- работает с уже готовыми DTO из очереди.

**EventRegistrar.cs**

- подписывает `DocumentCollection` и `Document` события;
- регистрирует уже открытые документы;
- регистрирует новые документы;
- снимает подписки при остановке.

**StateTracker.cs**

- хранит глобальный `seq`;
- хранит `doc_id` на документ;
- ведёт `command_depth`, `lisp_depth`, `busy`;
- знает `active_doc_id`.

**EventQueue.cs**

- bounded queue;
- при переполнении удаляет старые сообщения;
- увеличивает `dropped_count`.

**EventModels.cs**

- DTO для `hello`, `event`, `heartbeat`;
- сериализация в компактный JSON.

**DebugCommands.cs**

Команды только для ручной диагностики:

- `AEB_STATUS`
- `AEB_PING`
- `AEB_DIAG`

---

## 6. Какие события подписывать

### 6.1. На уровне `DocumentCollection`

Нужны события для жизненного цикла документов:

- `DocumentCreated` → `document_created`
- `DocumentBecameCurrent` → `document_activated`
- `DocumentToBeDestroyed` → `document_destroyed`

Это отдельный слой от событий самого `Document`.

### 6.2. На уровне `Document`

Нужны события команды/скрипта:

- `CommandWillStart` → `command_will_start`
- `CommandEnded` → `command_ended`
- `CommandCancelled` → `command_cancelled`
- `CommandFailed` → `command_failed`
- `LispWillStart` → `lisp_will_start`
- `LispEnded` → `lisp_ended`
- `LispCancelled` → `lisp_cancelled`
- `UnknownCommand` → `unknown_command`
- `ImpliedSelectionChanged` → `implied_selection_changed`

### 6.3. Не включать по умолчанию

Под флагом позже:

- `object_appended`
- `object_modified`
- `object_erased`

Причина: слишком шумно для первой версии.

---

## 7. Контракт pipe-протокола

### 7.1. Pipe name

```text
\\.\pipe\acad-event-bridge-<pid>
```

Где `<pid>` — PID того AutoCAD процесса, к которому уже подключён Python.

### 7.2. Формат

- `UTF-8`
- `NDJSON`
- одно JSON-сообщение на строку

### 7.3. Сообщения

#### hello

```json
{"type":"hello","protocol":1,"plugin":"AcadEventBridge","version":"0.1.0","pid":12345}
```

#### event

```json
{
  "type": "event",
  "seq": 101,
  "ts": "2026-04-21T10:12:33.456Z",
  "source": "document",
  "doc_id": "b6fb6a30-9f4e-4b7e-b1f9-a0ad8d43d9b1",
  "doc_name": "plan.dwg",
  "doc_path": "C:\\work\\plan.dwg",
  "event": "command_will_start",
  "payload": {"name": "LINE"}
}
```

#### heartbeat

```json
{
  "type": "heartbeat",
  "seq": 150,
  "ts": "2026-04-21T10:12:35.000Z",
  "busy": true,
  "command_depth": 1,
  "lisp_depth": 0,
  "active_doc_id": "b6fb6a30-9f4e-4b7e-b1f9-a0ad8d43d9b1",
  "queue_depth": 0,
  "dropped_count": 0
}
```

### 7.4. Обязательные поля

У каждого `event` должны быть:

- `seq` — монотонный номер;
- `ts` — UTC ISO8601;
- `event` — нормализованное имя события;
- `doc_id` — стабильный ID документа внутри жизни процесса;
- `payload` — только короткие поля без тяжёлых объектов.

---

## 8. Правила для event handlers

Жёсткие правила:

1. handler делает минимум работы;
2. handler не вызывает UI;
3. handler не вызывает `SendStringToExecute`;
4. handler не делает долгих I/O операций;
5. handler не меняет объект, который сам вызвал событие;
6. handler только собирает payload и кладёт сообщение в очередь.

Следствие: все тяжёлые действия выполняются вне handler-а.

---

## 9. Как должен выглядеть Python-слой

### 9.1. Новые файлы

```text
src/acad_cmd/
  bridge_plugin_client.py
  event_state.py
  command_waiter.py
```

### 9.2. Что делать в новых файлах

**bridge_plugin_client.py**

- подключение к named pipe;
- чтение NDJSON в фоне;
- reconnect;
- heartbeat health;
- хранение `plugin_version`, `pid`, `last_heartbeat`.

**event_state.py**

- хранение последних событий;
- `last_seq`;
- `busy`, `command_depth`, `lisp_depth`;
- `active_doc_id`;
- последние `command_*` события.

**command_waiter.py**

- логика ожидания completion для `send_command()`;
- сначала ждёт event stream;
- если bridge сломан — вызывает старый `wait_for_idle()`.

---

## 10. Что менять в текущих Python-файлах

### `src/acad_cmd/autocad_bridge.py`

Добавить:

- `get_hwnd()`
- `get_pid()`
- `get_pipe_name()`
- опционально `is_bridge_candidate_ready()`

Не ломать существующие методы:

- `connect()`
- `send_command()`
- `wait_for_idle()`
- `get_last_prompt()`

### `src/acad_cmd/server.py`

Добавить:

- инициализацию `EventBridgeClient`;
- lazy connect к pipe при первом `get_status()` или `send_command()`;
- новый блок `event_bridge` в `get_status()`;
- использование `CommandWaiter` в `send_command()`, `run_lisp()`, `load_lisp_file()`.

### `src/acad_cmd/output_log.py`

Не ломать API.

Оставить как:

- debug transcript;
- fallback;
- источник сырых текстовых хвостов.

### `src/acad_cmd/session_log.py`

Опционально добавить в audit:

- `event_bridge_connected`
- `event_bridge_pid`
- `event_bridge_seq`

Но это не критично для MVP.

---

## 11. Алгоритм `send_command()` после внедрения bridge

```text
1. Убедиться, что COM соединение живо.
2. Попробовать подключиться к event bridge для текущего PID AutoCAD.
3. Считать текущий last_seq.
4. Отправить команду через COM SendCommand.
5. Если bridge подключён:
   - ждать события после last_seq;
   - preferred path: увидеть start и затем completion;
   - accepted completion events:
       command_ended
       command_cancelled
       command_failed
       lisp_ended
       lisp_cancelled
6. Если completion не пришёл за timeout:
   - если bridge жив и busy=true -> вернуть needs_input=true;
   - иначе fallback на wait_for_idle().
7. Если bridge не подключён с самого начала:
   - сразу использовать старую схему.
```

### Важное упрощение

В первой версии корреляция только по времени и `seq`, а не по `request_id`.

Это допустимо, пока одновременно сервер ведёт только одну MCP-команду.

---

## 12. Как должен выглядеть `get_status()`

Добавить блок:

```json
{
  "event_bridge": {
    "available": true,
    "connected": true,
    "pipe_name": "acad-event-bridge-12345",
    "plugin_version": "0.1.0",
    "last_heartbeat": "2026-04-21T10:12:35Z",
    "busy": false,
    "last_seq": 150
  }
}
```

Если bridge не найден:

```json
{
  "event_bridge": {
    "available": false,
    "connected": false
  }
}
```

---

## 13. Загрузка плагина

### MVP

- плагин грузится вручную через `NETLOAD`;
- Python только пытается подключиться к pipe;
- отсутствие плагина не считается ошибкой сервера.

### После стабилизации

Выбрать один из вариантов:

1. app bundle / автозагрузка AutoCAD-плагина;
2. scripted load из Python, если на вашей версии AutoCAD это стабильно;
3. отдельный установщик.

Для первой итерации **не связывать успех проекта с автозагрузкой**.

---

## 14. Порядок внедрения по коммитам

### Commit 1 — skeleton plugin

Сделать:

- `AcadEventBridge.csproj`
- `EntryPoint`
- `PipeServer`
- `hello`
- `heartbeat`

Критерий:

- `NETLOAD` работает;
- Python видит `hello` и heartbeat.

### Commit 2 — Python pipe client

Сделать:

- `bridge_plugin_client.py`
- `event_state.py`
- блок `event_bridge` в `get_status()`

Критерий:

- `get_status()` показывает состояние bridge.

### Commit 3 — document lifecycle

Сделать:

- `DocumentCreated`
- `DocumentBecameCurrent`
- `DocumentToBeDestroyed`

Критерий:

- открытие/переключение/закрытие DWG видно в Python.

### Commit 4 — command + lisp lifecycle

Сделать:

- `CommandWillStart`
- `CommandEnded`
- `CommandCancelled`
- `CommandFailed`
- `LispWillStart`
- `LispEnded`
- `LispCancelled`
- `UnknownCommand`

Критерий:

- Python видит start/end/cancel/fail.

### Commit 5 — интеграция в `send_command()`

Сделать:

- `command_waiter.py`
- event-first ожидание completion;
- fallback на `wait_for_idle()`.

Критерий:

- `send_command(wait=true)` работает с bridge и без bridge.

### Commit 6 — selection + hardening

Сделать:

- `ImpliedSelectionChanged`
- bounded queue
- dropped counter
- reconnect logic

Критерий:

- выбор отражается, pipe переживает disconnect.

---

## 15. Минимальные smoke-тесты

### Smoke 1 — загрузка плагина

1. Собрать DLL.
2. В AutoCAD выполнить `NETLOAD`.
3. Убедиться, что Python видит `hello`.

### Smoke 2 — жизненный цикл команды

1. Подключиться MCP-сервером.
2. Выполнить простую неинтерактивную команду.
3. Проверить, что пришли `command_will_start` и одно из completion-событий.

### Smoke 3 — LISP lifecycle

1. Выполнить `run_lisp()`.
2. Проверить `lisp_will_start` и `lisp_ended`.

### Smoke 4 — fallback

1. Не грузить plugin.
2. Проверить, что `send_command(wait=true)` всё ещё работает по старой схеме.

### Smoke 5 — document switching

1. Открыть 2 DWG.
2. Переключиться между ними.
3. Убедиться, что приходит `document_activated` и корректный `doc_id`.

---

## 16. Что считать готовым MVP

MVP готов, если:

- bridge можно загрузить через `NETLOAD`;
- Python умеет подключаться к pipe нужного AutoCAD PID;
- `get_status()` показывает health bridge;
- `send_command()` умеет ждать completion по событиям;
- при падении bridge сервер продолжает работать по старому пути;
- `LOGFILEMODE` по-прежнему доступен как debug/fallback.

---

## 17. Что оставить на phase 2

- request/response поверх pipe;
- выполнение команд через сам плагин, а не через COM;
- строгая корреляция по `request_id`;
- object events под feature flag;
- app bundle packaging;
- richer status API внутри плагина.

---

## 18. Короткое ТЗ для Codex

Нужно внести в `acad-cmd` гибридный event bridge для AutoCAD. Текущий COM-слой и `LOGFILEMODE` остаются. Добавить .NET plugin `AcadEventBridge`, загружаемый через `NETLOAD`, который публикует `DocumentCollection` и `Document` события в локальный named pipe `\\.\pipe\acad-event-bridge-<pid>` в формате UTF-8 NDJSON. На Python-стороне добавить pipe client, event state и command waiter. `send_command(wait=true)` должен сначала ждать completion по event stream, а при недоступности bridge автоматически деградировать на текущий `wait_for_idle()` и существующую текстовую схему.

---

## 19. Источники

- `acad-cmd` repo: https://github.com/EmptyKot/acad-cmd
- AutoCAD .NET: Handle Document Events
  https://help.autodesk.com/view/ACD/2027/ENU/?caas=caas%2Fdocumentation%2FACD%2F2014%2FENU%2Ffiles%2FGUID-F432E285-8B94-4ACD-A186-89E1218DEC07-htm.html
- AutoCAD .NET: Handle DocumentCollection Events
  https://help.autodesk.com/view/OARX/2027/CHT/?guid=GUID-E619BB54-D531-4640-BB74-B61E6CA13238
- AutoCAD .NET: Guidelines for Event Handlers
  https://help.autodesk.com/view/ACD/2027/ENU/?caas=caas%2Fdocumentation%2FACD%2F2014%2FENU%2Ffiles%2FGUID-FE7D58D5-28A0-4C98-A876-D4D48F06D0B2-htm.html
- AutoCAD .NET: IExtensionApplication / Initialize / Terminate
  https://help.autodesk.com/view/OARX/2027/CSY/?guid=GUID-FA3B4125-F7BD-4E89-969F-9DCC90AC6977
- AutoCAD .NET: NETLOAD / Load an Assembly
  https://help.autodesk.com/view/OARX/2027/HUN/?guid=GUID-4EB83A6B-9903-4BF7-9F19-767A4D419CE3
