# Roadmap: in-process .NET event bridge for `acad-cmd`

## Current Progress (2026-04-22)

- Completed: steps 1 to 21 from `docs/acad-event-bridge-codex-task-roadmap.md`.
- Verified in AutoCAD runtime:
  - `NETLOAD` of `AcadEventBridge.dll`;
  - pipe `hello`;
  - pipe `heartbeat`.
  - Python `EventBridgeClient` receives `hello` + `heartbeat` and exposes snapshot state.
  - `event_state.py` is in place and tracks last seq/busy/depth/active doc/heartbeat plus latest `command_*` event fields.
  - `get_status()` now lazily connects to bridge by AutoCAD PID and returns bridge health fields.
  - `command_waiter.py` now supports event-first completion waiting and fallback to COM idle wait.
  - `send_command(wait=true)` now uses `CommandWaiter` (event-first + fallback).
  - `run_lisp(wait=true)` and `load_lisp_file(wait=true)` now use LISP-focused waiter completion profile (`lisp_ended`/`lisp_cancelled`) with COM fallback.
  - hardening for disconnect/overload added: heartbeat-timeout guard, dropped-count guard, reconnect handling and fallback diagnostics.
  - final MVP smoke pass completed: fallback/no-bridge, NETLOAD, command, LISP, document switching, disconnect/recovery.
- Step 21 decision (research result): keep COM as the primary command execution path; do not introduce plugin-side command execution at this stage.
- Next planned step: continue hardening/observability on the current hybrid model (COM execution + bridge events/status).

## 1. Цель

Сделать гибридное соединение с AutoCAD:

- внешний MCP-сервер на Python остаётся точкой входа для LLM/Codex;
- внутри AutoCAD загружается `C#/.NET` плагин;
- плагин отдаёт **структурированные события** о состоянии AutoCAD;
- `LOGFILEMODE` остаётся только как **fallback/debug-канал**, а не как основной источник состояния.

Главная практическая цель: перестать определять завершение команды через разбор текстового лога и сделать более надёжную синхронизацию `command start/end/cancel/fail`, документов и выбора.

---

## 2. Что есть сейчас в `acad-cmd`

Текущая схема:

- Python подключается к AutoCAD через COM.
- Команды отправляются через `SendCommand`.
- Ожидание завершения идёт через polling: `GetAcadState().IsQuiescent` + `CMDACTIVE`.
- Быстрый текст берётся через `LASTPROMPT`.
- Полный текст консоли добирается через `LOGFILEMODE` / `LOGFILENAME`.

Проблемы текущей схемы:

- лог консоли неструктурированный;
- бывают проблемы с путями/кодировкой и tail лог-файла;
- завершение команды определяется косвенно;
- COM приходится ретраить при `callee busy`;
- lifecycle команды и lifecycle документа смешаны с текстовым выводом.

---

## 3. Архитектурное решение

### 3.1. Решение для MVP

Использовать **`.NET`-плагин внутри AutoCAD**, а не `ObjectARX`, потому что:

- для событий документов/команд/выбора managed API уже достаточно;
- реализация и сопровождение будут проще, чем у native ARX;
- `ObjectARX` оставляем как запасной вариант только если managed API реально не хватит.

### 3.2. Границы MVP

В MVP:

- команды по-прежнему отправляются из Python через COM;
- `.NET`-плагин используется как **event bridge**;
- ожидание завершения команды переключается на события, а COM-polling остаётся fallback;
- `LOGFILEMODE` сохраняется для debug и случаев, где нужен именно текст командной строки.

### 3.3. Главные решения без обсуждений

1. **Не делать full rewrite** текущего сервера.
2. **Не заменять COM целиком** на первом этапе.
3. **Не убирать `LOGFILEMODE` полностью**.
4. **Не делать ObjectARX в MVP**.
5. **Не включать object-level events по умолчанию** — они слишком шумные.
6. **Считать, что в MVP одновременно ждём завершения только одной команды**.

---

## 4. Целевая схема приложения

```text
LLM / Codex
    ↓
MCP client
    ↓
Python server (acad-cmd)
    ├─ COM → AutoCAD (send_command, get/set variables, fallback)
    └─ Named Pipe → .NET plugin inside AutoCAD (events, status)

AutoCAD process
    └─ AcadEventBridge.dll
         ├─ registers document/application events
         ├─ pushes events to bounded queue
         └─ streams NDJSON over named pipe
```

### Ответственность слоёв

**Python / MCP слой**

- управляет жизненным циклом соединения;
- при необходимости загружает .NET-плагин;
- отправляет команды через COM;
- ждёт completion по event stream;
- при проблемах откатывается на старую схему (`IsQuiescent`/`CMDACTIVE`/`LOGFILEMODE`).

**.NET plugin слой**

- стартует внутри процесса AutoCAD;
- регистрирует события для текущих и новых документов;
- переводит события в компактный JSON;
- публикует события наружу через локальный IPC;
- не содержит тяжёлой бизнес-логики.

---

## 5. IPC: что использовать

### Выбор

Использовать **Named Pipe** на Windows.

Почему именно так:

- локальный IPC, без сети;
- хорошо подходит для одного AutoCAD процесса;
- не зависит от файловой системы, в отличие от `LOGFILEMODE`;
- легко подключить из Python и C#.

### Рекомендация по роли сторон

Для MVP плагин внутри AutoCAD должен быть **pipe server**, а Python — **pipe client**.

Причина:

- у Python уже есть способ узнать PID окна/процесса AutoCAD через COM/`HWND`;
- имя pipe можно строить как `acad-event-bridge-<pid>`;
- это упрощает поиск нужного AutoCAD экземпляра.

### Формат сообщений

Использовать **NDJSON** (`1 JSON object = 1 line`, кодировка `UTF-8`).

Не надо придумывать сложный бинарный протокол в первой версии.

---

## 6. Минимальный контракт протокола

### 6.1. `hello`

Первое сообщение от плагина после подключения:

```json
{"type":"hello","protocol":1,"plugin":"AcadEventBridge","version":"0.1.0","pid":12345}
```

### 6.2. `event`

Основное сообщение:

```json
{
  "type": "event",
  "seq": 101,
  "ts": "2026-04-21T10:12:33.456Z",
  "session_id": "acad-12345",
  "source": "document",
  "doc_id": "b6fb6a30-9f4e-4b7e-b1f9-a0ad8d43d9b1",
  "doc_name": "plan.dwg",
  "doc_path": "C:\\work\\plan.dwg",
  "event": "command_will_start",
  "payload": {
    "name": "LINE"
  }
}
```

### 6.3. `heartbeat`

Периодически, например раз в 2 секунды:

```json
{"type":"heartbeat","seq":150,"busy":true,"active_doc_id":"...","queue_depth":0}
```

### 6.4. Резерв под phase 2

Сразу зарезервировать формат:

- `request`
- `response`

Но **не реализовывать request/response в первой итерации**, если не понадобится.

---

## 7. Какие события нужны в MVP

### Обязательные

#### Document / command lifecycle

- `bridge_ready`
- `document_created`
- `document_activated`
- `document_destroyed`
- `command_will_start`
- `command_ended`
- `command_cancelled`
- `command_failed`
- `lisp_will_start`
- `lisp_ended`
- `lisp_cancelled`
- `unknown_command`
- `implied_selection_changed`

### Не включать по умолчанию

#### Database / object events

Оставить как opt-in флаг:

- `object_appended`
- `object_modified`
- `object_erased`

Причина: при обычной работе они могут генерировать очень много событий.

---

## 8. Правила реализации обработчиков событий

Это критично.

1. Обработчик события делает **минимум работы**.
2. Обработчик **не должен** показывать UI и запрашивать ввод.
3. Обработчик **не должен** запускать команды через `SendStringToExecute`.
4. Обработчик **не должен** менять объект, который сам вызвал событие.
5. Обработчик только:
   - собирает простой payload;
   - кладёт событие в очередь;
   - сразу выходит.
6. Если позже понадобится менять БД из modeless/session/COM-контекста — делать это отдельно и с `DocumentLock`.

Практическое следствие: никакой тяжёлой сериализации, файловых операций, COM-вызовов наружу и логики ожидания прямо из event handler.

---

## 9. Внутренняя структура .NET-плагина

Рекомендуемая структура проекта:

```text
/AcadEventBridge
  AcadEventBridge.csproj
  EntryPoint.cs
  BridgeHost.cs
  PipeServer.cs
  EventRegistrar.cs
  EventQueue.cs
  EventModels.cs
  StateTracker.cs
  DebugCommands.cs
```

### Роли классов

**EntryPoint**

- `IExtensionApplication.Initialize()`
- старт `BridgeHost`
- регистрация текущих документов и подписка на `DocumentCreated`

**BridgeHost**

- связывает `EventRegistrar`, `StateTracker`, `EventQueue`, `PipeServer`

**EventRegistrar**

- подписывает события документа;
- подписывает уже открытые документы;
- подписывает новые документы;
- снимает подписки при выгрузке

**StateTracker**

- хранит `seq`;
- хранит `doc_id` для каждого документа;
- считает `busy` / `commandDepth` / `lispDepth`;
- знает активный документ

**EventQueue**

- `ConcurrentQueue<BridgeEvent>` + bounded buffer;
- при переполнении удаляет старые элементы и увеличивает `dropped_count`

**PipeServer**

- поднимает named pipe;
- шлёт `hello`, `event`, `heartbeat`;
- не обращается к AutoCAD API напрямую

**DebugCommands**

необязательно, но полезно:

- `AEB_STATUS`
- `AEB_PING`
- `AEB_DIAG`

---

## 10. Внутренняя структура Python-части

Рекомендуемое добавление в `acad-cmd`:

```text
src/acad_cmd/
  bridge_plugin_client.py
  event_state.py
  command_waiter.py
```

### Роли модулей

**bridge_plugin_client.py**

- подключение к named pipe;
- чтение NDJSON;
- reconnect;
- health state (`connected`, `last_heartbeat`, `plugin_version`)

**event_state.py**

- хранение последнего `seq`;
- состояние `busy`;
- последний активный документ;
- последние события по командам/документам

**command_waiter.py**

- логика ожидания завершения `send_command`;
- сначала ждёт по event stream;
- при сбое уходит в старый COM fallback

---

## 11. Алгоритм `send_command` после внедрения bridge

### MVP алгоритм

1. Python получает текущий `seq` из `event_state`.
2. Python отправляет команду через COM `SendCommand`.
3. Python ждёт события после `seq`.
4. Если появился `command_will_start` или `lisp_will_start`, считаем команду принятой.
5. Далее ждём:
   - `command_ended`, или
   - `command_cancelled`, или
   - `command_failed`, или
   - `lisp_ended`, или
   - `lisp_cancelled`
6. Если completion не пришёл за timeout:
   - если `busy=true`, вернуть `needs_input=true`;
   - если bridge не отвечает, перейти к старой схеме `IsQuiescent + CMDACTIVE`.

### Упрощение MVP

Для первой версии считаем, что сервер ждёт **одну активную команду одновременно**. Это допустимо для MCP-сценария и сильно упрощает корреляцию.

### Что не решаем в MVP

Мы **не** решаем строгую корреляцию вложенных команд/ручных пользовательских действий.

Если позже понадобится строгая корреляция, делать phase 2:

- отдельный bridge-command;
- `request_id`;
- или выполнение команды через плагин, а не напрямую через COM.

---

## 12. Как bridge должен влиять на существующие MCP tools

### `get_status()`

Добавить блок:

```json
"event_bridge": {
  "available": true,
  "connected": true,
  "pipe_name": "acad-event-bridge-12345",
  "plugin_version": "0.1.0",
  "last_heartbeat": "..."
}
```

### `send_command()`

- основное ожидание завершения перевести на event stream;
- текущий COM wait сохранить как fallback;
- `lastprompt` можно сохранить как быстрый текстовый хвост;
- `logfile` вернуть только если реально запущен logging mode.

### `start_logging()` / `get_new_output_since()`

Оставить как есть, но позиционировать как:

- debug;
- текстовый transcript;
- fallback при отсутствии bridge.

---

## 13. Порядок внедрения

### Этап 1. Скелет плагина

Сделать минимальный `.NET`-плагин, который:

- грузится через `NETLOAD`;
- поднимает named pipe;
- шлёт `hello` и `heartbeat`.

**Критерий готовности:** Python может подключиться к pipe и увидеть `hello`.

### Этап 2. Document events

Добавить:

- регистрация событий на уже открытых документах;
- `DocumentCreated` для новых документов;
- `document_created`, `document_activated`, `document_destroyed`.

**Критерий готовности:** открытие/активация/закрытие DWG видны в Python.

### Этап 3. Command/LISP events

Добавить:

- `command_will_start`
- `command_ended`
- `command_cancelled`
- `command_failed`
- `lisp_will_start`
- `lisp_ended`
- `lisp_cancelled`

**Критерий готовности:** простой `send_command()` можно довести до completion только на событиях.

### Этап 4. Интеграция в Python server

Добавить:

- `bridge_plugin_client.py`
- `event_state.py`
- `command_waiter.py`
- блок `event_bridge` в `get_status()`

**Критерий готовности:** если bridge доступен, `send_command()` сначала ждёт события, а не `IsQuiescent`.

### Этап 5. Fallback logic

Добавить правила:

- если pipe не поднят — работаем по старой схеме;
- если heartbeat пропал — деградируем на COM polling;
- `LOGFILEMODE` не ломается и остаётся доступен.

**Критерий готовности:** старый функционал продолжает работать без плагина.

### Этап 6. Selection + optional object events

Добавить:

- `implied_selection_changed`
- опциональные object events за feature flag

**Критерий готовности:** выбор синхронизируется, объектные события можно включить вручную.

### Этап 7. Автозагрузка плагина

После стабилизации:

- либо автоматический `NETLOAD` из Python;
- либо нормальная упаковка/автозагрузка как AutoCAD plugin bundle.

**Критерий готовности:** bridge поднимается без ручных действий пользователя.

---

## 14. Что считать успехом проекта

Решение считается успешным, если выполнены все пункты:

- `send_command()` больше не зависит от `LOGFILEMODE` для определения completion;
- статус команды определяется по событиям `start/end/cancel/fail`;
- при нескольких открытых DWG события правильно привязаны к документу;
- отсутствие bridge не ломает текущий `acad-cmd`;
- лог-файл остаётся доступным как вспомогательный канал;
- событие не вызывает зависаний/циклов/интерактивных ошибок внутри AutoCAD.

---

## 15. Риски и ограничения

### Ожидаемые ограничения

- порядок событий нельзя считать абсолютно детерминированным;
- во время modal dialog AutoCAD события могут не приходить;
- ручные действия пользователя могут смешиваться с командами MCP;
- object events могут быть слишком шумными;
- строгая корреляция вложенных команд в MVP не решается.

### Как снизить риски

- держать event handlers максимально тупыми;
- все тяжёлые действия вынести в очередь/pipe thread;
- держать bounded queue и heartbeat;
- включать object events только флагом;
- оставить COM и `LOGFILEMODE` как fallback.

---

## 16. Что не делать сейчас

Не тратить время в первой реализации на:

- ObjectARX/C++;
- полный отказ от COM;
- сложный двусторонний RPC;
- database modifications прямо из event handlers;
- полный replacement текстового console transcript.

---

## 17. Практические подсказки по реализации

- Начать с версии, где bridge только **слушает** и публикует события.
- Команду в AutoCAD по-прежнему отправлять через текущий COM-слой.
- Для совместимости сначала ориентироваться на тот runtime AutoCAD, который уже используется в работе.
- Если нужна поддержка нескольких линий AutoCAD, заранее продумать разные сборки плагина под соответствующие runtime/SDK.
- Для отладки сделать очень маленький ручной smoke-test: `NETLOAD` → pipe connect → одна команда → событие `start/end`.

---

## 18. Краткое техническое ТЗ для Codex

Ниже формулировка, от которой можно отталкиваться в следующих сессиях:

> Нужно расширить текущий `acad-cmd` гибридной схемой: оставить внешний Python MCP-сервер и COM-отправку команд, но добавить in-process C#/.NET plugin для AutoCAD, который публикует document/command/LISP/selection events через local named pipe в формате NDJSON UTF-8. Python-сервер должен уметь подключаться к bridge, хранить event state, использовать события для ожидания completion в `send_command()` и автоматически деградировать на старую схему (`IsQuiescent`, `CMDACTIVE`, `LASTPROMPT`, `LOGFILEMODE`) при недоступности bridge. MVP не должен требовать ObjectARX, полного отказа от COM и строгой корреляции вложенных команд.

---

## 19. Полезные источники

- `acad-cmd` repo: <https://github.com/EmptyKot/acad-cmd>
- AutoCAD .NET: Out-of-Process vs In-Process
  <https://help.autodesk.com/view/OARX/2027/ENU/?guid=GUID-C8C65D7A-EC3A-42D8-BF02-4B13C2EA1A4B>
- AutoCAD .NET: Handle Document Events
  <https://help.autodesk.com/view/ACD/2027/ENU/?caas=caas%2Fdocumentation%2FACD%2F2014%2FENU%2Ffiles%2FGUID-F432E285-8B94-4ACD-A186-89E1218DEC07-htm.html>
- AutoCAD .NET: Guidelines for Event Handlers
  <https://help.autodesk.com/view/ACD/2027/ENU/?caas=caas%2Fdocumentation%2FACD%2F2014%2FENU%2Ffiles%2FGUID-FE7D58D5-28A0-4C98-A876-D4D48F06D0B2-htm.html>
- AutoCAD .NET: Lock and Unlock a Document
  <https://help.autodesk.com/cloudhelp/2026/ITA/OARX-DevGuide-Managed/files/GUID-A2CD7540-69C5-4085-BCE8-2A8ACE16BFDD.htm>
- AutoCAD ActiveX events overview
  <https://help.autodesk.com/cloudhelp/2023/CHS/AutoCAD-ActiveX/files/GUID-07494559-EA6C-4D84-B494-A49930C67E91.htm>
