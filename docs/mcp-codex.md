# MCP-интерфейс MS Project Agent

`tools/mcp_server.py` — локальный STDIO-сервер без внешних Python-зависимостей.
Он предоставляет Codex пять инструментов поверх единого Schedule IR:

| Инструмент | Назначение | Изменяет файлы |
| --- | --- | --- |
| `schedule_summary` | Компактная сводка IR | нет |
| `schedule_validate_ir` | Проверка контракта, WBS, дат и связей IR | нет |
| `schedule_build` | Новый Excel ГРП и IR из JSON ТЭП | да, только новые файлы |
| `mpp_export` | Новый MPP из валидного IR | да, только новый MPP |
| `mpp_validate` | Чтение MPP без сохранения, проверка сети и сверка с IR | нет; опционально новый JSON-отчёт |

## Ограничения безопасности MVP

- доступны только пути внутри `D:\Claude\ClaudeVS\agent1`;
- существующие Excel, IR, MPP и JSON-отчёты не перезаписываются;
- перед экспортом MPP автоматически проверяется Schedule IR;
- `mpp_validate` закрывает Microsoft Project с режимом «не сохранять»;
- редактирование и удаление задач существующего MPP не поддерживаются.

## Подключение в Codex

В проект уже добавлен `.codex/config.toml` с локальным MCP-сервером типа
**STDIO**. Codex загружает эту конфигурацию, когда проект отмечен доверенным.
Проверка выполняется из корня проекта:

```powershell
codex mcp list
```

В списке должен появиться включённый сервер `agent1-ms-project`. Его параметры:

- имя: `agent1-ms-project`;
- команда: `C:\Users\Qbal\AppData\Local\Programs\Python\Python312\python.exe`;
- аргумент: `D:\Claude\ClaudeVS\agent1\tools\mcp_server.py`.

Фрагмент проектной конфигурации:

```toml
[mcp_servers.agent1-ms-project]
command = 'C:\Users\Qbal\AppData\Local\Programs\Python\Python312\python.exe'
args = ['D:\Claude\ClaudeVS\agent1\tools\mcp_server.py']
startup_timeout_sec = 10
tool_timeout_sec = 1800
enabled = true
```

После изменения конфигурации Codex нужно перезапустить. Для первой проверки
попросите агента вызвать `schedule_summary` для существующего IR, затем
`schedule_validate_ir`. Запись MPP следует проверять отдельно на новом имени
файла, после чего запускать `mpp_validate` со ссылкой на исходный IR.

## Ручная проверка сервера

```powershell
python tools\mcp_server.py
```

Сервер ожидает JSON-RPC по одной строке и обычно запускается самим Codex, а не
используется человеком напрямую. В stdout он пишет только MCP-сообщения;
поэтому обычное отсутствие приглашения командной строки является нормальным.
