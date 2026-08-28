# -*- coding: utf-8 -*-
"""Быстрые локальные hooks для MS Project Agent.

Скрипт читает одно событие Codex в JSON из stdin и пишет только JSON в stdout.
Он не обращается к сети и не записывает пользовательские данные на диск.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
CORPORATE_TEMPLATE = (ROOT / "data" / "Шаблон ГРП.mpp").resolve()
EXPECTED = {
    "AGENTS.md": "Версия файла: 7.4",
    "instructions/typGRP.md": "v4.3",
    "instructions/bindings.md": "v3.2",
    "instructions/standards.md": "Версия 3.3",
}


def _output_context(message: str) -> dict[str, Any]:
    return {"hookSpecificOutput": {"additionalContext": message}}


def _deny(reason: str) -> dict[str, Any]:
    return {
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": "deny",
            "permissionDecisionReason": reason,
        }
    }


def _resolve_project_path(value: Any) -> Path | None:
    if not isinstance(value, str) or not value.strip():
        return None
    path = Path(value)
    if not path.is_absolute():
        path = ROOT / path
    try:
        return path.resolve()
    except OSError:
        return None


def _session_start() -> dict[str, Any]:
    problems: list[str] = []
    for relative, marker in EXPECTED.items():
        path = ROOT / relative
        if not path.is_file():
            problems.append(f"нет {relative}")
            continue
        try:
            content = path.read_text(encoding="utf-8")
        except OSError as exc:
            problems.append(f"не читается {relative}: {exc}")
            continue
        if marker not in content:
            problems.append(f"версия {relative} не соответствует ожидаемой ({marker})")
    templates = list((ROOT / "data").glob("*.mpp")) if (ROOT / "data").is_dir() else []
    if templates != [CORPORATE_TEMPLATE]:
        problems.append("в data должен быть ровно один корпоративный файл Шаблон ГРП.mpp")

    if problems:
        return _output_context(
            "MS Project Agent не готов: " + "; ".join(problems)
            + ". Не начинай расчёт или экспорт до устранения проблемы."
        )
    return _output_context(
        "MS Project Agent: комплект версий 7.4/4.3/3.2/3.3 найден. "
        "Перед каждым расчётом полностью загрузи typGRP.md, bindings.md и standards.md. "
        "Для MPP используй только data/Шаблон ГРП.mpp, записывай новый файл и после экспорта вызывай mpp_validate."
    )


def _pre_tool_use(event: dict[str, Any]) -> dict[str, Any]:
    name = str(event.get("tool_name") or event.get("toolName") or "")
    tool_input = event.get("tool_input") or event.get("toolInput") or {}
    if not isinstance(tool_input, dict):
        tool_input = {}

    if name.endswith("mpp_export"):
        output = _resolve_project_path(tool_input.get("mpp_path"))
        if output == CORPORATE_TEMPLATE:
            return _deny("Корпоративный Шаблон ГРП.mpp является только источником и не может быть выходным файлом.")
        if output is not None and output.exists():
            return _deny("Экспорт MPP разрешён только в новый файл; указанный путь уже существует.")
        template = _resolve_project_path(tool_input.get("template_path"))
        if template is not None and template != CORPORATE_TEMPLATE:
            return _deny("Для экспорта MPP разрешён только data/Шаблон ГРП.mpp.")

    command = str(tool_input.get("command") or tool_input.get("cmd") or "")
    normalized = command.replace("/", "\\").lower()
    mentions_template = "шаблон грп.mpp" in normalized
    mutating = re.search(
        r"\b(remove-item|move-item|rename-item|set-content|add-content|clear-content|"
        r"out-file|del|erase|move|ren|rm|mv)\b|(?:>>?|2>)",
        normalized,
    )
    if mentions_template and mutating:
        return _deny("Команда может изменить или переместить корпоративный Шаблон ГРП.mpp; операция запрещена.")

    patch_text = str(tool_input.get("patch") or tool_input.get("input") or "")
    if "шаблон грп.mpp" in patch_text.replace("/", "\\").lower():
        return _deny("Корпоративный Шаблон ГРП.mpp нельзя изменять через patch-инструмент.")
    return {}


def _post_tool_use(event: dict[str, Any]) -> dict[str, Any]:
    name = str(event.get("tool_name") or event.get("toolName") or "")
    if name.endswith("mpp_export"):
        return _output_context(
            "Экспорт MPP завершён. До передачи результата обязательно вызови mpp_validate с тем же Schedule IR "
            "и новым report_path; MPP с ошибками не помечай как готовый."
        )
    return {}


def handle(event: dict[str, Any]) -> dict[str, Any]:
    event_name = str(event.get("hook_event_name") or event.get("hookEventName") or "")
    if event_name == "SessionStart":
        return _session_start()
    if event_name == "PreToolUse":
        return _pre_tool_use(event)
    if event_name == "PostToolUse":
        return _post_tool_use(event)
    return {}


def main() -> int:
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        event = json.load(sys.stdin)
        if not isinstance(event, dict):
            raise ValueError("hook input must be a JSON object")
        result = handle(event)
    except Exception as exc:  # Hook must fail visibly, without a traceback in stdout.
        result = _output_context(f"Ошибка локального hook MS Project Agent: {exc}")
    json.dump(result, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
