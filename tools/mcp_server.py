# -*- coding: utf-8 -*-
"""Локальный MCP STDIO-сервер для агента-планировщика.

Сервер намеренно не зависит от внешнего MCP SDK: сообщения JSON-RPC передаются
по одной строке через stdin/stdout. В stdout нельзя писать ничего кроме
протокола; диагностические сообщения направляются в stderr.
"""
from __future__ import annotations

import hashlib
import io
import json
import subprocess
import sys
from collections import Counter
from contextlib import redirect_stderr, redirect_stdout
from dataclasses import asdict
from pathlib import Path
from typing import Any, Callable

import build_grp
from mpp_validator import (
    compare_with_ir,
    read_mpp,
    report_dict,
    validate_snapshot,
)
from mspdi_adapter import export_mpp
from schedule_ir import ScheduleProject, validate_schedule_ir

ROOT = Path(__file__).resolve().parents[1]
SERVER_NAME = "agent1-ms-project"
SERVER_VERSION = "0.5.0"
DEFAULT_MPP_TEMPLATE = ROOT / "data" / "Шаблон ГРП.mpp"
CONTEXT_MANIFEST = ROOT / "instructions" / "context-manifest.json"
EXPECTED_CONTEXT_SCHEMA = 1
EXPECTED_CONTEXT_PROFILE = "agent1-grp-2026-08-25"
EXPECTED_AGENT_POLICY_VERSION = "7.5"
DEFAULT_PROTOCOL_VERSION = "2025-06-18"
MAX_INLINE_ISSUES = 10


class ToolError(Exception):
    """Ожидаемая ошибка параметров или выполнения MCP-инструмента."""


def _schema(properties: dict[str, Any], required: list[str]) -> dict[str, Any]:
    return {
        "type": "object",
        "properties": properties,
        "required": required,
        "additionalProperties": False,
    }


PATH = {"type": "string", "minLength": 1, "description": "Путь внутри корня проекта"}
TIMEOUT = {
    "type": "integer", "minimum": 30, "maximum": 1800, "default": 300,
    "description": "Таймаут Microsoft Project COM в секундах",
}

TOOLS = [
    {
        "name": "context_preflight",
        "description": "Проверить версии и хэши нормативов, расчётного ядра и корпоративного MPP-шаблона.",
        "inputSchema": _schema({}, []),
        "annotations": {"readOnlyHint": True, "destructiveHint": False,
                        "idempotentHint": True, "openWorldHint": False},
    },
    {
        "name": "schedule_summary",
        "description": "Прочитать Schedule IR JSON и вернуть компактную сводку без изменений файлов.",
        "inputSchema": _schema({"ir_path": PATH}, ["ir_path"]),
        "annotations": {"readOnlyHint": True, "destructiveHint": False,
                        "idempotentHint": True, "openWorldHint": False},
    },
    {
        "name": "schedule_validate_ir",
        "description": "Проверить контракт, WBS, даты и связи Schedule IR без пересчёта и записи.",
        "inputSchema": _schema({"ir_path": PATH}, ["ir_path"]),
        "annotations": {"readOnlyHint": True, "destructiveHint": False,
                        "idempotentHint": True, "openWorldHint": False},
    },
    {
        "name": "schedule_build",
        "description": (
            "Построить новый Excel ГРП и Schedule IR из JSON ТЭП. "
            "Существующие выходные файлы не перезаписываются."
        ),
        "inputSchema": _schema({
            "spec_path": PATH,
            "xlsx_path": PATH,
            "ir_path": PATH,
        }, ["spec_path", "xlsx_path", "ir_path"]),
        "annotations": {"readOnlyHint": False, "destructiveHint": False,
                        "idempotentHint": False, "openWorldHint": False},
    },
    {
        "name": "mpp_export",
        "description": (
            "Создать новый MPP из проверенного Schedule IR через Microsoft Project "
            "на основе корпоративного шаблона. Существующий MPP не перезаписывается."
        ),
        "inputSchema": _schema({
            "ir_path": PATH,
            "mpp_path": PATH,
            "template_path": PATH,
            "timeout_seconds": TIMEOUT,
        }, ["ir_path", "mpp_path"]),
        "annotations": {"readOnlyHint": False, "destructiveHint": False,
                        "idempotentHint": False, "openWorldHint": False},
    },
    {
        "name": "mpp_validate",
        "description": (
            "Прочитать MPP без сохранения, проверить сеть и при наличии IR выполнить точную сверку. "
            "Полный JSON-отчёт записывается только если указан новый report_path."
        ),
        "inputSchema": _schema({
            "mpp_path": PATH,
            "ir_path": PATH,
            "report_path": PATH,
            "timeout_seconds": TIMEOUT,
        }, ["mpp_path"]),
        "annotations": {"readOnlyHint": False, "destructiveHint": False,
                        "idempotentHint": True, "openWorldHint": False},
    },
]


def _path(value: Any, *, suffix: str | None = None, exists: bool = True) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ToolError("Путь должен быть непустой строкой")
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = ROOT / candidate
    candidate = candidate.resolve()
    try:
        candidate.relative_to(ROOT)
    except ValueError as exc:
        raise ToolError(f"Путь вне корня проекта запрещён: {candidate}") from exc
    if suffix and candidate.suffix.lower() != suffix:
        raise ToolError(f"Ожидается файл {suffix}: {candidate}")
    if exists and not candidate.is_file():
        raise ToolError(f"Файл не найден: {candidate}")
    if not exists and candidate.exists():
        raise ToolError(f"Выходной файл уже существует: {candidate}")
    return candidate


def _timeout(arguments: dict[str, Any]) -> int:
    value = arguments.get("timeout_seconds", 300)
    if isinstance(value, bool) or not isinstance(value, int) or not 30 <= value <= 1800:
        raise ToolError("timeout_seconds должен быть целым числом от 30 до 1800")
    return value


def _load_ir(path_value: Any) -> tuple[Path, ScheduleProject]:
    path = _path(path_value, suffix=".json")
    try:
        schedule = ScheduleProject.from_json(path.read_text(encoding="utf-8"))
    except (KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
        raise ToolError(f"Некорректный Schedule IR: {exc}") from exc
    return path, schedule


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _context_preflight(arguments: dict[str, Any]) -> dict[str, Any]:
    try:
        manifest = json.loads(CONTEXT_MANIFEST.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ToolError(f"Не читается манифест контекста: {exc}") from exc

    issues: list[dict[str, str]] = []
    verified: list[dict[str, Any]] = []
    expected_header = {
        "schema_version": EXPECTED_CONTEXT_SCHEMA,
        "profile": EXPECTED_CONTEXT_PROFILE,
        "agent_policy_version": EXPECTED_AGENT_POLICY_VERSION,
    }
    for key, expected in expected_header.items():
        if manifest.get(key) != expected:
            issues.append({
                "code": "MANIFEST_HEADER", "path": "instructions/context-manifest.json",
                "message": f"{key}={manifest.get(key)!r}, ожидалось {expected!r}",
            })
    entries = manifest.get("files")
    if not isinstance(entries, list):
        raise ToolError("Некорректный context-manifest.json: files должен быть массивом")

    for entry in entries:
        if not isinstance(entry, dict):
            issues.append({"code": "MANIFEST_ENTRY", "message": "Некорректная запись files"})
            continue
        relative = entry.get("path")
        try:
            path = _path(relative)
        except ToolError as exc:
            issues.append({"code": "FILE_MISSING", "path": str(relative), "message": str(exc)})
            continue
        actual_hash = _sha256(path)
        expected_hash = str(entry.get("sha256", "")).upper()
        entry_ok = actual_hash == expected_hash
        if actual_hash != expected_hash:
            issues.append({
                "code": "HASH_MISMATCH", "path": str(relative),
                "message": f"SHA256 {actual_hash}, ожидался {expected_hash}",
            })
        marker = entry.get("version_marker")
        if marker:
            try:
                text = path.read_text(encoding="utf-8")
            except (OSError, UnicodeError) as exc:
                issues.append({"code": "READ_ERROR", "path": str(relative), "message": str(exc)})
            else:
                if str(marker) not in text:
                    entry_ok = False
                    issues.append({
                        "code": "VERSION_MISMATCH", "path": str(relative),
                        "message": f"не найден маркер {marker}",
                    })
        verified.append({
            "path": str(relative), "version": entry.get("version"),
            "verified": entry_ok,
        })

    templates = sorted((ROOT / "data").glob("*.mpp"))
    if templates != [DEFAULT_MPP_TEMPLATE]:
        issues.append({
            "code": "TEMPLATE_SET",
            "message": "В data должен быть ровно один корпоративный Шаблон ГРП.mpp",
        })
    return {
        "ready": not issues,
        "profile": manifest.get("profile"),
        "agent_policy_version": manifest.get("agent_policy_version"),
        "verified": verified,
        "issue_count": len(issues),
        "issues": issues[:MAX_INLINE_ISSUES],
        "issues_truncated": len(issues) > MAX_INLINE_ISSUES,
        "template_path": str(DEFAULT_MPP_TEMPLATE),
    }


def _require_context_ready() -> None:
    preflight = _context_preflight({})
    if not preflight["ready"]:
        first = preflight["issues"][0]["message"] if preflight["issues"] else "неизвестная ошибка"
        raise ToolError(
            f"Операция остановлена: context_preflight обнаружил "
            f"{preflight['issue_count']} ошибок. Первая: {first}"
        )


def _summary(schedule: ScheduleProject) -> dict[str, Any]:
    types = Counter(task.task_type for task in schedule.tasks)
    links = sum(len(task.predecessors) for task in schedule.tasks)
    dated = [task for task in schedule.tasks if task.start and task.finish]
    return {
        "schedule_id": schedule.schedule_id,
        "name": schedule.name,
        "schema_version": schedule.schema_version,
        "project_start": schedule.project_start.isoformat() if schedule.project_start else None,
        "project_finish": max((task.finish for task in dated), default=None).isoformat()
        if dated else None,
        "task_count": len(schedule.tasks),
        "task_types": dict(sorted(types.items())),
        "link_count": links,
        "critical_leaf_count": sum(
            task.critical and task.task_type != "summary" for task in schedule.tasks),
    }


def _schedule_summary(arguments: dict[str, Any]) -> dict[str, Any]:
    path, schedule = _load_ir(arguments.get("ir_path"))
    return {"ir_path": str(path), "summary": _summary(schedule)}


def _schedule_validate_ir(arguments: dict[str, Any]) -> dict[str, Any]:
    path, schedule = _load_ir(arguments.get("ir_path"))
    issues = [asdict(issue) for issue in validate_schedule_ir(schedule)]
    return {
        "ir_path": str(path),
        "valid": not issues,
        "issue_count": len(issues),
        "issues": issues[:MAX_INLINE_ISSUES],
        "issues_truncated": len(issues) > MAX_INLINE_ISSUES,
        "summary": _summary(schedule),
    }


def _schedule_build(arguments: dict[str, Any]) -> dict[str, Any]:
    _require_context_ready()
    spec = _path(arguments.get("spec_path"), suffix=".json")
    xlsx = _path(arguments.get("xlsx_path"), suffix=".xlsx", exists=False)
    ir = _path(arguments.get("ir_path"), suffix=".json", exists=False)
    xlsx.parent.mkdir(parents=True, exist_ok=True)
    ir.parent.mkdir(parents=True, exist_ok=True)
    captured = io.StringIO()
    with redirect_stdout(captured), redirect_stderr(captured):
        code = build_grp.main([str(spec), str(xlsx), "--ir", str(ir)])
    if code != 0 or not xlsx.is_file() or not ir.is_file():
        raise ToolError(f"Генератор завершился с кодом {code}: {captured.getvalue().strip()}")
    schedule = ScheduleProject.from_json(ir.read_text(encoding="utf-8"))
    issues = validate_schedule_ir(schedule)
    return {
        "xlsx_path": str(xlsx),
        "ir_path": str(ir),
        "ir_valid": not issues,
        "ir_issue_count": len(issues),
        "ir_issues": [asdict(issue) for issue in issues[:MAX_INLINE_ISSUES]],
        "ir_issues_truncated": len(issues) > MAX_INLINE_ISSUES,
        "summary": _summary(schedule),
    }


def _mpp_export(arguments: dict[str, Any]) -> dict[str, Any]:
    _require_context_ready()
    ir, schedule = _load_ir(arguments.get("ir_path"))
    issues = validate_schedule_ir(schedule)
    if issues:
        raise ToolError(
            f"Экспорт остановлен: Schedule IR содержит {len(issues)} ошибок; "
            "сначала вызовите schedule_validate_ir"
        )
    mpp = _path(arguments.get("mpp_path"), suffix=".mpp", exists=False)
    template_value = arguments.get("template_path")
    template = (_path(template_value, suffix=".mpp") if template_value is not None
                else DEFAULT_MPP_TEMPLATE)
    report = export_mpp(ir, mpp, template_path=template,
                        timeout_seconds=_timeout(arguments))
    return {"mpp_path": str(mpp), "ir_path": str(ir),
            "template_path": str(template), "report": report}


def _mpp_validate(arguments: dict[str, Any]) -> dict[str, Any]:
    mpp = _path(arguments.get("mpp_path"), suffix=".mpp")
    snapshot = read_mpp(mpp, timeout_seconds=_timeout(arguments))
    issues = validate_snapshot(snapshot)
    ir_value = arguments.get("ir_path")
    ir_path: Path | None = None
    if ir_value is not None:
        ir_path, schedule = _load_ir(ir_value)
        issues.extend(compare_with_ir(snapshot, schedule))
    report = report_dict(snapshot, issues)
    report_path_value = arguments.get("report_path")
    report_path: Path | None = None
    if report_path_value is not None:
        report_path = _path(report_path_value, suffix=".json", exists=False)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    compact = {
        "mpp_path": str(mpp),
        "ir_path": str(ir_path) if ir_path else None,
        "report_path": str(report_path) if report_path else None,
        "project": report["project"],
        "result": report["result"],
        "issues": report["issues"][:MAX_INLINE_ISSUES],
        "issues_truncated": len(report["issues"]) > MAX_INLINE_ISSUES,
    }
    return compact


HANDLERS: dict[str, Callable[[dict[str, Any]], dict[str, Any]]] = {
    "context_preflight": _context_preflight,
    "schedule_summary": _schedule_summary,
    "schedule_validate_ir": _schedule_validate_ir,
    "schedule_build": _schedule_build,
    "mpp_export": _mpp_export,
    "mpp_validate": _mpp_validate,
}


def call_tool(name: str, arguments: Any) -> dict[str, Any]:
    """Выполнить инструмент и вернуть MCP CallToolResult."""
    handler = HANDLERS.get(name)
    if handler is None:
        return _tool_result({"error": f"Неизвестный инструмент: {name}"}, is_error=True)
    if not isinstance(arguments, dict):
        return _tool_result({"error": "arguments должен быть JSON-объектом"}, is_error=True)
    try:
        result = handler(arguments)
    except (ToolError, OSError, RuntimeError, ValueError, TypeError, KeyError,
            UnicodeError, subprocess.SubprocessError) as exc:
        return _tool_result({"error": str(exc), "tool": name}, is_error=True)
    return _tool_result(result)


def _tool_result(value: dict[str, Any], *, is_error: bool = False) -> dict[str, Any]:
    return {
        "content": [{"type": "text", "text": json.dumps(
            value, ensure_ascii=False, indent=2, default=str)}],
        "structuredContent": value,
        "isError": is_error,
    }


def handle_request(message: dict[str, Any]) -> dict[str, Any] | None:
    """Обработать одно JSON-RPC сообщение. Уведомления ответа не требуют."""
    request_id = message.get("id")
    method = message.get("method")
    if request_id is None:
        return None
    try:
        if method == "initialize":
            requested = message.get("params", {}).get("protocolVersion")
            result = {
                "protocolVersion": requested or DEFAULT_PROTOCOL_VERSION,
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
                "instructions": (
                    "Файлы только внутри agent1. Начните с context_preflight. "
                    "schedule_build и mpp_export сами проверяют IR; MPP всегда проверяйте mpp_validate."
                ),
            }
        elif method == "ping":
            result = {}
        elif method == "tools/list":
            result = {"tools": TOOLS}
        elif method == "tools/call":
            params = message.get("params") or {}
            result = call_tool(params.get("name", ""), params.get("arguments", {}))
        else:
            return _rpc_error(request_id, -32601, f"Метод не поддерживается: {method}")
        return {"jsonrpc": "2.0", "id": request_id, "result": result}
    except Exception as exc:  # последний рубеж: протокол не должен падать от одного запроса
        return _rpc_error(request_id, -32603, f"Внутренняя ошибка: {exc}")


def _rpc_error(request_id: Any, code: int, message: str) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id,
            "error": {"code": code, "message": message}}


def serve() -> int:
    """Запустить построчный JSON-RPC цикл STDIO."""
    # MCP требует UTF-8. На Windows Python иначе может унаследовать OEM/ANSI
    # кодировку консоли и повредить русские названия задач в JSON-RPC.
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8")
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    for raw in sys.stdin:
        if not raw.strip():
            continue
        try:
            message = json.loads(raw)
            if not isinstance(message, dict):
                raise ValueError("сообщение должно быть JSON-объектом")
            response = handle_request(message)
        except (json.JSONDecodeError, ValueError) as exc:
            response = _rpc_error(None, -32700, f"Ошибка JSON: {exc}")
        if response is not None:
            sys.stdout.write(json.dumps(response, ensure_ascii=False, separators=(",", ":")) + "\n")
            sys.stdout.flush()
    return 0


if __name__ == "__main__":
    raise SystemExit(serve())
