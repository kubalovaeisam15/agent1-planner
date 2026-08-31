# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import subprocess
import sys
from datetime import date
from pathlib import Path
from uuid import uuid4

import mcp_server
from schedule_ir import ScheduleProject, ScheduleTask

ROOT = Path(__file__).resolve().parents[1]


def sample_schedule() -> ScheduleProject:
    return ScheduleProject(
        schedule_id="mcp-test",
        name="Тест MCP",
        project_start=date(2026, 1, 1),
        tasks=[
            ScheduleTask("1", "Раздел", 1, "summary",
                         start=date(2026, 1, 1), finish=date(2026, 1, 6)),
            ScheduleTask("2", "Работа", 2, "task", parent_id="1", duration_days=5,
                         start=date(2026, 1, 1), finish=date(2026, 1, 6), critical=True),
        ],
    )


def test_initialize_and_tool_catalog_are_mcp_json_rpc():
    initialized = mcp_server.handle_request({
        "jsonrpc": "2.0", "id": 1, "method": "initialize",
        "params": {"protocolVersion": "2025-06-18"},
    })
    assert initialized is not None
    assert initialized["result"]["serverInfo"]["name"] == "agent1-ms-project"
    assert initialized["result"]["serverInfo"]["version"] == "0.5.0"
    assert initialized["result"]["protocolVersion"] == "2025-06-18"

    listed = mcp_server.handle_request({
        "jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {},
    })
    assert listed is not None
    names = {tool["name"] for tool in listed["result"]["tools"]}
    assert names == {
        "context_preflight",
        "schedule_summary", "schedule_validate_ir", "schedule_build",
        "mpp_export", "mpp_validate",
    }
    export_tool = next(tool for tool in listed["result"]["tools"]
                       if tool["name"] == "mpp_export")
    assert "template_path" in export_tool["inputSchema"]["properties"]
    assert "template_path" not in export_tool["inputSchema"]["required"]


def test_context_preflight_verifies_compiled_context():
    result = mcp_server.call_tool("context_preflight", {})
    assert result["isError"] is False
    content = result["structuredContent"]
    assert content["ready"] is True
    assert content["agent_policy_version"] == "7.5"
    assert content["issue_count"] == 0
    assert len(content["verified"]) == 17
    assert all(item["verified"] for item in content["verified"])
    assert content["template_path"].endswith("Шаблон ГРП.mpp")


def test_context_preflight_detects_hash_mismatch(tmp_path, monkeypatch):
    manifest = json.loads(mcp_server.CONTEXT_MANIFEST.read_text(encoding="utf-8"))
    manifest["files"][0]["sha256"] = "0" * 64
    changed = tmp_path / "context-manifest.json"
    changed.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setattr(mcp_server, "CONTEXT_MANIFEST", changed)

    result = mcp_server.call_tool("context_preflight", {})["structuredContent"]
    assert result["ready"] is False
    assert result["issue_count"] == 1
    assert result["issues"][0]["code"] == "HASH_MISMATCH"

    token = uuid4().hex
    xlsx = ROOT / "out" / f"blocked-{token}.xlsx"
    ir = ROOT / "out" / f"blocked-{token}.json"
    blocked = mcp_server.call_tool("schedule_build", {
        "spec_path": "tests/etalon_project.json",
        "xlsx_path": str(xlsx.relative_to(ROOT)),
        "ir_path": str(ir.relative_to(ROOT)),
    })
    assert blocked["isError"] is True
    assert "context_preflight" in blocked["structuredContent"]["error"]
    assert not xlsx.exists()
    assert not ir.exists()


def test_notification_has_no_response():
    assert mcp_server.handle_request({
        "jsonrpc": "2.0", "method": "notifications/initialized",
    }) is None


def test_summary_and_validation_return_structured_content():
    token = uuid4().hex
    ir = ROOT / "out" / f"mcp-test-{token}.ir.json"
    ir.parent.mkdir(exist_ok=True)
    try:
        ir.write_text(sample_schedule().to_json(), encoding="utf-8")
        relative = str(ir.relative_to(ROOT))
        summary = mcp_server.call_tool("schedule_summary", {"ir_path": relative})
        assert summary["isError"] is False
        assert summary["structuredContent"]["summary"]["task_count"] == 2
        assert summary["structuredContent"]["summary"]["critical_leaf_count"] == 1

        validation = mcp_server.call_tool(
            "schedule_validate_ir", {"ir_path": relative})
        assert validation["isError"] is False
        assert validation["structuredContent"]["valid"] is True
    finally:
        ir.unlink(missing_ok=True)


def test_paths_outside_root_and_existing_outputs_are_rejected():
    denied = mcp_server.call_tool(
        "schedule_summary", {"ir_path": str(ROOT.parent / "outside.json")})
    assert denied["isError"] is True
    assert "вне корня проекта" in denied["structuredContent"]["error"]

    token = uuid4().hex
    existing = ROOT / "out" / f"mcp-existing-{token}.xlsx"
    try:
        existing.write_bytes(b"do not overwrite")
        blocked = mcp_server.call_tool("schedule_build", {
            "spec_path": "tests/etalon_project.json",
            "xlsx_path": str(existing.relative_to(ROOT)),
            "ir_path": f"out/mcp-existing-{token}.json",
        })
        assert blocked["isError"] is True
        assert existing.read_bytes() == b"do not overwrite"
    finally:
        existing.unlink(missing_ok=True)


def test_stdio_transport_emits_only_json_lines():
    messages = [
        {"jsonrpc": "2.0", "id": 1, "method": "initialize",
         "params": {"protocolVersion": "2025-06-18"}},
        {"jsonrpc": "2.0", "method": "notifications/initialized"},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}},
        {"jsonrpc": "2.0", "id": 3, "method": "ping", "params": {}},
    ]
    payload = "".join(json.dumps(item) + "\n" for item in messages)
    completed = subprocess.run(
        [sys.executable, str(ROOT / "tools" / "mcp_server.py")],
        input=payload, capture_output=True, text=True, encoding="utf-8",
        cwd=ROOT, timeout=10, check=False,
    )
    assert completed.returncode == 0
    assert completed.stderr == ""
    responses = [json.loads(line) for line in completed.stdout.splitlines()]
    assert [response["id"] for response in responses] == [1, 2, 3]
    assert all(response["jsonrpc"] == "2.0" for response in responses)
    assert "context_preflight" in responses[0]["result"]["instructions"]
