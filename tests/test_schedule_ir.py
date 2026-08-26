# -*- coding: utf-8 -*-
"""Schedule IR v1.0: единый контракт для Excel, MCP и Microsoft Project."""
from __future__ import annotations

import json
from pathlib import Path

import build_grp
from schedule_ir import ScheduleProject, schedule_from_grp, validate_schedule_ir

ROOT = Path(__file__).resolve().parents[1]


def build_ir() -> ScheduleProject:
    project = json.loads((ROOT / "tests" / "etalon_project.json").read_text(encoding="utf-8"))
    b = build_grp.Build(project)
    b.load_skeleton()
    b.check_stages()
    b.repair_defects()
    b.inherit_summary_links()
    b.apply_site_conditions()
    b.configure_corpuses()
    b.configure_zero_cycle()
    b.configure_parking()
    b.apply_finishing_scope()
    b.apply_standards()
    for corpus in project["корпуса"]:
        b.rebuild_monolith(corpus)
    nodes = b.schedule()
    b.thermal(nodes)
    b.wire_zos()
    nodes = b.schedule()
    rows = b.finalize(nodes)
    return schedule_from_grp(project, rows)


def test_etalon_converts_to_valid_ir():
    schedule = build_ir()
    assert schedule.schema_version == "1.0"
    assert len(schedule.tasks) == 1677
    assert validate_schedule_ir(schedule) == []
    assert {link.type for task in schedule.tasks for link in task.predecessors} <= {
        "FS", "SS", "FF", "SF",
    }


def test_ir_preserves_wbs_phase_and_source_identity():
    schedule = build_ir()
    heat = next(task for task in schedule.tasks if task.name == "К1. Пуск тепла корпус")
    by_id = {task.task_id: task for task in schedule.tasks}
    predecessor_names = {by_id[link.predecessor_id].name for link in heat.predecessors}
    assert heat.parent_id is not None
    assert heat.source_key is not None
    assert heat.phase == "B"
    assert "К1. По договору Отопление (контур для пуска тепла)" in predecessor_names
    assert "К1. По договору Отопление (все система полностью)" not in predecessor_names


def test_ir_json_round_trip_is_lossless():
    schedule = build_ir()
    payload = schedule.to_dict()
    assert "Название задачи" not in payload["tasks"][0]
    assert "name" in payload["tasks"][0]
    restored = ScheduleProject.from_json(schedule.to_json())
    assert restored.to_dict() == schedule.to_dict()
    assert validate_schedule_ir(restored) == []


def test_ir_validator_rejects_self_link():
    schedule = build_ir()
    task = next(task for task in schedule.tasks if task.predecessors)
    task.predecessors[0] = type(task.predecessors[0])(
        predecessor_id=task.task_id,
        type="FS",
        lag_days=0,
    )
    assert "IR-LINK-SELF" in {issue.code for issue in validate_schedule_ir(schedule)}
