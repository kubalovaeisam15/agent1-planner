# -*- coding: utf-8 -*-
"""Регрессии решений владельца DEC-34…37 от 29.08.2026."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

import build_grp
from schedule_ir import schedule_from_grp, validate_schedule_ir
from shared_sections import PROJECT_SHARED_SECTIONS


ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def corrected_schedule():
    project = json.loads(
        (ROOT / "tests" / "regression_two_corpuses.json").read_text(encoding="utf-8"))
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
    rows = b.finalize(b.schedule())
    schedule = schedule_from_grp(project, rows)
    assert validate_schedule_ir(schedule) == []
    return schedule


def test_dec34_underground_task_names_follow_tep(corrected_schedule):
    names = [t.name for t in corrected_schedule.tasks
             if "По договору Монолитные конструкции ниже отм 0,000" in t.name]
    assert names
    assert all("2 подземных этажа" in name for name in names)
    assert all("1 подземный этаж" not in name for name in names)


def test_dec35_pile_types_keep_individual_durations(corrected_schedule):
    piles = {t.name: t.duration_days for t in corrected_schedule.tasks
             if "По договору Свайное основание" in t.name}
    assert next(v for k, v in piles.items() if "БНС" in k) == 35
    assert next(v for k, v in piles.items() if "Забивные" in k) == 20


def test_dec36_first_floor_waits_only_own_below_zero(corrected_schedule):
    by_id = {t.task_id: t for t in corrected_schedule.tasks}
    for corpus in ("К1", "К2"):
        first = next(t for t in corrected_schedule.tasks
                     if t.name == f"{corpus}. 1 этаж Монолит")
        assert len(first.predecessors) == 1
        predecessor = by_id[first.predecessors[0].predecessor_id]
        assert first.predecessors[0].type == "FS"
        assert first.predecessors[0].lag_days == 0
        assert predecessor.name.startswith(f"{corpus}. ")
        assert "Монолитные конструкции ниже отм 0,000" in predecessor.name


def test_dec37_finishing_preserves_physical_links_and_snet(corrected_schedule):
    by_id = {t.task_id: t for t in corrected_schedule.tasks}
    leaves = [t for t in corrected_schedule.tasks
              if t.name.endswith("По договору Вестибюль")]
    assert {t.name[:2] for t in leaves} == {"К1", "К2"}
    for task in leaves:
        predecessor_names = [by_id[p.predecessor_id].name for p in task.predecessors]
        assert any("Кладка перегородок" in name for name in predecessor_names)
        assert any("Заключение договора Отделка" in name for name in predecessor_names)
        assert not any("Монолит" in name for name in predecessor_names)
        assert task.constraint_type == "start_no_earlier_than"
        assert task.constraint_date is not None


def test_dec38_dd_constraint_equals_mz_project_start(corrected_schedule):
    task = next(t for t in corrected_schedule.tasks
                if t.name == "Проведение ДД и выкуп 100% акций")
    assert task.constraint_type == "start_no_earlier_than"
    assert task.constraint_date == corrected_schedule.project_start
    assert task.start == corrected_schedule.project_start


def test_dec40_shared_closeout_uses_both_corpuses(corrected_schedule):
    tasks = corrected_schedule.tasks
    for name in PROJECT_SHARED_SECTIONS:
        matches = [task for task in tasks if task.name == name]
        assert len(matches) == 1
        assert matches[0].outline_level == 1
    bti = [task for task in tasks if task.name == "Готовность к обмерам БТИ"]
    assert len(bti) == 1
    by_id = {task.task_id: task for task in tasks}
    predecessors = [by_id[link.predecessor_id].name for link in bti[0].predecessors]
    assert any(name.startswith("К1. ") for name in predecessors)
    assert any(name.startswith("К2. ") for name in predecessors)
