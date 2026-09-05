"""V03 / DEC-25: дерево IR и свёртка дат, без изменения входа."""
import copy
from datetime import date, timedelta

import pytest
from schedule_ir import ScheduleLink, ScheduleProject, ScheduleTask, validate_schedule_ir


def example():
    start = date(2026, 1, 1)
    def row(key, level, kind, parent, begin, end):
        return ScheduleTask(key, key, level, kind, parent_id=parent,
            start=start+timedelta(days=begin), finish=start+timedelta(days=end),
            duration_days=None if kind == "summary" else end-begin)
    return ScheduleProject("test", "test", start, [
        row("A", 1, "summary", None, 0, 10),
        row("A1", 2, "summary", "A", 0, 5),
        row("leaf1", 3, "task", "A1", 0, 5),
        row("leaf2", 2, "task", "A", 5, 10),
        row("B", 1, "summary", None, 0, 10),
        row("leaf3", 2, "task", "B", 0, 10),
    ])


def codes(ir):
    return {i.code for i in validate_schedule_ir(ir)}


def test_valid_nested_tree_and_validator_does_not_mutate():
    ir = example()
    before = copy.deepcopy(ir)
    assert validate_schedule_ir(ir, require_all_sections=False) == []
    assert ir == before


@pytest.mark.parametrize("index,parent", [(2,"A"), (5,"A"), (5,"leaf2"), (4,"A"), (2,None), (2,"missing"), (2,"B"), (2,"leaf1")])
def test_parent_must_be_immediate_active_ancestor(index, parent):
    ir = example()
    ir.tasks[index].parent_id = parent
    assert "IR-PARENT" in codes(ir)


@pytest.mark.parametrize("kind", ["task", "milestone"])
def test_parent_must_be_summary(kind):
    ir = example()
    ir.tasks[1].task_type = kind
    ir.tasks[1].duration_days = 5 if kind == "task" else 0
    assert "IR-PARENT-TYPE" in codes(ir)


def test_summary_requires_children():
    ir = example()
    ir.tasks[2].task_type = "summary"
    ir.tasks[2].duration_days = None
    assert "IR-SUMMARY-EMPTY" in codes(ir)


@pytest.mark.parametrize("field,delta", [("start",-1), ("start",1), ("finish",-1), ("finish",1)])
def test_summary_dates_are_exact_rollup(field, delta):
    ir = example()
    setattr(ir.tasks[0], field, getattr(ir.tasks[0], field)+timedelta(days=delta))
    assert "IR-SUMMARY-DATES" in codes(ir)


def test_nested_summary_dates_are_checked_even_when_outer_dates_match():
    ir = example()
    ir.tasks[1].finish += timedelta(days=1)
    issues = validate_schedule_ir(ir)
    assert any(i.code == "IR-SUMMARY-DATES" and i.task_id == "A1" for i in issues)


def test_missing_child_date_is_reported_without_crash():
    ir = example()
    ir.tasks[2].start = None
    assert "IR-DATES-REQUIRED" in codes(ir)


@pytest.mark.parametrize("level", [None, "2", True, 0, -1, 4])
def test_bad_levels_are_reported_without_crash(level):
    ir = example()
    ir.tasks[1].outline_level = level
    assert "IR-WBS-LEVEL" in codes(ir)


def test_duplicate_ids_are_reported_without_crash():
    ir = example()
    ir.tasks[4].task_id = "A"
    assert "IR-DUPLICATE-ID" in codes(ir)


def test_single_milestone_child_and_summary_links_are_valid():
    ir = example()
    ir.tasks = ir.tasks[:3]
    for task in ir.tasks:
        task.finish = task.start
    ir.tasks[2].task_type = "milestone"
    ir.tasks[2].duration_days = 0
    # НН +0 допускает равные даты; связи сводок не разворачиваются.
    ir.tasks[1].predecessors = [ScheduleLink("A", "SS", 0)]
    ir.tasks[2].predecessors = [ScheduleLink("A1", "SS", 0)]
    before = copy.deepcopy(ir)
    assert validate_schedule_ir(ir, require_all_sections=False) == []
    assert ir == before


@pytest.mark.parametrize("mutation", ["parent", "dates"])
def test_mcp_rejects_invalid_ir_before_export(tmp_path, monkeypatch, mutation):
    import mcp_server
    from ir_test_fixtures import complete_sections
    ir = complete_sections(example())
    if mutation == "parent":
        ir.tasks[5].parent_id = "A"
    else:
        ir.tasks[0].start -= timedelta(days=1)
    source = tmp_path / "invalid.ir.json"
    target = tmp_path / "never-created.mpp"
    source.write_text(ir.to_json(), encoding="utf-8")
    result = mcp_server._schedule_validate_ir({"ir_path": str(source)})
    assert result["valid"] is False
    found = {i["code"] for i in result["issues"]}
    assert ("IR-PARENT" if mutation == "parent" else "IR-SUMMARY-DATES") in found
    assert "IR-SHARED-SECTION" not in found
    calls = []
    monkeypatch.setattr(mcp_server, "export_mpp", lambda *args, **kwargs: calls.append(args))
    with pytest.raises(mcp_server.ToolError, match="Экспорт остановлен"):
        mcp_server._mpp_export({"ir_path": str(source), "mpp_path": str(target)})
    assert calls == []
    assert not target.exists()
