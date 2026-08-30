# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import date
from pathlib import Path

from mpp_validator import (
    MPPLink,
    MPPSnapshot,
    MPPTask,
    compare_with_ir,
    duration_minutes,
    parse_mspdi,
    report_dict,
    validate_snapshot,
)
from schedule_ir import ScheduleLink, ScheduleProject, ScheduleTask


XML = """<?xml version="1.0" encoding="utf-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Тест</Name><StartDate>2026-01-01T00:00:00</StartDate>
  <FinishDate>2026-01-06T00:00:00</FinishDate>
  <Tasks>
    <Task><UID>0</UID><ID>0</ID><Name>Тест</Name><OutlineLevel>0</OutlineLevel><Summary>1</Summary></Task>
    <Task><UID>1</UID><ID>1</ID><Name>Раздел</Name><OutlineLevel>1</OutlineLevel>
      <Start>2026-01-01T00:00:00</Start><Finish>2026-01-06T00:00:00</Finish>
      <Duration>PT120H0M0S</Duration><Summary>1</Summary><Milestone>1</Milestone>
      <ConstraintType>0</ConstraintType>
    </Task>
    <Task><UID>2</UID><ID>2</ID><Name>Работа</Name><OutlineLevel>2</OutlineLevel>
      <Start>2026-01-01T00:00:00</Start><Finish>2026-01-06T00:00:00</Finish>
      <Duration>PT120H0M0S</Duration><Summary>0</Summary><Milestone>0</Milestone>
      <PercentComplete>0</PercentComplete><Critical>1</Critical><TotalSlack>0</TotalSlack>
      <ConstraintType>0</ConstraintType>
    </Task>
    <Task><UID>3</UID><ID>3</ID><Name>Веха</Name><OutlineLevel>2</OutlineLevel>
      <Start>2026-01-06T00:00:00</Start><Finish>2026-01-06T00:00:00</Finish>
      <Duration>PT0H0M0S</Duration><Summary>0</Summary><Milestone>1</Milestone>
      <PercentComplete>0</PercentComplete><Critical>1</Critical><TotalSlack>0</TotalSlack>
      <ConstraintType>0</ConstraintType>
      <PredecessorLink><PredecessorUID>2</PredecessorUID><Type>1</Type>
        <LinkLag>0</LinkLag><LagFormat>8</LagFormat></PredecessorLink>
    </Task>
  </Tasks>
</Project>"""


def schedule() -> ScheduleProject:
    return ScheduleProject("test", "Тест", date(2026, 1, 1), [
        ScheduleTask("1", "Раздел", 1, "summary", start=date(2026, 1, 1),
                     finish=date(2026, 1, 6)),
        ScheduleTask("2", "Работа", 2, "task", parent_id="1", duration_days=5,
                     start=date(2026, 1, 1), finish=date(2026, 1, 6), critical=True,
                     percent_complete=0),
        ScheduleTask("3", "Веха", 2, "milestone", parent_id="1", duration_days=0,
                     start=date(2026, 1, 6), finish=date(2026, 1, 6), critical=True,
                     percent_complete=0, reserve_finish=date(2026, 1, 10),
                     predecessors=[ScheduleLink("2")]),
    ])


def test_duration_parser_supports_project_iso_values():
    assert duration_minutes("P5D") == 7200
    assert duration_minutes("PT120H0M0S") == 7200


def test_parse_and_compare_round_trip(tmp_path: Path):
    path = tmp_path / "snapshot.xml"
    path.write_text(XML, encoding="utf-8")
    snapshot = parse_mspdi(path)
    assert len(snapshot.tasks) == 3
    assert compare_with_ir(snapshot, schedule()) == []
    codes = {issue.code for issue in validate_snapshot(snapshot)}
    assert "MPP-OPEN-START" in codes
    assert "MPP-OPEN-FINISH" in codes


def test_validator_detects_negative_slack(tmp_path: Path):
    path = tmp_path / "snapshot.xml"
    path.write_text(XML.replace("<TotalSlack>0</TotalSlack>",
                                "<TotalSlack>-1440</TotalSlack>", 1), encoding="utf-8")
    snapshot = parse_mspdi(path)
    assert "MPP-NEGATIVE-SLACK" in {issue.code for issue in validate_snapshot(snapshot)}


def test_ir_comparison_rejects_deadlines_used_for_dec31_reserve(tmp_path: Path):
    path = tmp_path / "snapshot-with-deadline.xml"
    payload = XML.replace(
        "<PredecessorLink>",
        "<Deadline>2026-01-10T00:00:00</Deadline><PredecessorLink>",
        1,
    )
    path.write_text(payload, encoding="utf-8")
    issues = compare_with_ir(parse_mspdi(path), schedule())
    deadline = next(issue for issue in issues if issue.code == "MPP-IR-DEADLINES")
    assert deadline.severity == "error"


def mpp_task(uid: int, name: str, level: int, *, summary: bool = False,
             milestone: bool = False, critical: bool = False,
             predecessors: list[MPPLink] | None = None) -> MPPTask:
    return MPPTask(
        uid=uid, task_id=uid, name=name, outline_level=level, summary=summary,
        milestone=milestone, start=date(2026, 1, 1), finish=date(2026, 1, 2),
        duration_minutes=0 if milestone else 1440, percent_complete=0,
        critical=critical, total_slack_minutes=0, constraint_type=0,
        constraint_date=None, deadline=None, predecessors=predecessors or [],
    )


def test_summary_links_cover_nested_leaf_network_and_critical_path():
    snapshot = MPPSnapshot("Тест", date(2026, 1, 1), date(2026, 1, 2), [
        mpp_task(1, "Старт", 1, milestone=True, critical=True),
        mpp_task(2, "Блок", 1, summary=True, critical=True,
                 predecessors=[MPPLink(1, "FS", 0)]),
        mpp_task(3, "Внутренняя работа", 2, critical=True),
        mpp_task(4, "Финиш", 1, milestone=True, critical=True,
                 predecessors=[MPPLink(2, "FS", 0)]),
    ])
    issues = validate_snapshot(snapshot)
    nested_codes = {issue.code for issue in issues if issue.task_id == 3}
    assert "MPP-OPEN-START" not in nested_codes
    assert "MPP-OPEN-FINISH" not in nested_codes
    assert "MPP-CRITICAL-GAP-IN" not in nested_codes
    assert "MPP-CRITICAL-GAP-OUT" not in nested_codes
    coverage = next(issue for issue in issues
                    if issue.code == "MPP-SUMMARY-LINK-COVERAGE")
    assert coverage.severity == "info"


def test_manual_calculation_mode_is_an_error():
    snapshot = MPPSnapshot(
        "Тест", date(2026, 1, 1), date(2026, 1, 1),
        [mpp_task(1, "Старт", 1, milestone=True)],
        calculation_mode=0,
    )
    issue = next(item for item in validate_snapshot(snapshot)
                 if item.code == "MPP-CALCULATION-MODE")
    assert issue.severity == "error"


def test_reporting_milestone_is_not_treated_as_open_finish():
    snapshot = MPPSnapshot("Тест", date(2026, 1, 1), date(2026, 1, 1), [
        mpp_task(1, "Контрольные вехи", 1, summary=True),
        mpp_task(2, "РВЭ получено", 2, milestone=True),
    ])
    issues = validate_snapshot(snapshot)
    assert not any(issue.code == "MPP-OPEN-FINISH" and issue.task_id == 2
                   for issue in issues)
    reporting = next(issue for issue in issues
                     if issue.code == "MPP-REPORTING-MILESTONES")
    assert reporting.severity == "info"
    result = report_dict(snapshot, issues)["result"]
    assert result["infos"] == 1
    assert result["warnings"] == 2  # открытое начало + отсутствие критического пути
