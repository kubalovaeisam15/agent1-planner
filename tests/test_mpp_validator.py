# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import date
from pathlib import Path

from mpp_validator import compare_with_ir, duration_minutes, parse_mspdi, validate_snapshot
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
      <ConstraintType>0</ConstraintType><Deadline>2026-01-10T00:00:00</Deadline>
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
