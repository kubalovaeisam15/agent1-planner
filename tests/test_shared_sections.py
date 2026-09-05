"""DEC-40: одинаковые ID не требуются, чтобы обнаружить повтор раздела."""
import json
from datetime import date

import pytest

import build_grp
import parse_template
from schedule_ir import ScheduleProject, ScheduleTask, validate_schedule_ir
from shared_sections import PROJECT_SHARED_SECTIONS, shared_section_errors
from mpp_validator import MPPSnapshot, validate_snapshot
from test_mpp_validator import mpp_task


@pytest.mark.parametrize("name", PROJECT_SHARED_SECTIONS)
def test_duplicate_sections_block_ir_and_mpp(name):
    schedule = ScheduleProject("test", "Тест", date(2026, 8, 1), [
        ScheduleTask(str(i), name, 1, "summary") for i in (1, 2)
    ])
    assert any(i.code == "IR-SHARED-SECTION" and "повторён" in i.message
               for i in validate_schedule_ir(schedule))
    snapshot = MPPSnapshot("Тест", None, None, [
        mpp_task(i, name, 1, summary=True) for i in (1, 2)
    ])
    assert any(i.code == "MPP-SHARED-SECTION" and "повторён" in i.message
               for i in validate_snapshot(snapshot))


def test_shared_root_cannot_be_nested_in_corpus():
    assert shared_section_errors([("Корпус 2", 1), ("ЗОС и РВЭ", 2)])


def test_stage_children_and_normal_repeated_work_are_allowed():
    assert shared_section_errors([
        ("ЗОС и РВЭ", 1), ("1-й Этап", 2), ("Получение РВЭ", 3),
        ("2-й Этап", 2), ("Получение РВЭ", 3),
        ("Корпус 1", 1), ("Монолит", 2), ("Корпус 2", 1), ("Монолит", 2),
    ]) == []


def test_skeleton_is_exact_current_excel_parse():
    assert json.loads(build_grp.SKELETON.read_text(encoding="utf-8")) == parse_template.load()


def test_skeleton_rejects_duplicate_tail_before_scheduling(tmp_path, monkeypatch):
    rows = parse_template.load()
    start = next(i for i, row in enumerate(rows) if row["Название задачи"] == "ЗОС и РВЭ")
    rows.extend(rows[start:])
    path = tmp_path / "bad-skeleton.json"
    path.write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setattr(build_grp, "SKELETON", path)
    with pytest.raises(ValueError, match="DEC-40"):
        build_grp.Build({"старт_проекта": "01.08.2026"}).load_skeleton()


def test_missing_template_section_is_rejected():
    rows = [(name, 1) for name in PROJECT_SHARED_SECTIONS[:-1]]
    assert shared_section_errors(rows, require_all=True)


def test_whitespace_does_not_hide_duplicates():
    assert shared_section_errors([("ЗОС и РВЭ", 1), ("  ЗОС  и РВЭ ", 1)])


def test_corpus_prefix_cannot_hide_shared_section_copy():
    assert shared_section_errors([("К2. ЗОС и РВЭ", 1)])
