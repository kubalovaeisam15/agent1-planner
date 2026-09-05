"""DEC-40: полный ГРП и явно запрошенный частичный фрагмент."""
from datetime import date
import pytest
import validate_grp
from schedule_ir import ScheduleProject, ScheduleTask, validate_schedule_ir
from mpp_validator import MPPSnapshot, validate_snapshot
from shared_sections import PROJECT_SHARED_SECTIONS
from test_mpp_validator import mpp_task


def ir_for(names):
    return ScheduleProject("test", "test", date(2026, 1, 1), [
        ScheduleTask(str(n), name, 1, "milestone", duration_days=0,
                     start=date(2026, 1, 1), finish=date(2026, 1, 1))
        for n, name in enumerate(names, 1)])


@pytest.mark.parametrize("missing", PROJECT_SHARED_SECTIONS)
def test_each_missing_section_blocks_ir_mpp_excel(missing, capsys):
    names = [name for name in PROJECT_SHARED_SECTIONS if name != missing]
    ir = ir_for(names)
    assert any(i.code == "IR-SHARED-SECTION" and missing in i.message for i in validate_schedule_ir(ir))
    snapshot = MPPSnapshot("test", None, None, [mpp_task(n, name, 1) for n, name in enumerate(names, 1)])
    assert any(i.code == "MPP-SHARED-SECTION" and missing in i.message for i in validate_snapshot(snapshot))
    rows = [dict.fromkeys(validate_grp.COLUMNS, "") for _ in names]
    for n, (row, name) in enumerate(zip(rows, names), 1):
        row.update({"Ид.": str(n), "Название задачи": name, "Уровень структуры": 1,
                    "Начало": "01.01.2026", "Окончание": "01.01.2026", "Длительность": "0 дней"})
    assert validate_grp.validate(rows) > 0
    output = capsys.readouterr().out
    assert "DEC-40" in output and missing in output


def test_partial_requires_explicit_opt_out():
    ir = ir_for(["Фрагмент"])
    assert "IR-SHARED-SECTION" in {i.code for i in validate_schedule_ir(ir)}
    assert validate_schedule_ir(ir, require_all_sections=False) == []
    snapshot = MPPSnapshot("test", None, None, [mpp_task(1, "Фрагмент", 1)])
    assert "MPP-SHARED-SECTION" in {i.code for i in validate_snapshot(snapshot)}
    assert "MPP-SHARED-SECTION" not in {i.code for i in validate_snapshot(snapshot, require_all_sections=False)}


def test_partial_still_rejects_duplicate_sections():
    ir = ir_for(["ЗОС и РВЭ", "ЗОС и РВЭ"])
    assert "IR-SHARED-SECTION" in {i.code for i in validate_schedule_ir(ir, require_all_sections=False)}


def test_all_sections_present_have_no_completeness_errors():
    ir = ir_for(PROJECT_SHARED_SECTIONS)
    assert "IR-SHARED-SECTION" not in {i.code for i in validate_schedule_ir(ir)}


def test_empty_ir_reports_all_missing_sections():
    issues = [i for i in validate_schedule_ir(ir_for([])) if i.code == "IR-SHARED-SECTION"]
    assert len(issues) == 7


@pytest.mark.parametrize("name,level", [("ЗОС и РВЭ",2), ("К2. ЗОС и РВЭ",1)])
def test_partial_still_checks_level_and_corpus_prefix(name, level):
    ir = ir_for([name])
    ir.tasks[0].outline_level = level
    assert "IR-SHARED-SECTION" in {i.code for i in validate_schedule_ir(ir, require_all_sections=False)}
    snapshot = MPPSnapshot("test", None, None, [mpp_task(1,name,level)])
    assert "MPP-SHARED-SECTION" in {i.code for i in validate_snapshot(snapshot, require_all_sections=False)}


def test_empty_excel_is_controlled(capsys):
    assert validate_grp.validate([]) == 1
    assert capsys.readouterr().out.count("DEC-40") == 7


@pytest.mark.parametrize("value", [None, "false", 0])
def test_partial_mode_cannot_be_selected_by_wrong_type(value):
    with pytest.raises(ValueError, match="bool"):
        validate_schedule_ir(ir_for([]), require_all_sections=value)


@pytest.mark.parametrize("missing", PROJECT_SHARED_SECTIONS)
def test_full_mcp_and_mspdi_reject_missing_section(missing, tmp_path, monkeypatch):
    import mcp_server
    from mspdi_adapter import schedule_to_mspdi
    ir = ir_for([name for name in PROJECT_SHARED_SECTIONS if name != missing])
    source = tmp_path / "incomplete.ir.json"
    target = tmp_path / "blocked.mpp"
    source.write_text(ir.to_json(), encoding="utf-8")
    result = mcp_server._schedule_validate_ir({"ir_path":str(source)})
    assert result["valid"] is False
    assert any(i["code"] == "IR-SHARED-SECTION" and missing in i["message"] for i in result["issues"])
    def forbidden(*args, **kwargs):
        pytest.fail("Экспортёр не должен запускаться для неполного ГРП")
    monkeypatch.setattr(mcp_server,"export_mpp",forbidden)
    with pytest.raises(mcp_server.ToolError,match="Экспорт остановлен"):
        mcp_server._mpp_export({"ir_path":str(source),"mpp_path":str(target)})
    with pytest.raises(ValueError,match="IR-SHARED-SECTION"):
        schedule_to_mspdi(ir)
    assert not target.exists()
