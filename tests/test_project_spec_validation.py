from __future__ import annotations

import copy
import importlib.util
import json
import sys
import tempfile
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import build_grp  # noqa: E402
import mcp_server  # noqa: E402

CAMPAIGN_SCRIPT = ROOT / ".agents" / "skills" / "grp-agent-tester" / "scripts" / "campaign.py"
CAMPAIGN_SPEC = importlib.util.spec_from_file_location("project_spec_campaign", CAMPAIGN_SCRIPT)
assert CAMPAIGN_SPEC and CAMPAIGN_SPEC.loader
campaign = importlib.util.module_from_spec(CAMPAIGN_SPEC)
CAMPAIGN_SPEC.loader.exec_module(campaign)


def test_validate_project_spec_accepts_complete_project() -> None:
    project = json.loads((ROOT / "tests" / "etalon_project.json").read_text(encoding="utf-8"))
    assert build_grp.validate_project_spec(project) == []


@pytest.mark.parametrize("value", [None, True, 1, 1.5, [], {}, ["НВФ"]])
def test_facade_type_errors_are_controlled(value):
    project = json.loads((ROOT / "tests" / "etalon_project.json").read_text(encoding="utf-8"))
    project["фасад"]["тип"] = value
    assert any("фасад.тип" in issue for issue in build_grp.validate_project_spec(project))


@pytest.mark.parametrize("value", [[], {}])
@pytest.mark.parametrize("entrypoint", ["cli", "mcp"])
def test_invalid_facade_never_writes_outputs(value, entrypoint, tmp_path, capsys):
    project = json.loads((ROOT / "tests" / "etalon_project.json").read_text(encoding="utf-8"))
    project["фасад"]["тип"] = value
    spec, xlsx, ir = (tmp_path / name for name in ("input.json", "output.xlsx", "output.ir.json"))
    spec.write_text(json.dumps(project, ensure_ascii=False), encoding="utf-8")
    if entrypoint == "cli":
        assert build_grp.main([str(spec), str(xlsx), "--ir", str(ir)]) == 2
        assert "Traceback" not in capsys.readouterr().err
    else:
        with pytest.raises(mcp_server.ToolError, match="ТЭП не прошёл входную проверку"):
            mcp_server._schedule_build({"spec_path": str(spec), "xlsx_path": str(xlsx), "ir_path": str(ir)})
    assert not xlsx.exists()
    assert not ir.exists()


def test_validate_project_spec_collects_all_blocking_errors() -> None:
    project, _ = campaign.make_spec(0, campaign.random.Random(0))
    del project["корпуса"][0]["этажей_подземных"]
    del project["корпуса"][0]["остекление"]
    del project["фасад"]
    del project["нулевой_цикл"]["сваи"]

    issues = build_grp.validate_project_spec(project)

    assert any("этажей_подземных" in issue for issue in issues)
    assert any("остекление" in issue for issue in issues)
    assert any("фасад.тип" in issue for issue in issues)
    assert any("нулевой_цикл.сваи" in issue for issue in issues)


@pytest.mark.parametrize("label, project", campaign.make_invalid_specs())
def test_each_blocking_group_is_rejected_without_traceback(
        label: str, project: dict, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    spec = tmp_path / f"{label}.json"
    xlsx = tmp_path / f"{label}.xlsx"
    ir = tmp_path / f"{label}.ir.json"
    spec.write_text(json.dumps(project, ensure_ascii=False), encoding="utf-8")

    code = build_grp.main([str(spec), str(xlsx), "--ir", str(ir)])
    captured = capsys.readouterr()

    assert code == 2
    assert "ТЭП не прошёл входную проверку" in captured.err
    assert "Traceback" not in captured.err
    assert not xlsx.exists()
    assert not ir.exists()


def test_validation_does_not_mutate_project() -> None:
    project, _ = campaign.make_spec(0, campaign.random.Random(0))
    before = copy.deepcopy(project)
    build_grp.validate_project_spec(project)
    assert project == before


def test_top_level_glazing_gets_migration_hint() -> None:
    project, _ = campaign.make_spec(0, campaign.random.Random(0))
    project["остекление"] = project["корпуса"][0].pop("остекление")
    issues = build_grp.validate_project_spec(project)
    assert any("верхнеуровневое поле «остекление» не используется" in issue for issue in issues)


def test_empty_pile_list_is_rejected() -> None:
    project, _ = campaign.make_spec(0, campaign.random.Random(0))
    project["нулевой_цикл"]["сваи"] = []
    issues = build_grp.validate_project_spec(project)
    assert any("непустой список" in issue for issue in issues)


def test_full_height_glazing_requires_stained_glass() -> None:
    project, _ = campaign.make_spec(0, campaign.random.Random(0))
    project["корпуса"][0]["остекление"] = {
        "пвх": True,
        "витражи": False,
        "витражи_на_всю_высоту": True,
    }
    issues = build_grp.validate_project_spec(project)
    assert any("не может быть true" in issue for issue in issues)


def test_fresh_mcp_build_uses_the_same_validation() -> None:
    project, _ = campaign.make_spec(0, campaign.random.Random(0))
    del project["корпуса"][0]["остекление"]
    with tempfile.TemporaryDirectory(dir=ROOT) as value:
        temp_dir = Path(value)
        spec = temp_dir / "invalid.json"
        xlsx = temp_dir / "invalid.xlsx"
        ir = temp_dir / "invalid.ir.json"
        spec.write_text(json.dumps(project, ensure_ascii=False), encoding="utf-8")

        with pytest.raises(mcp_server.ToolError, match="ТЭП не прошёл входную проверку"):
            mcp_server._schedule_build({
                "spec_path": str(spec),
                "xlsx_path": str(xlsx),
                "ir_path": str(ir),
            })

        assert not xlsx.exists()
        assert not ir.exists()


@pytest.mark.parametrize('floors', [44, 45, 46, 59, 60, 61])
def test_floor_boundaries_with_different_underground_levels(floors, tmp_path):
    project, _ = campaign.make_spec(1, campaign.random.Random(0))
    for corpus, level in zip(project['корпуса'], (0, 2)):
        corpus['этажей_надземных'] = floors
        corpus['этажей_подземных'] = level
    spec = tmp_path / 'boundary.json'
    spec.write_text(json.dumps(project, ensure_ascii=False), encoding='utf-8')
    assert build_grp.main([str(spec), str(tmp_path / 'result.xlsx'),
                           '--ir', str(tmp_path / 'result.ir.json')]) == 0


@pytest.mark.parametrize('value', [None, True, '2', -1, 1.5])
def test_underground_rejects_wrong_types(value):
    project, _ = campaign.make_spec(0, campaign.random.Random(0))
    project['корпуса'][0]['этажей_подземных'] = value
    assert any('этажей_подземных' in i for i in build_grp.validate_project_spec(project))
