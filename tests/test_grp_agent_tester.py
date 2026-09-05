from __future__ import annotations

import importlib.util
import json
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / ".agents" / "skills" / "grp-agent-tester" / "scripts" / "campaign.py"
SPEC = importlib.util.spec_from_file_location("grp_agent_campaign", SCRIPT)
assert SPEC and SPEC.loader
campaign = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(campaign)


def test_campaign_generation_is_reproducible(tmp_path: Path) -> None:
    first = tmp_path / "first"
    second = tmp_path / "second"
    campaign.create_campaign(first, count=12, seed=42, include_invalid=True)
    campaign.create_campaign(second, count=12, seed=42, include_invalid=True)

    first_manifest = json.loads((first / "manifest.json").read_text(encoding="utf-8"))
    second_manifest = json.loads((second / "manifest.json").read_text(encoding="utf-8"))
    assert first_manifest == second_manifest
    for case in first_manifest["cases"]:
        assert (first / case["input"]).read_bytes() == (second / case["input"]).read_bytes()


def test_valid_cases_cover_core_dimensions(tmp_path: Path) -> None:
    target = tmp_path / "campaign"
    campaign.create_campaign(target, count=36, seed=20260904, include_invalid=False)
    manifest = json.loads((target / "manifest.json").read_text(encoding="utf-8"))
    dimensions = [case["dimensions"] for case in manifest["cases"]]

    assert {item["corpus_count"] for item in dimensions} == {1, 2, 3}
    assert {item["underground_floors"] for item in dimensions} == {0, 1, 2}
    assert {item["facade"] for item in dimensions} == set(campaign.FACADES)
    assert {item["enclosure"] for item in dimensions} == set(campaign.ENCLOSURES)
    assert {item["finish_share"] for item in dimensions} == {0.0, 0.5, 1.0}
    assert {item["pile_profile"] for item in dimensions} == set(range(len(campaign.PILE_SETS)))


def test_invalid_cases_target_all_blocking_groups() -> None:
    assert {label for label, _ in campaign.make_invalid_specs()} == {
        "missing-corpuses",
        "missing-underground",
        "missing-facade",
        "missing-glazing",
        "missing-piles",
    }


def test_campaign_does_not_overwrite_existing_path(tmp_path: Path) -> None:
    target = tmp_path / "existing"
    target.mkdir()
    with pytest.raises(SystemExit, match="already exists"):
        campaign.create_campaign(target, count=1, seed=1, include_invalid=False)


def test_manifest_path_cannot_escape_campaign(tmp_path: Path) -> None:
    with pytest.raises(SystemExit, match="escapes campaign"):
        campaign.contained_path(tmp_path, "../outside.json")


def test_case_ids_are_restricted() -> None:
    assert campaign.SAFE_CASE_ID.fullmatch("valid-001")
    assert not campaign.SAFE_CASE_ID.fullmatch("../outside")
    assert not campaign.SAFE_CASE_ID.fullmatch("nested/case")


def test_campaign_cannot_be_created_in_corporate_data(tmp_path: Path, monkeypatch) -> None:
    fake_root = tmp_path / "repo"
    data_dir = fake_root / "data"
    data_dir.mkdir(parents=True)
    monkeypatch.setattr(campaign, "ROOT", fake_root)
    with pytest.raises(SystemExit, match="must not be inside data"):
        campaign.create_campaign(data_dir / "campaign", count=1, seed=1, include_invalid=False)


def test_valid_smoke_campaign_runs_excel_and_ir_checks(tmp_path: Path) -> None:
    target = tmp_path / "smoke"
    campaign.create_campaign(target, count=1, seed=20260904, include_invalid=False)
    assert campaign.run_campaign(target, timeout=120) == 0

    result_files = list((target / "runs").glob("*/results.json"))
    assert len(result_files) == 1
    result = json.loads(result_files[0].read_text(encoding="utf-8"))
    assert result["passed"] == 1
    assert result["results"][0]["build_exit"] == 0
    assert result["results"][0]["excel_validation_exit"] == 0
    assert result["results"][0]["ir_validation_exit"] == 0
    assert result["results"][0]["input_sha256"]
    assert result["results"][0]["xlsx_sha256"]
    assert result["results"][0]["ir_sha256"]
