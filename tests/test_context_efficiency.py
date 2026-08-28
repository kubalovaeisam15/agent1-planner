import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_codex_always_on_context_stays_compact():
    policy = (ROOT / "AGENTS.md").read_text(encoding="utf-8")
    assert len(policy.encode("utf-8")) <= 12_000
    assert "context_preflight" in policy
    assert "Не загружай целиком" in policy
    assert len((ROOT / "CLAUDE.md").read_bytes()) <= 4_000


def test_manifest_covers_norms_runtime_and_template():
    manifest = json.loads(
        (ROOT / "instructions" / "context-manifest.json").read_text(encoding="utf-8")
    )
    paths = {item["path"] for item in manifest["files"]}
    assert paths == {
        "AGENTS.md",
        "CLAUDE.md",
        "instructions/typGRP.md",
        "instructions/bindings.md",
        "instructions/standards.md",
        "instructions/agent-policy-full.md",
        "tools/grp_model.py",
        "tools/build_grp.py",
        "tests/project.template.json",
        "tests/template_parsed.json",
        "data/Шаблон ГРП.mpp",
    }
    assert all(len(item["sha256"]) == 64 for item in manifest["files"])


def test_skill_entrypoints_do_not_require_full_context_load():
    for skill in (ROOT / ".agents" / "skills").glob("*/SKILL.md"):
        text = skill.read_text(encoding="utf-8")
        assert "Полностью прочитай" not in text
        assert len(text.encode("utf-8")) <= 3_500
