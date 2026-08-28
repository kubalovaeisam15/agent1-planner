import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
HOOK = ROOT / ".codex" / "hooks" / "agent1_hook.py"
TEMPLATE = ROOT / "data" / "Шаблон ГРП.mpp"


def run_hook(event: dict) -> dict:
    completed = subprocess.run(
        [sys.executable, str(HOOK)],
        input=json.dumps(event, ensure_ascii=False),
        text=True,
        encoding="utf-8",
        capture_output=True,
        check=True,
    )
    return json.loads(completed.stdout)


class AgentHookTests(unittest.TestCase):
    def test_session_start_reports_ready_context(self):
        result = run_hook({"hook_event_name": "SessionStart"})
        context = result["hookSpecificOutput"]["additionalContext"]
        self.assertIn("context_preflight", context)
        self.assertIn("Шаблон ГРП.mpp", context)

    def test_export_to_corporate_template_is_denied(self):
        result = run_hook({
            "hook_event_name": "PreToolUse",
            "tool_name": "mcp__agent1-ms-project__mpp_export",
            "tool_input": {"mpp_path": str(TEMPLATE)},
        })
        output = result["hookSpecificOutput"]
        self.assertEqual("deny", output["permissionDecision"])

    def test_export_to_existing_file_is_denied(self):
        with tempfile.NamedTemporaryFile(dir=ROOT, suffix=".mpp") as existing:
            result = run_hook({
                "hook_event_name": "PreToolUse",
                "tool_name": "mcp__agent1-ms-project__mpp_export",
                "tool_input": {"mpp_path": existing.name},
            })
        self.assertEqual("deny", result["hookSpecificOutput"]["permissionDecision"])

    def test_read_only_template_command_is_allowed(self):
        result = run_hook({
            "hook_event_name": "PreToolUse",
            "tool_name": "PowerShell",
            "tool_input": {"command": f"Get-FileHash -LiteralPath '{TEMPLATE}'"},
        })
        self.assertEqual({}, result)

    def test_mutating_template_command_is_denied(self):
        result = run_hook({
            "hook_event_name": "PreToolUse",
            "tool_name": "PowerShell",
            "tool_input": {"command": f"Remove-Item -LiteralPath '{TEMPLATE}'"},
        })
        self.assertEqual("deny", result["hookSpecificOutput"]["permissionDecision"])

    def test_post_export_requires_validation(self):
        result = run_hook({
            "hook_event_name": "PostToolUse",
            "tool_name": "mcp__agent1-ms-project__mpp_export",
        })
        self.assertIn("mpp_validate", result["hookSpecificOutput"]["additionalContext"])


if __name__ == "__main__":
    unittest.main()
