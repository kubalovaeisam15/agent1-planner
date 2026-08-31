"""Portable output locations; no directory is created at import time."""
from __future__ import annotations

import argparse
import os
from pathlib import Path
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[1]


def _absolute_directory(value: str, source: str) -> Path:
    expanded = os.path.expandvars(os.path.expanduser(value))
    path = Path(expanded)
    if not path.is_absolute() or "%" in expanded or "$" in expanded:
        raise ValueError(f"{source} must be an absolute, fully expanded path")
    path = path.resolve()
    if path == Path(path.anchor):
        raise ValueError(f"{source} must not be a filesystem root")
    return path


def _windows_desktop() -> str | None:
    """Read the user's redirected Desktop (including OneDrive), without a shell."""
    if os.name != "nt":
        return None
    import winreg

    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "Desktop")
            return value if isinstance(value, str) and value.strip() else None
    except OSError:
        return None


def _desktop_dir() -> Path | None:
    # Explicit environment overrides are useful for redirected and non-Windows desktops.
    for key in ("AGENT1_DESKTOP_DIR", "XDG_DESKTOP_DIR"):
        if os.environ.get(key):
            return _absolute_directory(os.environ[key], key)
    redirected = _windows_desktop()
    if redirected:
        return _absolute_directory(redirected, "Windows Desktop")
    for key in ("USERPROFILE", "HOME"):
        if os.environ.get(key):
            return _absolute_directory(os.environ[key], key) / "Desktop"
    try:
        return Path.home().resolve() / "Desktop"
    except RuntimeError:
        return None


def get_output_dir(
    repo_root: str | Path | None = None, *, create: bool = True,
    runtime: str | None = None,
) -> Path:
    """Return Desktop, or <repo>/artifacts/output in cloud/headless mode.

    AGENT1_RUNTIME=cloud is a project setting, not a Codex-provided detection API.
    auto uses cloud output in CI or when Desktop does not exist. local requires
    an existing Desktop. Permission errors propagate; never silently bypass them.
    create=False resolves the same location without any writes.
    """
    mode = (runtime or os.environ.get("AGENT1_RUNTIME", "auto")).strip().lower()
    if mode not in {"auto", "local", "cloud"}:
        raise ValueError("AGENT1_RUNTIME must be auto, local or cloud")
    root = Path(repo_root if repo_root is not None else ROOT).resolve()
    ci = os.environ.get("CI", "").strip().lower() in {"1", "true", "yes"}
    desktop = None if mode == "cloud" or (mode == "auto" and ci) else _desktop_dir()
    if desktop is not None and desktop.is_dir():
        result = desktop.resolve()
    elif mode == "local":
        raise FileNotFoundError("Desktop not found; set AGENT1_DESKTOP_DIR or use cloud mode")
    else:
        result = (root / "artifacts" / "output").resolve()
        if not result.is_relative_to(root):
            raise ValueError("artifacts/output must not point outside the repository")
    if create:
        result.mkdir(parents=True, exist_ok=True)
    return result


def new_output_path(suffix: str, *, repo_root: str | Path | None = None) -> Path:
    """Generate a new name without touching the filesystem."""
    if suffix not in {".xlsx", ".ir.json", ".mpp", ".mpp-report.json"}:
        raise ValueError("Unsupported output suffix")
    return get_output_dir(repo_root, create=False) / f"GRP-{uuid4().hex}{suffix}"


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Print the current output directory")
    parser.add_argument("--runtime", choices=("auto", "local", "cloud"))
    parser.add_argument("--create", action="store_true", help="create the directory if needed")
    args = parser.parse_args()
    print(get_output_dir(runtime=args.runtime, create=args.create))
