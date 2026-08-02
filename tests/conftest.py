# -*- coding: utf-8 -*-
"""Общая настройка pytest: модули генератора лежат в tools/, а не в пакете."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
