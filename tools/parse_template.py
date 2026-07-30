# -*- coding: utf-8 -*-
"""Разбор Шаблон_ГРП.xlsx в машинный вид + сверка статистики typGRP.md §3.

Выход:
  tests/template_parsed.json  — 1 545 задач с вычисленным «Уровнем структуры»
  печать отчёта сверки в stdout

Уровень структуры считается по правилу typGRP.md §2.2: число сегментов кода СДР.
В самом шаблоне этой колонки нет — она обязательна только в выдаче агента (§2, §3).
"""
from __future__ import annotations

import json
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "instructions" / "Шаблон_ГРП.xlsx"
DST = ROOT / "tests" / "template_parsed.json"

COLUMNS = [
    "СДР", "Ид.", "Название задачи", "% завершения", "Длительность",
    "Начало", "Окончание", "Предшественники", "Последователи", "комментарий",
]

# Заявлено в typGRP.md §3 — сверяем, а не доверяем.
DECLARED = {
    "всего задач": 1545,
    "вех": 292,
    "со связями": 1140,
    "с комментариями": 243,
    "уровни": {1: 19, 2: 84, 3: 284, 4: 655, 5: 317, 6: 90, 7: 58, 8: 38},
}


def outline_level(sdr: str) -> int:
    """typGRP.md §2.2: уровень = число сегментов кода СДР."""
    return sdr.count(".") + 1


def cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def load() -> list[dict]:
    wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [cell(c) for c in next(rows)]
    if header != COLUMNS:
        raise SystemExit(f"Шапка шаблона изменилась.\nОжидалось: {COLUMNS}\nПолучено:  {header}")

    tasks = []
    for raw in rows:
        rec = {name: cell(v) for name, v in zip(COLUMNS, raw)}
        if not rec["СДР"] and not rec["Название задачи"]:
            continue
        rec["Уровень структуры"] = outline_level(rec["СДР"])
        rec["Название задачи"] = rec["Название задачи"].strip()
        tasks.append(rec)
    return tasks


def report(tasks: list[dict]) -> int:
    levels = Counter(t["Уровень структуры"] for t in tasks)
    facts = {
        "всего задач": len(tasks),
        "вех": sum(1 for t in tasks if t["Длительность"].startswith("0 дн")),
        "со связями": sum(1 for t in tasks if t["Предшественники"]),
        "с комментариями": sum(1 for t in tasks if t["комментарий"]),
        "уровни": dict(sorted(levels.items())),
    }

    print("=== Сверка со статистикой typGRP.md §3 ===")
    bad = 0
    for key in ("всего задач", "вех", "со связями", "с комментариями"):
        ok = facts[key] == DECLARED[key]
        bad += 0 if ok else 1
        print(f"{'OK  ' if ok else 'РАСХ'} {key}: факт {facts[key]}, заявлено {DECLARED[key]}")

    ok_levels = facts["уровни"] == DECLARED["уровни"]
    bad += 0 if ok_levels else 1
    print(f"{'OK  ' if ok_levels else 'РАСХ'} распределение по уровням: факт {facts['уровни']}")
    if not ok_levels:
        print(f"     заявлено {DECLARED['уровни']}")

    # Проверки, важные для импорта в MS Project (typGRP.md §2.2)
    print("\n=== Проверки импорта (typGRP.md §2.2) ===")
    first = tasks[0]["Уровень структуры"]
    print(f"{'OK  ' if first == 1 else 'РАСХ'} уровень первой строки: {first}")

    jumps = [
        (i, tasks[i - 1]["СДР"], tasks[i]["СДР"])
        for i in range(1, len(tasks))
        if tasks[i]["Уровень структуры"] - tasks[i - 1]["Уровень структуры"] > 1
    ]
    print(f"{'OK  ' if not jumps else 'РАСХ'} скачков уровня >1: {len(jumps)}")
    for i, prev, cur in jumps[:10]:
        print(f"     строка {i + 2}: {prev} -> {cur}")

    ids = [t["Ид."] for t in tasks]
    seq = ids == [str(i) for i in range(len(tasks))]
    print(f"{'OK  ' if seq else 'РАСХ'} Ид. сквозной от 0 без разрывов")

    has_children = {t["СДР"] for t in tasks}
    parents = set()
    for t in tasks:
        sdr = t["СДР"]
        if "." in sdr:
            parent = sdr.rsplit(".", 1)[0]
            if parent in has_children:
                parents.add(parent)
    dirty = [
        t["СДР"] for t in tasks
        if t["СДР"] in parents and (t["Длительность"] or t["Предшественники"])
    ]
    print(f"{'ИНФО'} суммарных строк с собственной длительностью/связями: {len(dirty)}")
    print("     (в шаблоне это норма; в выдаче агента такие строки должны быть пусты — §2.2 п.4)")

    return bad


def main() -> int:
    tasks = load()
    DST.parent.mkdir(exist_ok=True)
    DST.write_text(
        json.dumps(tasks, ensure_ascii=False, indent=1),
        encoding="utf-8",
    )
    print(f"Записано: {DST.relative_to(ROOT)} ({len(tasks)} задач)\n")
    return 1 if report(tasks) else 0


if __name__ == "__main__":
    sys.exit(main())
