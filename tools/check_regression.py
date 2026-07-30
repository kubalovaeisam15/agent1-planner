# -*- coding: utf-8 -*-
"""Регрессионная проверка по критериям приёмки CLAUDE.md §10.

    python tools/check_regression.py out/ГРП_эталон.xlsx

Сверяет выдачу с tests/etalon_expected.md §1: даты РС и РВЭ и длительности
двух фаз по отдельности. Сквозная длительность не проверяется — DEC-19.

Код возврата: 0 — все жёсткие критерии в допуске; 1 — есть выход за ±14 дн.
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TOL = 14  # ±2 недели — CLAUDE.md §10

# Эталон — tests/etalon_expected.md §1 и §2. Правка «под результат» запрещена.
HARD = {
    "РС": ("7.5", "02.04.2024"),
    "РВЭ": ("1.2.3.1.7", "22.06.2026"),
}
PHASES = {
    "Фаза A (МЗ → РС)": ("1.1.1.4", "7.5", 447),
    "Фаза РС → РВЭ": ("7.5", "1.2.3.1.7", 811),
}
LANDMARKS = {
    "ЯКОРЬ (кровля/парапет)": ("11.4.2.1.1.27", "19.04.2025"),
    "TC_perm": ("11.4.2.1.8", "16.08.2025"),
    "Пуск тепла К1": ("11.4.2.2.9", "20.10.2025"),
    "Завершено свайное поле": ("1.2.2.1.4", "04.07.2024"),
    "Получен ЗОС": ("1.2.3.1.5", "15.06.2026"),
}


def d(s: str):
    return datetime.strptime(s, "%d.%m.%Y").date()


def load(path: Path) -> dict[str, dict]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c else "" for c in next(rows)]
    out = {}
    for raw in rows:
        rec = {h: ("" if v is None else str(v).strip()) for h, v in zip(header, raw)}
        if rec.get("СДР"):
            out[rec["СДР"]] = rec
        if rec.get("Ид.") == "1":
            pass
    return out


def finish_of(tasks: dict, sdr: str) -> str | None:
    r = tasks.get(sdr)
    if not r:
        return None
    return r.get("Окончание") or None


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    path = Path(sys.argv[1])
    if not path.is_absolute():
        path = ROOT / path
    tasks = load(path)

    print(f"=== Регрессия: {path.name} (CLAUDE.md §10, допуск ±{TOL} дн) ===\n")
    failed = 0

    print("-- Жёсткие критерии --")
    for name, (sdr, expected) in HARD.items():
        got = finish_of(tasks, sdr)
        if not got:
            print(f"  НЕТ ДАННЫХ {name}: задача {sdr} отсутствует или без даты")
            failed += 1
            continue
        delta = (d(got) - d(expected)).days
        ok = abs(delta) <= TOL
        failed += 0 if ok else 1
        print(f"  {'OK  ' if ok else 'МИМО'} {name}: {got} против эталона {expected} "
              f"({delta:+d} дн)")

    print("\n-- Длительности фаз (DEC-19, пофазно) --")
    for name, (a_sdr, b_sdr, expected) in PHASES.items():
        a, b = finish_of(tasks, a_sdr), finish_of(tasks, b_sdr)
        if not a or not b:
            print(f"  НЕТ ДАННЫХ {name}")
            failed += 1
            continue
        got = (d(b) - d(a)).days
        delta = got - expected
        ok = abs(delta) <= TOL
        failed += 0 if ok else 1
        print(f"  {'OK  ' if ok else 'МИМО'} {name}: {got} дн против эталона {expected} дн "
              f"({delta:+d} дн)")

    print("\n-- Ориентиры второго уровня (не критерии приёмки) --")
    for name, (sdr, expected) in LANDMARKS.items():
        got = finish_of(tasks, sdr)
        if not got:
            print(f"  —    {name}: нет данных")
            continue
        delta = (d(got) - d(expected)).days
        mark = "OK  " if abs(delta) <= TOL else "!   "
        print(f"  {mark} {name}: {got} против {expected} ({delta:+d} дн)")

    print(f"\nИтог: нарушений жёстких критериев — {failed}")
    if failed:
        print("Разбирается причина в standards.md / bindings.md / входных данных.")
        print("Правка tests/etalon_expected.md «под результат» запрещена (CLAUDE.md §10).")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
