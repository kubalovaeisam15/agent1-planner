# -*- coding: utf-8 -*-
"""Валидатор выдачи ГРП по машинно проверяемой части чек-листа CLAUDE.md §9.

Использование:
    python tools/validate_grp.py <файл.xlsx>
    python tools/validate_grp.py tests/template_parsed.json     # самопроверка на эталоне

Проверяет только то, что проверяется формально. Пункты чек-листа, требующие
предметного суждения (сезонный гейт, здравый смысл критического пути, оценка
уверенности), остаются за планировщиком и здесь НЕ подменяются.

Код возврата: 0 — нарушений уровня ОШИБКА нет; 1 — есть.
"""
from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

LATIN_LINK = re.compile(r"\d+\s*(FS|SS|FF|SF)\b", re.IGNORECASE)
LINK_RE = re.compile(r"^(\d+)\s*([А-Я]{2})?\s*([+-]\s*\d+)?")
VAGUE = re.compile(r"(≥|≤|не\s+менее|не\s+более|более|около|примерно|~)")
DUR_RE = re.compile(r"^\d+([.,]\d+)?\s*(дн|день|дня|дней)\.?\??$", re.IGNORECASE)

# CLAUDE.md §9 «Все обязательные вехи присутствуют»
REQUIRED_MILESTONES = {
    "АК (архитектурный конкурс)": ("архитектурн", "конкурс"),
    "ГПЗУ": ("гпзу",),
    "РС": ("разрешени", "строительств"),
    "Старт СМР": ("старт", "смр"),
    "Завершение монолита": ("монолит",),
    "Замыкание теплового контура": ("теплов", "контур"),
    "Пуск тепла": ("пуск", "тепл"),
    "ЗОС": ("зос",),
    "РВЭ": ("рвэ",),
    "Передача": ("передач",),
}

COLUMNS = ["СДР", "Ид.", "Название задачи", "% завершения", "Длительность",
           "Начало", "Окончание", "Предшественники", "Последователи", "комментарий"]


class Report:
    def __init__(self) -> None:
        self.errors: list[str] = []
        self.warns: list[str] = []
        self.oks: list[str] = []

    def check(self, cond: bool, name: str, detail: str = "", warn_only: bool = False) -> None:
        if cond:
            self.oks.append(name)
        elif warn_only:
            self.warns.append(f"{name}{': ' + detail if detail else ''}")
        else:
            self.errors.append(f"{name}{': ' + detail if detail else ''}")

    def dump(self) -> int:
        for n in self.oks:
            print(f"  OK      {n}")
        for n in self.warns:
            print(f"  ВНИМАНИЕ {n}")
        for n in self.errors:
            print(f"  ОШИБКА  {n}")
        print(f"\nИтог: OK {len(self.oks)} · внимание {len(self.warns)} · ошибок {len(self.errors)}")
        return 1 if self.errors else 0


def cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def load(path: Path) -> list[dict]:
    if path.suffix.lower() == ".json":
        return json.loads(path.read_text(encoding="utf-8"))

    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [cell(c) for c in next(rows)]
    tasks = []
    for raw in rows:
        rec = {h: cell(v) for h, v in zip(header, raw)}
        if not rec.get("СДР") and not rec.get("Название задачи"):
            continue
        tasks.append(rec)
    return tasks


def level_of(sdr: str) -> int:
    return sdr.count(".") + 1


def validate(tasks: list[dict]) -> int:
    r = Report()
    n = len(tasks)
    print(f"Задач: {n}\n")

    # --- Колонка «Уровень структуры» -------------------------------------
    has_level_col = "Уровень структуры" in tasks[0]
    r.check(has_level_col, "колонка «Уровень структуры» присутствует",
            "в выдаче она обязательна — typGRP.md §2.2", warn_only=True)
    for t in tasks:
        t["_lvl"] = int(t["Уровень структуры"]) if has_level_col and t.get("Уровень структуры") \
            else level_of(t["СДР"])

    mismatch = [t["СДР"] for t in tasks if t["_lvl"] != level_of(t["СДР"])]
    r.check(not mismatch, "уровень структуры = числу сегментов СДР",
            f"{len(mismatch)} расхождений, напр. {mismatch[:5]}")

    # --- Иерархия (typGRP.md §2.2) ---------------------------------------
    r.check(tasks[0]["_lvl"] == 1, "уровень первой строки = 1", f"фактически {tasks[0]['_lvl']}")

    jumps = [(i, tasks[i - 1]["СДР"], tasks[i]["СДР"])
             for i in range(1, n) if tasks[i]["_lvl"] - tasks[i - 1]["_lvl"] > 1]
    r.check(not jumps, "уровень не растёт больше чем на 1",
            f"{len(jumps)} скачков, напр. {jumps[:3]}")

    # порядок обхода дерева: родитель непосредственно выше
    seen: dict[str, int] = {}
    order_bad = []
    for i, t in enumerate(tasks):
        sdr = t["СДР"]
        seen[sdr] = i
        if "." in sdr:
            parent = sdr.rsplit(".", 1)[0]
            if parent not in seen:
                order_bad.append(sdr)
    r.check(not order_bad, "строки идут в порядке обхода дерева",
            f"{len(order_bad)} потомков раньше родителя, напр. {order_bad[:5]}")

    # --- Суммарные строки -------------------------------------------------
    children = defaultdict(list)
    all_sdr = {t["СДР"] for t in tasks}
    for t in tasks:
        if "." in t["СДР"]:
            p = t["СДР"].rsplit(".", 1)[0]
            if p in all_sdr:
                children[p].append(t)
    summaries = {s for s in children}

    # §2.2 п.4 в редакции с 31.07.2026: у суммарных строк пусты ДЛИТЕЛЬНОСТЬ и
    # ПРЕДШЕСТВЕННИКИ (именно они ломают импорт). Даты выводятся — решение
    # владельца, они свёрнуты из потомков и нужны для чтения выгрузки.
    dirty = [t["СДР"] for t in tasks if t["СДР"] in summaries and (
        t.get("Длительность") or t.get("Предшественники"))]
    r.check(not dirty, "у суммарных строк пусты длительность и предшественники",
            f"{len(dirty)} нарушений (§2.2 п.4), напр. {dirty[:5]}")

    undated = [t["СДР"] for t in tasks if t["СДР"] in summaries
               and not (t.get("Начало") and t.get("Окончание"))]
    r.check(not undated, "у суммарных строк проставлены даты",
            f"{len(undated)} без дат, напр. {undated[:5]}", warn_only=True)

    orphan_parents = [s for s in summaries if s not in all_sdr]
    r.check(not orphan_parents, "все промежуточные суммарные строки присутствуют",
            f"отсутствуют: {orphan_parents[:5]}")

    # --- Ид. ---------------------------------------------------------------
    ids = [t["Ид."] for t in tasks]
    nums = [int(x) for x in ids if x.isdigit()]
    r.check(len(nums) == n, "все Ид. заполнены и числовые", f"{n - len(nums)} пустых/нечисловых")
    if len(nums) == n:
        gaps = sorted(set(range(min(nums), max(nums) + 1)) - set(nums))
        r.check(not gaps, "столбец Ид. сквозной, без разрывов",
                f"{len(gaps)} пропущено: {gaps[:12]}")
        r.check(len(set(nums)) == len(nums), "Ид. уникальны",
                f"{len(nums) - len(set(nums))} дубликатов")

    present_ids = set(nums)

    # --- Связи -------------------------------------------------------------
    latin = [t["СДР"] for t in tasks if LATIN_LINK.search(t.get("Предшественники", ""))]
    r.check(not latin, "нотация связей русская (ОН · НН · ОО · НО)",
            f"латиница в {len(latin)} строках, напр. {latin[:5]}")

    vague = [t["СДР"] for t in tasks if VAGUE.search(t.get("Предшественники", ""))]
    r.check(not vague, "лаги численные, без формулировок «не менее»",
            f"{len(vague)} строк, напр. {vague[:5]}")

    dangling, to_summary = [], []
    leaf_sdr = all_sdr - summaries
    id_to_sdr = {int(t["Ид."]): t["СДР"] for t in tasks if t["Ид."].isdigit()}
    has_successor = set()
    for t in tasks:
        p = t.get("Предшественники", "")
        if not p:
            continue
        for part in (x.strip() for x in p.split(";")):
            if not part:
                continue
            m = LINK_RE.match(part)
            if not m:
                continue
            pid = int(m.group(1))
            if pid not in present_ids:
                dangling.append((t["СДР"], part))
            else:
                has_successor.add(pid)
                if id_to_sdr.get(pid) in summaries:
                    to_summary.append((t["СДР"], part))

    r.check(not dangling, "все связи ведут на существующие Ид.",
            f"{len(dangling)} битых, напр. {dangling[:5]}")
    # Связь на суммарную строку — норма, а не нарушение: решение владельца
    # 31.07.2026 (typGRP.md §2.2 п.5, §13 расх. 14). Выводится справочно.
    if to_summary:
        r.warns.append(f"справочно: {len(to_summary)} связей ведут на суммарные строки — "
                       f"это норма (typGRP.md §2.2 п.5), напр. {to_summary[:4]}")
    else:
        r.oks.append("связей на суммарные строки нет")

    leaves = [t for t in tasks if t["СДР"] in leaf_sdr]
    no_pred = [t["СДР"] for t in leaves if not t.get("Предшественники")]
    no_succ = [t["СДР"] for t in leaves
               if t["Ид."].isdigit() and int(t["Ид."]) not in has_successor
               and not t.get("Последователи")]
    r.check(len(no_pred) <= 1, "нет задач-листьев без предшественников (кроме стартовой)",
            f"{len(no_pred)} шт., напр. {no_pred[:8]}", warn_only=True)
    r.check(len(no_succ) <= 1, "нет задач-листьев без последователей (кроме финальной)",
            f"{len(no_succ)} шт., напр. {no_succ[:8]}", warn_only=True)

    # --- Длительности ------------------------------------------------------
    bad_dur = [(t["СДР"], t["Длительность"]) for t in leaves
               if t.get("Длительность") and not DUR_RE.match(t["Длительность"])]
    r.check(not bad_dur, "длительности в формате «N дней»",
            f"{len(bad_dur)} нарушений, напр. {bad_dur[:5]}")

    prelim = [t["СДР"] for t in tasks if t.get("Длительность", "").endswith("?")]
    r.check(not prelim, "нет предварительных оценок (суффикс «?»)",
            f"{len(prelim)} шт., напр. {prelim[:5]}", warn_only=True)

    # --- Обязательные вехи -------------------------------------------------
    milestones = [t for t in tasks if t.get("Длительность", "").startswith("0 дн")]
    ms_names = [t["Название задачи"].lower() for t in milestones]
    missing = [name for name, keys in REQUIRED_MILESTONES.items()
               if not any(all(k in nm for k in keys) for nm in ms_names)]
    r.check(not missing, "все обязательные вехи присутствуют", f"нет: {missing}")

    return r.dump()


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    path = Path(sys.argv[1])
    if not path.is_absolute():
        path = ROOT / path
    if not path.exists():
        print(f"Файл не найден: {path}")
        return 2
    print(f"=== Валидация: {path.name} (CLAUDE.md §9, машинная часть) ===\n")
    return validate(load(path))


if __name__ == "__main__":
    sys.exit(main())
