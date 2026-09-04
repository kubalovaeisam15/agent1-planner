# -*- coding: utf-8 -*-
"""Ограничение ТЗ тендера обеспечивает договор за 15 дней до раннего СМР."""
from __future__ import annotations

import json
from datetime import timedelta
from pathlib import Path

import build_grp

ROOT = Path(__file__).resolve().parents[1]


def prepared_build() -> tuple[build_grp.Build, dict]:
    project = json.loads((ROOT / "tests" / "etalon_project.json").read_text(encoding="utf-8"))
    b = build_grp.Build(project)
    b.load_skeleton()
    b.check_stages()
    b.repair_defects()
    b.inherit_summary_links()
    b.apply_site_conditions()
    b.configure_corpuses()
    b.configure_zero_cycle()
    b.configure_parking()
    b.apply_finishing_scope()
    b.apply_standards()
    b.apply_absent_piles()
    for corpus in project["корпуса"]:
        b.rebuild_monolith(corpus)
    nodes = b.schedule()
    b.thermal(nodes)
    b.wire_zos()
    return b, b.schedule()


def test_roofing_tz_is_anchored_15_days_before_earliest_linked_smr():
    b, baseline = prepared_build()
    contract_i = b.one("Заключение договора Кровля")
    tz_i = b.one("Подготовка ТЗ кровля")
    assert contract_i is not None and tz_i is not None

    contract_key = b.rows[contract_i]["key"]
    linked_smrs = [
        baseline[row["key"]]
        for row in b.rows
        if row.get("dur") and any(
            target == contract_key for target, _, _ in row.get("tpl_links", [])
        )
    ]
    earliest_smr = min(node.start for node in linked_smrs if node.start)

    assert b.constrain_tender_tz(baseline) > 0
    nodes = b.schedule()
    tz = b.rows[tz_i]
    contract = nodes[contract_key]

    assert tz["constraint_type"] == "Начало не ранее"
    assert build_grp.dparse(tz["constraint_date"]) == nodes[tz["key"]].start
    assert contract.finish == earliest_smr - timedelta(days=15)


def test_nomination_tz_does_not_receive_contract_constraint():
    b, baseline = prepared_build()
    nomination_i = b.one("Подготовка ТЗ и ключевых условия номинация КМ к ограждению кровли")
    assert nomination_i is not None

    b.constrain_tender_tz(baseline)

    assert b.rows[nomination_i].get("constraint_type") != "Начало не ранее"
    assert not b.rows[nomination_i].get("constraint_date")
