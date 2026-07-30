# -*- coding: utf-8 -*-
"""Генератор ГРП: развёртывание каркаса typGRP.md в Excel для импорта в MS Project.

    python tools/build_grp.py tests/etalon_project.json out/ГРП.xlsx

Конвейер:
    1. каркас из эталона (tests/template_parsed.json)
    2. устранение расхождений typGRP.md §13 (расх. 9–11)
    3. пересчёт параметризуемых блоков по standards.md
    4. тиражирование корпусов (typGRP.md §12.1)
    5. сетевой расчёт (CPM), проход 1
    6. тепловой контур и отделка по каждому корпусу (bindings.md §3.5, §3.7)
    7. сетевой расчёт, проход 2
    8. перенумерация Ид., очистка суммарных строк, запись Excel

Что генератор НЕ делает (по решениям владельца, а не по недосмотру):
    · сценарии и вероятностный расчёт — DEC-16, правила не заданы
    · ресурсный контур — CLAUDE.md §4
    · утверждение графика — CLAUDE.md §1
"""
from __future__ import annotations

import json
import math
import sys
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from grp_model import (  # noqa: E402
    Assumption, Bnd, Node, Std, backward_pass, compute_thermal_and_fit, dfmt, dparse,
    fmt_links, forward_pass, monolith_floor_durations, parse_links, pile_duration,
    seasonal_duration,
)

ROOT = Path(__file__).resolve().parents[1]
SKELETON = ROOT / "tests" / "template_parsed.json"

COLUMNS = ["СДР", "Ид.", "Уровень структуры", "Название задачи", "% завершения",
           "Длительность", "Начало", "Окончание", "Предшественники", "Последователи",
           "комментарий"]


class Build:
    def __init__(self, project: dict) -> None:
        self.p = project
        self.assumptions: list[Assumption] = []
        self.rationale: list[dict] = []
        self.notices: list[str] = []
        self.rows: list[dict] = []
        self.start = dparse(project["старт_проекта"])

    # -- служебное ------------------------------------------------------
    def note(self, code: str, text: str, scope: str = "") -> None:
        self.assumptions.append(Assumption(code, text, scope))

    def why(self, block: str, value: str, source: str, confidence: str, comment: str = "") -> None:
        self.rationale.append({"Блок": block, "Значение": value, "Источник": source,
                               "Уверенность": confidence, "Комментарий": comment})

    def by_sdr(self) -> dict[str, dict]:
        return {r["СДР"]: r for r in self.rows}

    # ==================================================================
    # 1. Каркас
    # ==================================================================
    def load_skeleton(self) -> None:
        raw = json.loads(SKELETON.read_text(encoding="utf-8"))
        id2sdr = {int(r["Ид."]): r["СДР"] for r in raw}
        self.rows = []
        for r in raw:
            links = []
            for pid, kind, lag in parse_links(r["Предшественники"]):
                if pid in id2sdr:
                    links.append((id2sdr[pid], kind, lag))
            self.rows.append({
                "СДР": r["СДР"],
                "Название задачи": r["Название задачи"],
                "Длительность": self._dur(r["Длительность"]),
                "links": links,
                "комментарий": r["комментарий"],
                "tpl_start": r["Начало"],
                "источник": "эталон (typGRP.md)",
            })
        self.why("Каркас", f"{len(self.rows)} строк",
                 "typGRP.md §3, шаблон 1 545 задач", "высокая",
                 "Регламентные блоки фазы A наследуют длительности шаблона — typGRP.md §6.0")

    @staticmethod
    def _dur(s: str) -> int | None:
        if not s:
            return None
        digits = "".join(c for c in s if c.isdigit())
        return int(digits) if digits else None

    def summaries(self) -> set[str]:
        all_sdr = {r["СДР"] for r in self.rows}
        out = {s.rsplit(".", 1)[0] for s in all_sdr if "." in s and s.rsplit(".", 1)[0] in all_sdr}
        # Корень проекта: СДР «0» — суммарная строка (typGRP.md §4), но потомков
        # вида «0.x» у неё нет, поэтому по общему правилу она бы стала задачей
        # и потянула бы за собой длительность 1 667 дней из шаблона.
        if "0" in all_sdr:
            out.add("0")
        return out

    # ==================================================================
    # 2. Расхождения эталона — typGRP.md §13, расх. 9–11
    # ==================================================================
    def repair_defects(self) -> None:
        idx = {r["СДР"]: i for i, r in enumerate(self.rows)}

        def insert_block(head_sdr: str, steps: list[tuple[str, int]], name: str) -> list[str]:
            pos = idx[head_sdr]
            self.rows[pos]["Длительность"] = None       # суммарная строка — §2.2 п.4
            self.rows[pos]["links"] = []
            new, prev = [], None
            for i, (label, dur) in enumerate(steps, 1):
                sdr = f"{head_sdr}.{i}"
                links = [(prev, "ОН", 0)] if prev else []
                new.append({"СДР": sdr, "Название задачи": f"{label} {name}",
                            "Длительность": dur, "links": links, "комментарий": "",
                            "tpl_start": "", "источник": "восстановлено, typGRP.md §13 расх. 11"})
                prev = sdr
            self.rows[pos + 1:pos + 1] = new
            return [r["СДР"] for r in new]

        # 10.2.9 — блок номинации, §9.1а (5 шагов)
        nom_steps = [("РД ВПР номинация", 0), ("Подготовка ТЗ и ключевых условий номинация", 7),
                     ("Подготовка оферты номинация", 14), ("Тендер номинация", 30),
                     ("Выбор поставщика / соглашение о намерениях номинация", 14)]
        insert_block("10.2.9", nom_steps, "Арматурный каркас сваи")

        idx = {r["СДР"]: i for i, r in enumerate(self.rows)}
        # 10.2.53 — основной блок, §9.1 (7 шагов)
        main_steps = [("РД ВПР", 0), ("Подготовка ТЗ", 7), ("Подготовка ВОР", 4),
                      ("Подготовка оферты", 3), ("Согласование тендерного пакета", 3),
                      ("Тендер", 30), ("Заключение договора", 14)]
        s53 = insert_block("10.2.53", main_steps, "Ограждение территории")

        # три битые связи → шаги 3, 6, 7 блока 10.2.53 (§13 расх. 10)
        remap = {"1.2.1.5.3": s53[6], "1.3.20.3": s53[2], "1.3.21.3": s53[5]}
        for sdr, target in remap.items():
            row = self.by_sdr()[sdr]
            row["links"] = [(target, "ОН", 0)] + row["links"]

        self.note("typGRP.md §13 расх. 9",
                  "Нумерация Ид. эталона содержала 12 разрывов и в выдачу не перенесена: "
                  "Ид. присвоены заново сквозным рядом, все ссылки пересчитаны.")
        self.note("typGRP.md §13 расх. 10",
                  "Три связи эталона вели на удалённые Ид. (1102, 1105, 1106). Вехи МАФ "
                  "1.2.1.5.3, 1.3.20.3, 1.3.21.3 привязаны к шагам 7, 3 и 6 блока 10.2.53. "
                  "Требует подтверждения владельца.")
        self.note("typGRP.md §13 расх. 11",
                  "Пустые блоки 10.2.9 и 10.2.53 развёрнуты по канону §9.1а и §9.1. "
                  "Требует подтверждения владельца.")
        self.why("Тендерные блоки", "10.2.9 — 5 шагов, 10.2.53 — 7 шагов",
                 "typGRP.md §9.1, §9.1а", "средняя", "Восстановление канона, DEC не выпущен")

    # ==================================================================
    # 3. Пересчёт по standards.md
    # ==================================================================
    def rebuild_monolith(self, corpus_idx: int, corpus: dict) -> None:
        base = f"11.4.{corpus_idx}.1.1"
        n = corpus["этажей_надземных"]
        first, typical, roof = monolith_floor_durations(n, corpus.get("сложный_конструктив", False))
        k = typical[0] if typical else Std.MON_TYPICAL[0]

        old = [r for r in self.rows if r["СДР"].startswith(base + ".")]
        pos = self.rows.index(old[0])
        head_links = old[0]["links"]                     # предшественник 1-го этажа сохраняем
        for r in old:
            self.rows.remove(r)

        new, prev = [], None
        for i in range(1, n + 1):
            dur = first if i == 1 else k
            links = [(prev, "ОН", 0)] if prev else list(head_links)
            sdr = f"{base}.{i}"
            new.append({"СДР": sdr, "Название задачи": f"{i} этаж Монолит", "Длительность": dur,
                        "links": links, "комментарий": f"STD-MON-001: {'первый этаж' if i == 1 else f'K = {k} дн'}",
                        "tpl_start": "", "источник": "standards.md §9 STD-MON-001"})
            prev = sdr
        sdr_roof = f"{base}.{n + 1}"
        new.append({"СДР": sdr_roof, "Название задачи": "Кровля/Парапет Монолит",
                    "Длительность": roof, "links": [(prev, "ОН", 0)],
                    "комментарий": "STD-MON-001: кровля/парапет 15 дн. ЯКОРЬ — последний куб бетона",
                    "tpl_start": "", "источник": "standards.md §9 STD-MON-001"})
        self.rows[pos:pos] = new

        total = first + k * (n - 1) + roof
        self.why(f"Монолит корпуса {corpus['код']}", f"{total} дн ({first} + {k} × {n - 1} + {roof})",
                 "standards.md §9 STD-MON-001", "высокая",
                 f"K = {k} дн ({'сложный конструктив или N > 60' if k == 7 else 'типовой'})")

        # лаги пересчитываются, а не копируются — typGRP.md §12.2
        floor6 = f"{base}.{min(6, n)}"
        floor15 = f"{base}.{min(15, n)}"
        old6, old15 = f"{base}.6", f"{base}.15"
        remapped = 0
        for r in self.rows:
            fixed = []
            for tgt, kind, lag in r["links"]:
                if tgt == old6 and floor6 != old6:
                    tgt, remapped = floor6, remapped + 1
                elif tgt == old15 and floor15 != old15:
                    tgt, remapped = floor15, remapped + 1
                elif tgt.startswith(base + ".") and tgt not in {x["СДР"] for x in new}:
                    tgt, remapped = sdr_roof, remapped + 1   # ссылка на исчезнувший этаж
                fixed.append((tgt, kind, lag))
            r["links"] = fixed
        if remapped:
            self.why(f"Лаги корпуса {corpus['код']}", f"{remapped} связей перепривязано",
                     "typGRP.md §12.2, bindings.md §1.2", "высокая",
                     f"Кладка → этаж min(6, {n}) = {min(6, n)}; ВИС и витражи → этаж "
                     f"min(15, {n}) = {min(15, n)}")

    def apply_standards(self) -> None:
        zc = self.p["нулевой_цикл"]
        wall_type = zc["ограждение_котлована"].lower()
        by = self.by_sdr()

        def set_dur(sdr: str, value: int, source: str) -> None:
            if sdr in by:
                by[sdr]["Длительность"] = value
                by[sdr]["источник"] = source

        # --- ограждение котлована и земляные работы (STD-ZC-001/002) ---
        if wall_type not in Std.PIT_WALL:
            self.note("standards.md §4", f"Тип ограждения котлована «{wall_type}» не распознан — "
                                         f"принят трубошпунт по умолчанию.")
            wall_type = "трубошпунт"
        wall_days, wall_inc, wall_lag = Std.PIT_WALL[wall_type]
        earth_days, earth_inc = Std.EARTH[wall_type]
        set_dur("11.4.1.1.2.1", wall_days, Std.PIT_WALL_SRC)
        set_dur("11.4.1.1.1.1", earth_days, Std.EARTH_SRC)
        if "11.4.1.1.2.1" in by:
            by["11.4.1.1.2.1"]["links"] = [("11.4.1.1.1.1", "НН", wall_lag)]
        self.why("Ограждение котлована", f"{wall_days} дн, НН {wall_lag:+d} дн от земляных",
                 Std.PIT_WALL_SRC, "высокая", f"Тип: {wall_type}")
        self.why("Земляные работы", f"{earth_days} дн", Std.EARTH_SRC, "высокая",
                 f"По типу ограждения, не по этажности подземной части (DEC-03). "
                 f"Приращение старта ФП следующих корпусов +{earth_inc} дн")

        # --- свайное основание (STD-ZC-003/004) ---
        pile_days, pile_trace = pile_duration(zc.get("сваи", []))
        for sdr in ("11.4.1.1.3.1", "11.4.1.1.3.2"):
            set_dur(sdr, pile_days, "standards.md §7 STD-ZC-003/004")
        self.why("Свайное основание", f"{pile_days} дн", "standards.md §7 STD-ZC-003/004",
                 "средняя", " · ".join(pile_trace) + " · статус норматива ⚠, калибровки не проходило")
        if pile_days > Std.PILE_WARN_DAYS[0]:
            msg = (f"Свайное основание {pile_days} дн превышает порог {Std.PILE_WARN_DAYS[0]} дн. "
                   f"Рекомендуется увеличить число буровых установок.")
            self.notices.append(msg)
            self.note("DEC-14", msg)

        # --- ФП и монолит ниже 0.000 (STD-ZC-005/006) ---
        n_max = max(c["этажей_надземных"] for c in self.p["корпуса"])
        raft = Std.RAFT_TALL[0] if n_max > 45 else Std.RAFT[0]
        set_dur("11.4.1.1.4.1", raft, Std.RAFT[1])
        self.why("Фундаментная плита", f"{raft} дн", Std.RAFT[1] if n_max <= 45 else Std.RAFT_TALL[1],
                 "высокая", f"Максимальная этажность {n_max}")
        park = self.p.get("паркинг", {})
        if park.get("есть"):
            lv = park.get("этажей_подземных", 1)
            set_dur("11.4.1.1.4.2", Std.PARKING_RAFT[0] + Std.PARKING_LEVEL[0] * lv,
                    Std.PARKING_RAFT[1])
            self.why("Фундаменты паркинга", f"{Std.PARKING_RAFT[0] + Std.PARKING_LEVEL[0] * lv} дн",
                     "standards.md §10 STD-MON-002", "высокая",
                     f"ФП 45 дн + {lv} подземн. этаж(а) × 45 дн")

        # --- фиксированные нормативы ---
        fixed = [
            ("11.4.2.1.4.1", Std.FACADE_MOCKUP, "МОКАП фасада"),
            ("11.4.2.2.1", Std.ELEVATORS, "Лифты"),
            ("11.4.2.1.9", Std.TEMP_CONTOUR, "ВТК"),
        ]
        for sdr, (val, src), label in fixed:
            if sdr in by:
                set_dur(sdr, val, src)
                self.why(label, f"{val} дн", src, "высокая")

        # --- ИТП корпуса: DEC-06 отменяет «1 день?» шаблона ---
        if "11.4.2.2.3" in by:
            by["11.4.2.2.3"]["Длительность"] = None      # суммарная, значение уходит на потомка
        self.why("ЦТП/ИТП", f"{Std.ITP[0]} дн", Std.ITP[1], "средняя",
                 "DEC-06: комментарий шаблона «1 день?» отменён. Длительность подтверждена "
                 "только для ИТП паркинга и перенесена на корпус без отдельного подтверждения")

        # --- кровля с сезонным коэффициентом SEA-02 ---
        self.why("Кровля", f"{Std.ROOF_PLAIN[0]} дн базовых",
                 Std.ROOF_PLAIN[1] + " · SEA-02 ×2,0 к попавшей части (DEC-20)", "высокая")

    # ==================================================================
    # 4. Тиражирование корпусов — typGRP.md §12.1
    # ==================================================================
    def replicate(self) -> None:
        corpuses = self.p["корпуса"]
        if len(corpuses) <= 1:
            return
        base_rows = [r for r in self.rows if r["СДР"].startswith("11.4.2")]
        tail = self.rows.index(base_rows[-1]) + 1
        earth_inc = Std.EARTH[self.p["нулевой_цикл"]["ограждение_котлована"].lower()][1]

        clones = []
        for n, corpus in enumerate(corpuses[1:], start=2):
            new_idx = n + 1                       # 11.4.3, 11.4.4, ...
            mapping = {r["СДР"]: r["СДР"].replace("11.4.2", f"11.4.{new_idx}", 1)
                       for r in base_rows}
            for r in base_rows:
                clone = dict(r)
                clone["СДР"] = mapping[r["СДР"]]
                clone["Название задачи"] = r["Название задачи"].replace(
                    corpuses[0]["код"], corpus["код"])
                clone["links"] = [(mapping.get(t, t), k, lg) for t, k, lg in r["links"]]
                clone["источник"] = f"тиражирование, typGRP.md §12.1"
                clones.append(clone)
            self.note("typGRP.md §14 п.8",
                      f"Корпус {corpus['код']} развёрнут тиражированием блока 11.4.2 в 11.4.{new_idx}. "
                      f"Правила нумерации 11.4.N и состав вех при 3+ корпусах владельцем "
                      f"не подтверждены.")
            self.why(f"Корпус {corpus['код']}", f"блок 11.4.{new_idx}",
                     "typGRP.md §12.1", "низкая",
                     f"Приращение DEC-10 (+{earth_inc * (n - 1)} дн к старту ФП) НЕ ПРИМЕНЕНО: "
                     f"в шаблоне фундаментная плита одна на все корпуса (11.4.1.1.4.1), "
                     f"покорпусного блока ФП нет. Требуется структурное решение владельца")
            self.note("DEC-10",
                      f"Корпус {corpus['код']}: приращение старта ФП (+{earth_inc * (n - 1)} дн) "
                      f"не применено — в шаблоне отсутствует покорпусный блок фундаментной плиты. "
                      f"Нулевые циклы корпусов не разнесены каскадом.", corpus["код"])
        self.rows[tail:tail] = clones

    # ==================================================================
    # 5. Сетевой расчёт
    # ==================================================================
    def build_nodes(self) -> dict[str, Node]:
        summaries = self.summaries()
        nodes: dict[str, Node] = {}
        for r in self.rows:
            if r["СДР"] in summaries:
                continue
            dur = r["Длительность"] or 0
            n = Node(key=r["СДР"], duration=dur,
                     links=[(t, k, lg) for t, k, lg in r["links"] if t not in summaries])
            if not n.links and r.get("tpl_start"):
                n.anchor_start = dparse(r["tpl_start"])
            nodes[r["СДР"]] = n
        return nodes

    def schedule(self) -> dict[str, Node]:
        nodes = self.build_nodes()
        forward_pass(nodes, self.start)
        backward_pass(nodes)
        return nodes

    def report_anchors(self, nodes: dict[str, Node]) -> None:
        """Задачи без предшественников удерживаются календарной датой эталона.

        Чек-лист CLAUDE.md §9 допускает такие ограничения только «согласованные
        и помеченные», поэтому их число выносится в допущения, а не умалчивается.
        """
        anchored = [k for k, n in nodes.items() if n.anchor_start and not n.links]
        if not anchored:
            return
        self.note("CLAUDE.md §9",
                  f"{len(anchored)} задач-листьев не имеют предшественников в эталоне и "
                  f"удерживаются календарной датой шаблона (регламентные вехи фазы A, "
                  f"контрольные вехи проектирования). Это помеченные ограничения; "
                  f"при поступлении правил привязки они заменяются связями.")
        self.why("Календарные якоря", f"{len(anchored)} задач",
                 "typGRP.md §6.0, регламентные привязки и длительности", "средняя",
                 "Пересчёту по standards.md не подлежат — длительности берутся как есть")

    # ==================================================================
    # 6. Тепловой контур и отделка — bindings.md §3.5, §3.7
    # ==================================================================
    def thermal(self, nodes: dict[str, Node]) -> None:
        by = self.by_sdr()
        glaze = self.p["остекление"]
        facade = self.p["фасад"]["тип"]
        has_parking = self.p.get("паркинг", {}).get("есть", False)

        for ci, corpus in enumerate(self.p["корпуса"], start=2):
            pre = f"11.4.{ci}"
            n_fl = corpus["этажей_надземных"]
            roof_key = f"{pre}.1.1.{n_fl + 1}"
            if roof_key not in nodes:
                continue
            anchor = nodes[roof_key].finish

            ext_masonry = nodes.get(f"{pre}.1.3.1")
            stained = nodes.get(f"{pre}.1.6.2")
            pvc = nodes.get(f"{pre}.1.6.1")
            partitions = nodes.get(f"{pre}.1.3.2")

            # ЯКОРЬ_ОГР — BND-KLD-001 / DEC-05
            if glaze.get("витражи_на_всю_высоту"):
                anchor_ogr = stained.start if stained else anchor
                self.note("BND-SPK-003", f"{corpus['код']}: витражи на всю высоту — наружной кладки "
                                         f"нет, ЯКОРЬ_ОГР = старт монтажа витражей.")
            else:
                anchor_ogr = ext_masonry.start if ext_masonry else anchor

            # финиши остекления и фасада — привязаны к ЯКОРЬ
            pvc_fin = anchor + timedelta(days=Bnd.FIN_PVC[0]) if glaze.get("пвх") else None
            st_fin = anchor + timedelta(days=Bnd.FIN_STAINED[0]) if glaze.get("витражи") else None
            fac_days = (Bnd.FIN_FACADE_MODULAR[0] if facade.lower() in ("модульный", "панельный")
                        else Bnd.FIN_FACADE[0])
            fac_fin = anchor + timedelta(days=fac_days)

            res = compute_thermal_and_fit(
                anchor=anchor, anchor_ogr=anchor_ogr, facade_type=facade,
                pvc_finish=pvc_fin, stained_finish=st_fin, facade_finish=fac_fin,
                partitions_start=partitions.start if partitions else anchor_ogr,
                has_parking=has_parking,
            )
            self.assumptions.extend(
                Assumption(a.code, a.text, corpus["код"]) for a in res.assumptions)

            # Запись в график. Даты не проставляются напрямую: зависимые работы
            # привязываются связью НН к вехе монолита кровли (ЯКОРЬ) с численным
            # лагом — чек-лист CLAUDE.md §9 «привязаны к вехам монолита, а не к датам».
            def link_to_anchor(sdr: str, start: date, dur: int, comment: str) -> None:
                if sdr not in by:
                    return
                # ЯКОРЬ — это ОКОНЧАНИЕ монолита кровли, поэтому связь ОН, а не НН:
                # при НН отсчёт пошёл бы от начала задачи кровли, то есть на её
                # длительность раньше якоря.
                lag = (start - anchor).days
                by[sdr]["Длительность"] = dur
                by[sdr]["links"] = [(roof_key, "ОН", lag)]
                by[sdr]["tpl_start"] = ""
                by[sdr]["комментарий"] = f"{comment} · ОН от ЯКОРЬ, лаг {lag:+d} дн"

            link_to_anchor(f"{pre}.1.8", res.tc_perm, 0,
                           f"TC_perm, BND-TC-001, тип фасада {facade}")
            if res.temp_contour:
                link_to_anchor(f"{pre}.1.9", res.temp_contour[0], Std.TEMP_CONTOUR[0],
                               "ВТК, BND-TC-002: финиш 30.09 года старта отделки, 45 дн")
            elif f"{pre}.1.9" in by:
                self.rows.remove(by[f"{pre}.1.9"])
                self.note("BND-TC-002", "ВТК не развёрнут: условие не выполнено либо "
                                        "развёртывание не даёт эффекта.", corpus["код"])
            link_to_anchor(f"{pre}.2.9", res.heat, 0,
                           "Пуск тепла, BND-TC-003: max(TC_eff, ИТП) + 15 дн; "
                           "отопление исключено (DEC-07)")

            for sdr in (f"{pre}.3.1", f"{pre}.3.2"):
                link_to_anchor(sdr, res.fit_start, res.fit_duration,
                               f"BND-OTD-002/003: старт от ЯКОРЬ_ОГР +90 через сезонный гейт, "
                               f"длительность {res.fit_duration} дн (минимум 210)")

            self.why(f"Тепловой контур {corpus['код']}",
                     f"TC_perm {dfmt(res.tc_perm)} · пуск тепла {dfmt(res.heat)}",
                     "bindings.md §3.5 BND-TC-001…003", "средняя", " · ".join(res.trace))
            self.why(f"Отделка {corpus['код']}",
                     f"{dfmt(res.fit_start)} → {dfmt(res.fit_finish)}, {res.fit_duration} дн",
                     "bindings.md §3.7 BND-OTD-002/003 · standards.md §15.1", "средняя",
                     f"ЯКОРЬ = {dfmt(res.anchor)} · ЯКОРЬ_ОГР = {dfmt(res.anchor_ogr)}")
            if res.overrun_days:
                self.notices.append(
                    f"{corpus['код']}: финиш отделки выходит за ЯКОРЬ + 365 на "
                    f"{res.overrun_days} дн. РВЭ и передача сдвинуты на ту же величину (DEC-09).")

            # SEA-02: кровля с коэффициентом к попавшей части
            for sdr in (f"{pre}.1.7.1", f"{pre}.1.7.2"):
                node = nodes.get(sdr)
                if node and sdr in by and by[sdr]["Длительность"]:
                    base = by[sdr]["Длительность"]
                    adj = seasonal_duration(node.start, base, "SEA-02", 2.0)
                    if adj != base:
                        by[sdr]["Длительность"] = adj
                        by[sdr]["комментарий"] = (f"SEA-02: базовые {base} дн → {adj} дн, "
                                                  f"коэффициент ×2,0 к попавшей в окно части (DEC-20)")
                        self.why(f"Кровля {corpus['код']}", f"{base} → {adj} дн",
                                 "bindings.md §5 SEA-02, DEC-20", "высокая",
                                 f"Старт {dfmt(node.start)}, окно 01.12 – последний день февраля")

    # ==================================================================
    # 7. Финализация
    # ==================================================================
    def finalize(self, nodes: dict[str, Node]) -> list[dict]:
        summaries = self.summaries()
        sdr2id = {r["СДР"]: i + 1 for i, r in enumerate(self.rows)}
        out = []
        for i, r in enumerate(self.rows, start=1):
            sdr = r["СДР"]
            is_sum = sdr in summaries
            n = nodes.get(sdr)
            links = [(sdr2id[t], k, lg) for t, k, lg in r["links"]
                     if t in sdr2id and t not in summaries]
            out.append({
                "СДР": sdr,
                "Ид.": i,
                "Уровень структуры": sdr.count(".") + 1,
                "Название задачи": r["Название задачи"],
                "% завершения": "" if is_sum else 0,
                "Длительность": "" if is_sum else f"{(r['Длительность'] or 0)} дней",
                "Начало": "" if is_sum or not n else dfmt(n.start),
                "Окончание": "" if is_sum or not n else dfmt(n.finish),
                "Предшественники": "" if is_sum else fmt_links(links),
                "Последователи": "",
                # Заливка при импорте в MS Project теряется, поэтому критический путь
                # помечается ещё и текстом — CLAUDE.md §3 «рассчитан, непрерывен, отмечен».
                "комментарий": ("КРИТИЧЕСКИЙ ПУТЬ · " if (n and n.critical and not is_sum) else "")
                               + r.get("комментарий", ""),
                "_critical": bool(n and n.critical) and not is_sum,
                "_source": r.get("источник", ""),
            })
        # обратные ссылки
        succ: dict[int, list[int]] = {}
        for row in out:
            for part in row["Предшественники"].split(";"):
                part = part.strip()
                if not part:
                    continue
                pid = int("".join(c for c in part if c.isdigit())[:6] or 0)
                num = ""
                for c in part:
                    if c.isdigit():
                        num += c
                    else:
                        break
                if num:
                    succ.setdefault(int(num), []).append(row["Ид."])
        for row in out:
            row["Последователи"] = "; ".join(str(x) for x in succ.get(row["Ид."], []))
        return out


# ======================================================================
# Запись Excel
# ======================================================================
def write_excel(path: Path, rows: list[dict], build: Build, nodes: dict[str, Node]) -> None:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ГРП"

    head = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="2F5597")
    sum_font = Font(bold=True)
    crit_fill = PatternFill("solid", fgColor="FFE0E0")

    ws.append(COLUMNS)
    for c in range(1, len(COLUMNS) + 1):
        ws.cell(1, c).font = head
        ws.cell(1, c).fill = head_fill
        ws.cell(1, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r in rows:
        ws.append([r[c] for c in COLUMNS])
        row_i = ws.max_row
        if not r["Длительность"]:
            for c in range(1, len(COLUMNS) + 1):
                ws.cell(row_i, c).font = sum_font
        elif r["_critical"]:
            for c in range(1, len(COLUMNS) + 1):
                ws.cell(row_i, c).fill = crit_fill
        ws.cell(row_i, 4).alignment = Alignment(indent=max(0, r["Уровень структуры"] - 1))

    widths = [16, 7, 9, 62, 11, 13, 12, 12, 30, 24, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # --- Обоснование ---
    wr = wb.create_sheet("Обоснование")
    wr.append(["Блок", "Значение", "Источник", "Уверенность", "Комментарий"])
    for c in range(1, 6):
        wr.cell(1, c).font = head
        wr.cell(1, c).fill = head_fill
    for item in build.rationale:
        wr.append([item["Блок"], item["Значение"], item["Источник"],
                   item["Уверенность"], item["Комментарий"]])
    for i, w in enumerate([34, 40, 46, 14, 110], start=1):
        wr.column_dimensions[get_column_letter(i)].width = w
    wr.freeze_panes = "A2"

    # --- Допущения ---
    wa = wb.create_sheet("Допущения")
    wa.append(["Код", "Область", "Формулировка"])
    for c in range(1, 4):
        wa.cell(1, c).font = head
        wa.cell(1, c).fill = head_fill
    for a in build.assumptions:
        wa.append([a.code, a.scope, a.text])
    wa.append(["DEC-16", "", "Правила сценариев не заданы — блок сценариев и вероятностный "
                             "расчёт не рассчитаны (CLAUDE.md §8.3)."])
    for i, w in enumerate([22, 12, 130], start=1):
        wa.column_dimensions[get_column_letter(i)].width = w
    wa.freeze_panes = "A2"

    # --- Чек-лист ---
    wc = wb.create_sheet("Чек-лист")
    wc.append(["Раздел", "Пункт CLAUDE.md §9", "Результат"])
    for c in range(1, 4):
        wc.cell(1, c).font = head
        wc.cell(1, c).fill = head_fill
    crit = sum(1 for r in rows if r["_critical"])
    leaves = [r for r in rows if r["Длительность"]]
    checks = [
        ("Структура", "Уровень первой строки = 1", rows[0]["Уровень структуры"] == 1),
        ("Структура", "Уровень не растёт больше чем на 1",
         all(rows[i]["Уровень структуры"] - rows[i - 1]["Уровень структуры"] <= 1
             for i in range(1, len(rows)))),
        ("Структура", "У суммарных строк пусты длительность, даты и связи",
         all(not r["Начало"] and not r["Предшественники"]
             for r in rows if not r["Длительность"])),
        ("Структура", "Ид. сквозной и совпадает с номером строки",
         all(r["Ид."] == i + 1 for i, r in enumerate(rows))),
        ("Связи", "Нотация русская (ОН · НН · ОО · НО), латиницы нет",
         not any(x in r["Предшественники"] for r in rows for x in ("FS", "SS", "FF", "SF"))),
        ("Связи", "Лаги численные", not any("≥" in r["Предшественники"] for r in rows)),
        ("Длительности", "Все длительности в календарных днях",
         all(r["Длительность"].endswith("дней") for r in leaves)),
        ("Критический путь", f"Рассчитан ({crit} задач)", crit > 0),
        ("Сценарии", "Исключены по DEC-16 — не проверяются", True),
    ]
    for section, name, ok in checks:
        wc.append([section, name, "OK" if ok else "НАРУШЕНО"])
    for i, w in enumerate([20, 62, 14], start=1):
        wc.column_dimensions[get_column_letter(i)].width = w
    wc.freeze_panes = "A2"

    # --- Сценарии (DEC-16) ---
    wsc = wb.create_sheet("Сценарии")
    wsc["A1"] = "Блок не рассчитан"
    wsc["A1"].font = Font(bold=True)
    wsc["A2"] = ("DEC-16: решение владельца — правила сценариев поступят отдельной инструкцией "
                 "(standards.md §18). До неё три параметрических сценария и вероятностный расчёт "
                 "не строятся, выдача графика по этой причине не блокируется (CLAUDE.md §8.3).")
    wsc.column_dimensions["A"].width = 140
    wsc["A2"].alignment = Alignment(wrap_text=True, vertical="top")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ======================================================================
def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    spec = Path(sys.argv[1])
    if not spec.is_absolute():
        spec = ROOT / spec
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "out" / "ГРП.xlsx"
    if not out.is_absolute():
        out = ROOT / out

    project = json.loads(spec.read_text(encoding="utf-8"))
    b = Build(project)

    print(f"Проект: {project['название']}")
    print(f"Корпусов: {len(project['корпуса'])} · старт {project['старт_проекта']}\n")

    b.load_skeleton()
    b.repair_defects()
    # Порядок существен: нормативы применяются к эталонному блоку 11.4.2 ДО
    # тиражирования, иначе клоны унаследуют длительности шаблона. Монолит
    # перестраивается ПОСЛЕ тиражирования — до него блоков 11.4.3+ не существует.
    b.apply_standards()
    b.replicate()
    for ci, corpus in enumerate(project["корпуса"], start=2):
        b.rebuild_monolith(ci, corpus)

    nodes = b.schedule()                 # проход 1
    b.thermal(nodes)
    nodes = b.schedule()                 # проход 2 — после теплового блока
    b.report_anchors(nodes)
    rows = b.finalize(nodes)

    write_excel(out, rows, b, nodes)

    finish = max(n.finish for n in nodes.values())
    crit = [n for n in nodes.values() if n.critical]
    print(f"Строк: {len(rows)} · вех: {sum(1 for r in rows if r['Длительность'] == '0 дней')}")
    print(f"Критический путь: {len(crit)} задач · финиш проекта {dfmt(finish)}")
    names = {r["СДР"]: r["Название задачи"] for r in b.rows}
    for n in sorted(crit, key=lambda x: x.start):
        print(f"    {n.key:<16} {dfmt(n.start)} → {dfmt(n.finish)}  {names.get(n.key, '')[:46]}")
    print(f"Обоснований: {len(b.rationale)} · допущений: {len(b.assumptions)}")
    if b.notices:
        print("\nУведомления:")
        for n in b.notices:
            print(f"  ! {n}")
    print(f"\nЗаписано: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
