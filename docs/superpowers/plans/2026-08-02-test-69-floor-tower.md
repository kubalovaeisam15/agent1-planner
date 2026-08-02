# Тестовый прогон «башня 69 этажей, без отделки квартир» — план реализации

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Довести генератор ГРП до чистой выдачи по проекту «одна башня 69 этажей сложного конструктива, встроенный двухуровневый паркинг, 90 свай БНС, без отделки квартир», не сломав регрессию эталона.

**Architecture:** Генератор `tools/build_grp.py` собирает график из каркаса шаблона v2 (`tests/template_parsed.json`, 1 673 строки), применяет нормативы `tools/grp_model.py`, считает сеть двумя проходами и пишет Excel. Новое правило `DEC-30` встраивается в конвейер отдельным шагом `apply_finishing_scope()` между `configure_parking()` и `apply_standards()`; оно снимает ветку чистовой отделки квартир через общий помощник `drop_rows()`, который переписывает внешние связи вместо их молчаливой потери. Проверки уходят в `tools/validate_grp.py` и в новый pytest-набор.

**Tech Stack:** Python 3.12, `openpyxl` 3.1.5, `pytest` 9.1.1 (устанавливается в задаче 2). Скрипты запускаются из корня репозитория.

**Спека:** `docs/superpowers/specs/2026-08-02-test-69-floor-tower-design.md`

## Global Constraints

Действуют для каждой задачи без исключения.

1. **Регрессия эталона после каждой задачи.** `python tools/build_grp.py tests/etalon_project.json out/ГРП_эталон.xlsx`, затем `python tools/check_regression.py out/ГРП_эталон.xlsx` — все четыре критерия `CLAUDE.md` §10 в допуске: РС ±14 дн, РВЭ ±14 дн, фаза A ±14 дн, фаза РС→РВЭ ±14 дн. Эталонные значения: РС 02.04.2024, РВЭ 31.12.2026, фаза A 447 дн, фаза РС→РВЭ 1 003 дн.
2. **Валидатор.** На эталоне — ноль ошибок **после каждой задачи**: `python tools/validate_grp.py out/ГРП_эталон.xlsx`, код возврата 0. На проекте 69 (`python tools/validate_grp.py out/ГРП_69.xlsx`) ноль ошибок требуется **начиная с задачи 6**: до неё краевые ветки ещё чинятся, и ошибки, отнесённые в отчёте задачи 1 к паркингу, высотности или остаточным поломкам, допустимы. Предупреждения («ВНИМАНИЕ») ошибками не считаются нигде.
3. **Ни одной цифры вне `CLAUDE.md` §0.3.** Каждая длительность и каждый лаг трассируется до пункта файла контекста, кода `DEC-*` или `R-*`. Оценки «по опыту», «по аналогии», «правдоподобное значение» запрещены. Нет правила — фиксируется открытый вопрос через `self.note(...)` и значение не выдумывается.
4. **Правило одной цифры (`CLAUDE.md` §0.2).** Расхождение между двумя файлами контекста, не покрытое `DEC-*` или `R-*`, — остановка и вопрос владельцу, а не выбор.
5. **Нет молчаливого отбрасывания (`CLAUDE.md` §12).** Удалённая задача, снятая связь и неперенесённая зависимость обязаны попасть в лист «Обоснование» через `self.why(...)` либо в «Допущения» через `self.note(...)`.
6. **Версии комплекта согласованы.** Правка `instructions/typGRP.md`, `instructions/bindings.md` или `instructions/standards.md` в том же коммите поднимает версию в шапке файла, таблицу `CLAUDE.md` §0 и сводку версий в `README.md`. Иначе агент обязан остановиться по §0 — промежуточных коммитов с рассогласованными версиями быть не должно.
7. **Файлы контекста правятся только при реальном расхождении**, а не для удобства реализации.
8. **Кодировка.** Все файлы читаются и пишутся в UTF-8. Вывод в консоль уже защищён `sys.stdout.reconfigure(errors="replace")` в существующих скриптах — в новых делать так же.
9. **Коммиты** — по-русски латиницей (как в истории репозитория), с завершающей строкой `Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>`.

## Файловая структура

| Файл | Ответственность | Действие |
| --- | --- | --- |
| `tests/project_69.json` | Вход генератора для тестового проекта | создать (задача 1) |
| `tests/conftest.py` | Путь к `tools/` для pytest | создать (задача 2) |
| `tests/test_drop_rows.py` | Юнит-тесты помощника удаления строк | создать (задача 2) |
| `tests/test_dec30.py` | Тесты правила `DEC-30` на собранном графике | создать (задача 3) |
| `tests/test_project69.py` | Тесты краевых веток проекта 69 этажей | создать (задачи 4–5) |
| `tools/build_grp.py` | Сборка графика | изменить (задачи 2–5) |
| `tools/grp_model.py` | Нормативы и сетевой расчёт | изменить при необходимости (задачи 4–5) |
| `tools/validate_grp.py` | Машинная часть чек-листа §9 | изменить (задача 3) |
| `instructions/typGRP.md` | Карта применимости блока отделки квартир | изменить (задача 3) |
| `instructions/standards.md` | §15.3 передача помещений | изменить (задача 3) |
| `CLAUDE.md` | §0 таблица версий, §0.4 индекс `DEC-30` | изменить (задача 3) |
| `README.md` | Сводка версий, состав, состояние генератора | изменить (задачи 3, 7) |
| `docs/superpowers/reports/task-1-diagnostics.md` | Отчёт диагностического прогона | создать (задача 1) |

---

### Task 1: Вход и диагностический прогон

Задача не меняет ни строки кода. Её результат — воспроизводимый снимок того, что генератор делает с новым входом сегодня. Без него задачи 4–6 не будут знать, что чинить.

**Files:**
- Create: `tests/project_69.json`
- Create: `docs/superpowers/reports/task-1-diagnostics.md`

**Interfaces:**
- Consumes: ничего.
- Produces: `tests/project_69.json` — вход для всех последующих задач; `docs/superpowers/reports/task-1-diagnostics.md` — список поломок, на который опираются задачи 4–6.

- [ ] **Step 1: Создать вход `tests/project_69.json`**

Записать ровно этот файл. Значения происходят из спеки §2: пять блокирующих параметров заданы владельцем, остальные — умолчания `standards.md` §3.

```json
{
  "_комментарий": "Тестовый вход для проверки агента на краевых ветках. Пять блокирующих параметров CLAUDE.md §2.2 заданы пользователем 02.08.2026: 1 корпус, 69 надземных этажей, сложный конструктив, фасад НВФ, остекление ПВХ, 90 свай БНС. Остальное — умолчания standards.md §3.",
  "название": "Тестовый проект: башня 69 этажей, сложный конструктив",
  "старт_проекта": "02.08.2026",

  "этапы_ввода": [
    {
      "номер": 1,
      "объекты": ["К1"]
    }
  ],

  "корпуса": [
    {
      "код": "К1",
      "этажей_надземных": 69,
      "секций": 1,
      "сложный_конструктив": true,
      "_секции": "Умолчание standards.md §3 — 1 секция, статус ⚠️. Каскад BND-MON-002 не применяется.",
      "остекление": {"пвх": true, "витражи": false, "витражи_на_всю_высоту": false},
      "_остекление": "Окна ПВХ, витражей нет. Следствие: BND-KLD-001 применяется, ЯКОРЬ_ОГР = финиш кладки наружных стен."
    }
  ],

  "паркинг": {"есть": false, "код": "П1", "этажей_подземных": 2},
  "_паркинг": "Паркинг встроенный: два подземных уровня корпуса и есть двухуровневый паркинг. Отдельным объектом не выделяется, иначе нулевой цикл удвоился бы. Нулевой цикл и ИТП переносятся в корпус — DEC-06, R-05.",

  "стилобат": {"есть": false},
  "_стилобат": "DEC-21: стилобат учитывается только из входных данных, здесь его нет.",

  "нулевой_цикл": {
    "этажей_подземных": 2,
    "ограждение_котлована": "трубошпунт",
    "сваи": [
      {"тип": "БНС", "количество": 90, "установок": 1}
    ]
  },
  "_нулевой_цикл": "Ограждение задано пользователем и совпадает с умолчанием standards.md §3. Буровых установок — 1 на очередь, умолчание standards.md §3.",

  "фасад": {"тип": "НВФ"},

  "отделка": {"доля_квартир_с_чистовой": 0.0},
  "_отделка": "Квартиры передаются без отделки. Применяется DEC-30: ветка чистовой отделки квартир в график не входит, передача = РВЭ + 180 дн (standards.md §15.3).",

  "индивидуальные_условия": {
    "ППТ": false,
    "снос_застройки": false,
    "вынос_сетей": false,
    "зона_метро": false,
    "ВРИ": false,
    "ПЗЗ": false,
    "ОКН": false,
    "СЗЗ": false
  },
  "_индивидуальные_условия": "Пользователем не заданы → все неприменимы, каждое даёт строку в «Допущения» (CLAUDE.md §2.3). У эталона ППТ и СЗЗ включены и лежат на критическом пути фазы A — здесь дата РС выйдет заметно раньше эталонной, это ожидаемое поведение."
}
```

- [ ] **Step 2: Прогнать генератор и записать вывод**

```bash
mkdir -p out docs/superpowers/reports
python tools/build_grp.py tests/project_69.json out/ГРП_69.xlsx > out/build_69.log 2>&1; echo "exit=$?"
cat out/build_69.log
```

Ожидаемое: либо успешная сборка с блоком «Чек-лист НАРУШЕН» / «Уведомления», либо трассировка исключения. Оба исхода допустимы — это диагностика, а не приёмка.

- [ ] **Step 3: Прогнать валидатор**

```bash
python tools/validate_grp.py out/ГРП_69.xlsx > out/validate_69.log 2>&1; echo "exit=$?"
cat out/validate_69.log
```

Если Step 2 упал с исключением и `out/ГРП_69.xlsx` не создан — записать это в отчёт и перейти к Step 5.

- [ ] **Step 4: Убедиться, что регрессия эталона в порядке ДО правок**

```bash
python tools/build_grp.py tests/etalon_project.json out/ГРП_эталон.xlsx
python tools/check_regression.py out/ГРП_эталон.xlsx; echo "exit=$?"
python tools/validate_grp.py out/ГРП_эталон.xlsx; echo "exit=$?"
```

Ожидаемое: `check_regression` — все четыре критерия ✅, код возврата 0. `validate_grp` — код возврата 0. Это базовая линия: если она уже красная, дальше идти нельзя — сообщить об этом статусом `BLOCKED`.

- [ ] **Step 5: Написать отчёт `docs/superpowers/reports/task-1-diagnostics.md`**

Отчёт пишется по этому шаблону, с фактическим выводом, а не пересказом:

```markdown
# Диагностический прогон `tests/project_69.json`

Дата: <дата> · Коммит базы: <git rev-parse --short HEAD>

## 1. Сборка

Команда: `python tools/build_grp.py tests/project_69.json out/ГРП_69.xlsx`
Код возврата: <N>

<полный вывод или трассировка, дословно>

## 2. Валидатор

Команда: `python tools/validate_grp.py out/ГРП_69.xlsx`
Код возврата: <N>

<полный список ОШИБКА и ВНИМАНИЕ, дословно>

## 3. Базовая линия эталона

check_regression: <четыре строки таблицы с отклонениями>
validate_grp: <итоговая строка «Итог: OK N · внимание N · ошибок N»>

## 4. Перечень поломок

| № | Симптом | Где | Похоже на зону риска спеки |
| --- | --- | --- | --- |
| 1 | ... | ... | отделка квартир / паркинг / высотность / прочее |

## 5. Поломки, не покрытые задачами 2–5

<список, который заберёт задача 6; либо «нет»>
```

Колонка «Похоже на зону риска» заполняется по спеке §3: `отделка квартир` → задача 3, `встроенный паркинг` → задача 4, `высотность` → задача 5, всё прочее → задача 6.

- [ ] **Step 6: Коммит**

```bash
git add tests/project_69.json docs/superpowers/reports/task-1-diagnostics.md
git commit -m "Task 1: vhod project_69.json i diagnosticheskiy progon

Vhod dlya testovogo proekta (1 korpus 69 etazhey, slozhnyy konstruktiv,
vstroennyy parking, 90 svay BNS, bez otdelki kvartir) i otchet o tom,
chto generator delaet s nim segodnya. Kod ne menyalsya.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 2: Помощник `drop_rows` — удаление строк без потери связей

`finalize()` в `tools/build_grp.py:1073` отбрасывает связи на несуществующие строки молча: `links = [(key2id[t], k, lg) for t, k, lg in r["links"] if t in key2id]`. Для `DEC-30` этого недостаточно — `CLAUDE.md` §12 требует, чтобы каждая снятая зависимость была записана. Задача добавляет помощник, который удаляет поддерево и осознанно переписывает внешние связи.

**Files:**
- Create: `tests/conftest.py`
- Create: `tests/test_drop_rows.py`
- Modify: `tools/build_grp.py` — добавить методы `_live_preds()` и `drop_rows()` в класс `Build`, сразу после `set_dur()` (строка 551–554)
- Modify: `README.md` — раздел «Требования», добавить `pytest`

**Interfaces:**
- Consumes: класс `Build` из `tools/build_grp.py`; структура строки — словарь с ключами `key`, `lvl`, `name`, `dur`, `links`, `comment`, `tpl_start`, `src`; связь — кортеж `(key: str, kind: str, lag: int)`; методы `subtree(i)`, `why(...)`, `note(...)`.
- Produces: `Build.drop_rows(roots: list[int], reason: str) -> int` — удаляет строки с индексами `roots` вместе с потомками, возвращает число удалённых строк. Используется задачей 3.

- [ ] **Step 1: Установить pytest**

```bash
python -m pip install pytest
python -c "import pytest; print(pytest.__version__)"
```

Ожидаемое: печатается версия (проверено — 9.1.1).

- [ ] **Step 2: Создать `tests/conftest.py`**

```python
# -*- coding: utf-8 -*-
"""Общая настройка pytest: модули генератора лежат в tools/, а не в пакете."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
```

- [ ] **Step 3: Написать падающий тест `tests/test_drop_rows.py`**

Тест строит `Build` в обход `__init__` (он требует полного словаря проекта), подменяя только то, что нужно помощнику.

```python
# -*- coding: utf-8 -*-
"""Юнит-тесты Build.drop_rows — удаление строк без молчаливой потери связей."""
from __future__ import annotations

import build_grp


def make_build(rows):
    """Build без прогона __init__: помощнику нужны только rows и журналы."""
    b = build_grp.Build.__new__(build_grp.Build)
    b.rows = rows
    b.assumptions = []
    b.rationale = []
    b.notices = []
    return b


def row(key, lvl, name, links=()):
    return {"key": key, "lvl": lvl, "name": name, "dur": 1,
            "links": [tuple(x) for x in links], "comment": "",
            "tpl_start": "", "src": ""}


def test_udalyaet_podderevo_celikom():
    b = make_build([
        row("a", 1, "Корень"),
        row("b", 2, "Ветка под удаление"),
        row("c", 3, "Лист ветки"),
        row("d", 2, "Соседняя ветка"),
    ])
    removed = b.drop_rows([1], "тест")
    assert removed == 2
    assert [r["key"] for r in b.rows] == ["a", "d"]


def test_potrebitel_s_drugimi_predshestvennikami_teryaet_tolko_bituyu_svyaz():
    b = make_build([
        row("live", 1, "Живой предшественник"),
        row("dead", 1, "Удаляемая", [("live", "ОН", 0)]),
        row("user", 1, "Потребитель", [("live", "ОН", 0), ("dead", "ОН", 0)]),
    ])
    b.drop_rows([1], "тест")
    assert b.rows[-1]["links"] == [("live", "ОН", 0)]
    assert any("Потребитель" in w["Комментарий"] for w in b.rationale)


def test_potrebitel_bez_drugih_svyazey_nasleduet_predshestvennikov_udalyaemoy():
    b = make_build([
        row("root", 1, "Корень цепочки"),
        row("dead", 1, "Удаляемая", [("root", "НН", 5)]),
        row("user", 1, "Потребитель", [("dead", "НН", 0)]),
    ])
    b.drop_rows([1], "тест")
    assert b.rows[-1]["links"] == [("root", "НН", 5)]


def test_nasledovanie_prohodit_skvoz_cepochku_udalyaemyh():
    b = make_build([
        row("root", 1, "Живой корень"),
        row("dead1", 1, "Удаляемая внешняя", [("root", "ОН", 3)]),
        row("dead2", 2, "Удаляемая вложенная", [("dead1", "ОН", 0)]),
        row("user", 1, "Потребитель", [("dead2", "ОН", 0)]),
    ])
    b.drop_rows([1], "тест")
    assert [r["key"] for r in b.rows] == ["root", "user"]
    assert b.rows[-1]["links"] == [("root", "ОН", 3)]


def test_pustoy_spisok_korney_nichego_ne_menyaet():
    b = make_build([row("a", 1, "Одна строка")])
    assert b.drop_rows([], "тест") == 0
    assert len(b.rows) == 1
```

- [ ] **Step 4: Прогнать тесты — убедиться, что падают**

```bash
python -m pytest tests/test_drop_rows.py -v
```

Ожидаемое: FAIL, `AttributeError: 'Build' object has no attribute 'drop_rows'` во всех пяти тестах.

- [ ] **Step 5: Реализовать `_live_preds` и `drop_rows`**

Вставить в класс `Build` в `tools/build_grp.py` сразу после метода `set_dur` (после строки 554), перед `def apply_standards`:

```python
    @staticmethod
    def _live_preds(key: str, dead: dict[str, list[tuple[str, str, int]]],
                    seen: set[str] | None = None) -> list[tuple[str, str, int]]:
        """Предшественники удаляемой строки, разрешённые до живых строк.

        Если предшественник удаляемой строки сам удаляется, спуск продолжается
        вглубь: иначе цепочка «живой → удаляемая → удаляемая → потребитель»
        оставила бы потребителя вовсе без связей.
        """
        seen = seen if seen is not None else set()
        out: list[tuple[str, str, int]] = []
        for pt, pk, plg in dead.get(key, []):
            if pt in seen:
                continue
            if pt in dead:
                # Копия множества на ветку, а не мутация общего: сходящиеся
                # ветви (диамант через две удаляемые строки на один и тот же
                # живой узел) обязаны разрешаться независимо, иначе вторая
                # ветвь молча теряется на проверке "уже видели".
                out.extend(Build._live_preds(pt, dead, seen | {pt}))
            else:
                out.append((pt, pk, plg))
        return out

    def drop_rows(self, roots: list[int], reason: str) -> int:
        """Удаляет строки `roots` вместе с потомками и переписывает внешние связи.

        CLAUDE.md §12 запрещает молчаливое отбрасывание связей. finalize()
        отбросил бы ссылки на удалённые строки без следа, поэтому решение
        принимается здесь и записывается в «Обоснование»:
          · у потребителя остаются другие предшественники → битая связь снимается;
          · других предшественников нет → потребитель наследует предшественников
            удаляемой строки (транзитивно, до первой живой).
        """
        doomed: set[int] = set()
        for i in roots:
            doomed.add(i)
            doomed.update(self.subtree(i))
        if not doomed:
            return 0

        dead = {self.rows[i]["key"]: list(self.rows[i]["links"]) for i in doomed}
        cut: list[str] = []
        rewired: list[str] = []
        for i, r in enumerate(self.rows):
            if i in doomed:
                continue
            lost = [ln for ln in r["links"] if ln[0] in dead]
            if not lost:
                continue
            kept = [ln for ln in r["links"] if ln[0] not in dead]
            if kept:
                r["links"] = kept
                cut.append(r["name"])
            else:
                inherited: list[tuple[str, str, int]] = []
                for t, _, _ in lost:
                    for ln in self._live_preds(t, dead):
                        if ln not in inherited:
                            inherited.append(ln)
                r["links"] = inherited
                rewired.append(r["name"])

        for i in sorted(doomed, reverse=True):
            del self.rows[i]

        self.why(f"Удаление блока ({reason})",
                 f"{len(doomed)} строк снято",
                 f"{reason} · CLAUDE.md §12", "средняя",
                 f"Связей снято у сохранённых задач: {len(cut)}"
                 + (f" — {'; '.join(cut[:6])}" if cut else "")
                 + f". Связей переопределено на предшественников удалённых: {len(rewired)}"
                 + (f" — {'; '.join(rewired[:6])}" if rewired else "")
                 + ". Ни одна зависимость не отброшена молча.")
        return len(doomed)
```

- [ ] **Step 6: Прогнать тесты — убедиться, что проходят**

```bash
python -m pytest tests/test_drop_rows.py -v
```

Ожидаемое: 5 passed.

- [ ] **Step 7: Проверить, что ничего не сломалось**

```bash
python tools/build_grp.py tests/etalon_project.json out/ГРП_эталон.xlsx
python tools/check_regression.py out/ГРП_эталон.xlsx; echo "regression exit=$?"
python tools/validate_grp.py out/ГРП_эталон.xlsx; echo "validate exit=$?"
```

Ожидаемое: оба кода возврата 0, четыре критерия ✅. Помощник ещё нигде не вызывается, поэтому выдача обязана совпасть с базовой линией задачи 1.

- [ ] **Step 8: Дополнить `README.md`**

В разделе «Требования» заменить строку

```
- Python 3.12+, `openpyxl` (`pip install openpyxl`)
```

на

```
- Python 3.12+, `openpyxl`, `pytest` (`pip install openpyxl pytest`)
```

В разделе «Команды» добавить в конец блока:

```powershell
# юнит-тесты генератора
python -m pytest tests -q
```

- [ ] **Step 9: Коммит**

```bash
git add tests/conftest.py tests/test_drop_rows.py tools/build_grp.py README.md
git commit -m "Task 2: pomoshchnik drop_rows - udalenie strok bez poteri svyazey

finalize() otbrasyval svyazi na nesushchestvuyushchie stroki molcha, chto
zapreshchaet CLAUDE.md 12. drop_rows udalyaet podderevo i pereopredelyaet
vneshnie svyazi: potrebitel s drugimi predshestvennikami teryaet tolko
bituyu svyaz, potrebitel bez nih nasleduet predshestvennikov udalyaemoy
tranzitivno. Kazhdyy sluchay popadaet v list Obosnovanie.

Dobavlena infrastruktura pytest.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 3: `DEC-30` — применимость блока отделки квартир

Правило и его реализация делаются одной задачей: `CLAUDE.md` §0 обязывает агента останавливаться при рассогласовании версий комплекта, поэтому коммит, где код умеет то, чего нет в нормативах (или наоборот), недопустим.

**Files:**
- Create: `tests/test_dec30.py`
- Modify: `tools/build_grp.py` — константы `FLAT_FIT_PER_CORPUS`, `FLAT_FIT_GLOBAL` и метод `apply_finishing_scope()` в классе `Build`; вызов в `main()`
- Modify: `tools/validate_grp.py` — две новые проверки в `validate()`
- Modify: `instructions/typGRP.md` — новый раздел «Применимость блока отделки квартир», версия v4.1 → v4.2
- Modify: `instructions/standards.md` — §15.3, версия v3.1 → v3.2
- Modify: `CLAUDE.md` — §0 таблица версий, §0.4 строка `DEC-30`, шапка v7.1 → v7.2
- Modify: `README.md` — сводка версий

**Interfaces:**
- Consumes: `Build.drop_rows(roots, reason)` из задачи 2; `Build.find(name, prefix=None, exact=True)`; `Build.corpus_prefix: dict[str, str]` — код корпуса → префикс в графике, заполняется в `configure_corpuses()`.
- Produces: `Build.apply_finishing_scope() -> None` — вызывается в `main()` между `configure_parking()` и `apply_standards()`. После неё в `self.rows` нет ни одной строки ветки чистовой отделки квартир, если `отделка.доля_квартир_с_чистовой == 0`.

- [ ] **Step 1: Написать падающий тест `tests/test_dec30.py`**

```python
# -*- coding: utf-8 -*-
"""DEC-30: при доле чистовой отделки 0 ветка отделки квартир в график не входит."""
from __future__ import annotations

import json
from pathlib import Path

import build_grp

ROOT = Path(__file__).resolve().parents[1]

# Маркеры ветки чистовой отделки квартир, в нижнем регистре.
MARKERS = (
    "отделочные работы квартиры",
    "подготовка под чистовую отделку квартир",
    "чистовая отделка квартир",
    "тендер отделка квартиры чистовая",
    "мокап отделки типового этажа",
    "завершены отделочные работы квартир",
    "передача квартир с отделкой",
    "переданы квартиры с отделкой",
    "отделка квартиры",
)


def build(spec_name: str) -> build_grp.Build:
    project = json.loads((ROOT / "tests" / spec_name).read_text(encoding="utf-8"))
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
    return b


def names(b: build_grp.Build) -> list[str]:
    return [r["name"].lower() for r in b.rows]


def test_pri_dole_nol_vetka_kvartirnoy_otdelki_snyata():
    b = build("project_69.json")
    found = [n for n in names(b) if any(m in n for m in MARKERS)]
    assert found == [], f"остались строки ветки отделки квартир: {found[:10]}"


def test_pri_dole_nol_peredacha_bez_otdelki_sohranena():
    b = build("project_69.json")
    assert any(n == "передача квартир без отделки" for n in names(b))
    assert any(n == "переданы квартиры без отделки" for n in names(b))


def test_pri_dole_nol_mokap_lobbi_sohranen_i_ne_ostalsya_bez_svyazey():
    b = build("project_69.json")
    i = b.one("Разработка рабочей документации на АИ МОКАП отделки лобби")
    assert i is not None, "блок МОКАП отделки лобби удалён вместе с типовым этажом"
    assert b.rows[i]["links"], "МОКАП лобби остался без предшественников"


def test_pri_dole_nol_vekha_zaversheny_otdelochnye_raboty_sohranena():
    b = build("project_69.json")
    i = b.one("Завершены отделочные работы")
    assert i is not None
    assert b.rows[i]["links"], "веха потеряла все связи вместе с квартирной"


def test_v_etalone_vetka_kvartirnoy_otdelki_na_meste():
    b = build("etalon_project.json")
    found = [n for n in names(b) if any(m in n for m in MARKERS)]
    assert found, "при доле 1,0 ветка отделки квартир обязана остаться"
```

- [ ] **Step 2: Прогнать тест — убедиться, что падает**

```bash
python -m pytest tests/test_dec30.py -v
```

Ожидаемое: FAIL, `AttributeError: 'Build' object has no attribute 'apply_finishing_scope'`.

- [ ] **Step 3: Реализовать `apply_finishing_scope`**

Вставить в класс `Build` в `tools/build_grp.py` сразу после `drop_rows` (перед `def apply_standards`):

```python
    # ==================================================================
    # 5а. Применимость блока отделки квартир — DEC-30
    # ==================================================================
    # Наименования покорпусных корней ветки. Удаляются вместе с потомками.
    FLAT_FIT_PER_CORPUS = ("Отделочные работы Квартиры",)
    # Наименования общих строк ветки. Наименования дословные, включая
    # опечатку шаблона «Старт СМР- отделка Квартиры» (без пробела после дефиса).
    FLAT_FIT_GLOBAL = (
        "Тендер Отделка Квартиры Чистовая",
        "Разработка рабочей документации АИ МОКАП отделки типового этажа",
        "Старт ВОР - отделка Квартиры",
        "Завершение Тендера - отделка Квартиры",
        "Старт СМР- отделка Квартиры",
        "Завершены отделочные работы квартир",
        "Передача квартир с отделкой",
        "Переданы квартиры с отделкой",
    )

    def apply_finishing_scope(self) -> None:
        """DEC-30: при отсутствии чистовой отделки квартир её ветка в график не входит.

        Шаблон v2 сам помечает обе квартирные задачи «(при наличии)», то есть
        исключение предусмотрено автором. Вместе с работами снимаются их
        контрактация, РД МОКАП типового этажа, контрольные вехи и веха передачи
        квартир с отделкой. Остаётся «Передача квартир без отделки» — она уже
        удерживается связью ОН +180 дн от вехи РВЭ (standards.md §15.3).
        """
        share = self.p.get("отделка", {}).get("доля_квартир_с_чистовой", 1.0)
        if share > 0:
            self.note("DEC-30",
                      f"Доля квартир с чистовой отделкой — {share:.0%}. Ветка чистовой отделки "
                      f"квартир сохранена полностью, передача считается по РВЭ + "
                      f"{Std.HANDOVER_FIN[0]} дн ({Std.HANDOVER_FIN[1]}). При любой доле больше "
                      f"нуля состав графика полный: standards.md §15.3 задаёт только полюса 0 % "
                      f"и 100 %, промежуточные значения правилом не покрыты — открытый вопрос "
                      f"владельцу (спека §4.6).")
            return

        roots: list[int] = []
        for pref in self.corpus_prefix.values():
            for name in self.FLAT_FIT_PER_CORPUS:
                roots.extend(self.find(name, pref))
        for name in self.FLAT_FIT_GLOBAL:
            roots.extend(self.find(name))
        removed = self.drop_rows(sorted(set(roots)), "DEC-30")

        self.note("DEC-30",
                  f"Квартиры передаются без отделки (доля чистовой отделки 0). Ветка чистовой "
                  f"отделки квартир в график не включена: снято {removed} строк — работы по "
                  f"корпусам, тендер «Отделка Квартиры Чистовая», РД АИ МОКАП отделки типового "
                  f"этажа, контрольные вехи ВОР/тендера/старта СМР, веха «Завершены отделочные "
                  f"работы квартир» и обе строки «Передача квартир с отделкой». Передача квартир "
                  f"без отделки удерживается связью ОН +{Std.HANDOVER_RAW[0]} дн от вехи РВЭ "
                  f"({Std.HANDOVER_RAW[1]}). Отделка МОП, вестибюлей и раздела ТХ сохранена и "
                  f"остаётся якорем блока ЗОС (BND-ZOS-001).")
        self.why("Отделка квартир", f"блок снят, {removed} строк",
                 "DEC-30 · standards.md §15.3 · typGRP.md, карта применимости", "средняя",
                 "Шаблон помечает квартирные задачи «(при наличии)». Финиш блока отделки "
                 "определяют МОП (410 дн) и раздел ТХ (230 дн), которые длиннее снятых "
                 "квартирных 320 дн, поэтому якорь ЗОС не смещается")
```

- [ ] **Step 4: Вызвать метод в `main()`**

В `tools/build_grp.py`, в функции `main()`, после строки `b.configure_parking()` и перед `b.apply_standards()` вставить:

```python
    b.apply_finishing_scope()   # DEC-30 — до нормативов: снятым строкам нормативы не нужны
```

- [ ] **Step 5: Прогнать тест — убедиться, что проходит**

```bash
python -m pytest tests/test_dec30.py -v
```

Ожидаемое: 5 passed.

- [ ] **Step 6: Добавить проверки в `tools/validate_grp.py`**

В функции `validate()`, непосредственно перед строкой `return r.dump()` (строка 249), вставить:

Список маркеров ниже намеренно повторяет `MARKERS` из `tests/test_dec30.py` и **не выносится в общий модуль**: тест проверяет структуру `Build.rows` до сборки Excel, валидатор — готовый файл выдачи, и это два независимых рубежа. Общая константа сделала бы ошибку в списке невидимой сразу для обоих.

```python
    # --- DEC-30: согласованность ветки чистовой отделки квартир -------------
    # Список намеренно дублирует MARKERS из tests/test_dec30.py — см. план,
    # задача 3: тест и валидатор должны падать независимо друг от друга.
    FLAT_FIT_MARKERS = (
        "отделочные работы квартиры",
        "подготовка под чистовую отделку квартир",
        "чистовая отделка квартир",
        "тендер отделка квартиры чистовая",
        "мокап отделки типового этажа",
        "завершены отделочные работы квартир",
        "передача квартир с отделкой",
        "переданы квартиры с отделкой",
    )
    lowered = [t["Название задачи"].lower() for t in tasks]
    present = [m for m in FLAT_FIT_MARKERS if any(m in nm for nm in lowered)]
    r.check(len(present) in (0, len(FLAT_FIT_MARKERS)),
            "ветка чистовой отделки квартир согласована (DEC-30)",
            f"снята частично: осталось {len(present)} из {len(FLAT_FIT_MARKERS)} маркеров — "
            f"{present}")

    # --- Передача квартир без отделки = РВЭ + 180 дн (standards.md §15.3) ---
    rve = next((t for t in tasks
                if t["Название задачи"].strip().lower().startswith("рвэ по этапу")), None)
    handovers = [t for t in tasks
                 if t["Название задачи"].strip().lower() == "передача квартир без отделки"]
    hand = min(handovers, key=lambda t: t["_lvl"]) if handovers else None
    if rve and hand and rve.get("Окончание") and hand.get("Начало"):
        delta = (datetime.strptime(hand["Начало"], "%d.%m.%Y")
                 - datetime.strptime(rve["Окончание"], "%d.%m.%Y")).days
        r.check(delta == 180,
                "передача квартир без отделки = РВЭ + 180 дн (standards.md §15.3)",
                f"фактически +{delta} дн")
    else:
        r.warns.append("справочно: веха «Передача квартир без отделки» или «РВЭ по этапу» "
                       "не найдена — проверка срока передачи пропущена")
```

- [ ] **Step 7: Записать правило в `instructions/standards.md`**

В §15.3, после таблицы «Тип помещений / Передача», перед строкой про срок по ДДУ, вставить:

```markdown
**`DEC-30` · Отсутствие чистовой отделки квартир.** При доле квартир с чистовой отделкой, равной нулю, задачи чистовой отделки квартир, их контрактация и связанная РД в график **не включаются** — состав определяется картой применимости `typGRP.md`. Передача считается по строке «Квартиры без отделки»: **РВЭ + 180 дн**. Веха «Передача квартир с отделкой» не выводится. Статус: ⚠️ — строка в «Допущения» обязательна.

При доле больше нуля состав графика полный, передача — **РВЭ + 270 дн**. Промежуточные значения доли (0 < доля < 1) правилом не покрыты: настоящий раздел задаёт только полюса. До ответа владельца применяется 270 дн как более поздний срок, с отдельной строкой в «Допущения». Статус: 🔴
```

В шапке файла поднять версию с v3.1 на v3.2 и указать дату 02.08.2026.

- [ ] **Step 8: Записать карту применимости в `instructions/typGRP.md`**

Добавить новый раздел (по образцу §6.8 «карта применимости»), в котором перечислены снимаемые задачи с их наименованиями:

```markdown
## Применимость блока отделки квартир — `DEC-30`

Признак: `отделка.доля_квартир_с_чистовой` во входных данных.

| Наименование задачи шаблона | Область | При доле 0 |
| --- | --- | --- |
| `Кn. Отделочные работы Квартиры` (с потомками «Подготовка под чистовую отделку квартир», «Чистовая отделка квартир» и их строками «По договору …») | покорпусно | снимается |
| `Тендер Отделка Квартиры Чистовая` (все 8 шагов блока) | общая | снимается |
| `Разработка рабочей документации АИ МОКАП отделки типового этажа` (с потомками) | общая | снимается |
| `Старт ВОР - отделка Квартиры` | контрольная веха | снимается |
| `Завершение Тендера - отделка Квартиры` | контрольная веха | снимается |
| `Старт СМР- отделка Квартиры` | контрольная веха | снимается |
| `Завершены отделочные работы квартир` | контрольная веха | снимается |
| `Передача квартир с отделкой` (обе строки: раздел «Передача собственникам» и раздел «Срок по ДДУ») | общая | снимается |
| `Переданы квартиры с отделкой` | контрольная веха | снимается |

Сохраняются и не затрагиваются: `Разработка рабочей документации АИ квартир` — архитектурные планировки нужны независимо от отделки; `Кn. Отделочные работы МОП и вестибюли`; `Кn. Отделочные работы, раздел ТХ`; `Разработка рабочей документации АИ МОКАП отделки лобби`; тендеры `Отделка мест общего пользования+WB` и `Отделка подземного паркинга и технических помещений`; `Чистовая отделка` офиса продаж.

Каждая связь, ссылавшаяся на снятую строку, обрабатывается по правилу `CLAUDE.md` §12: при наличии у потребителя других предшественников битая связь снимается, при их отсутствии потребитель наследует предшественников снятой строки. Каждый случай выводится в лист «Обоснование».

Финиш блока отделки после снятия определяют МОП и раздел ТХ, поэтому якорь `BND-ZOS-001` (`ОН −160 дн`) не смещается.
```

В шапке файла поднять версию с v4.1 на v4.2 и указать дату 02.08.2026.

- [ ] **Step 9: Внести `DEC-30` в `CLAUDE.md`**

1. В шапке заменить `**Версия файла: 7.1** от 02.08.2026. Комплект: `typGRP.md` v4.1 · `bindings.md` v3.1 · `standards.md` v3.1.` на версию 7.2 с комплектом `typGRP.md` v4.2 · `bindings.md` v3.1 · `standards.md` v3.2.
2. В таблице §0 поставить `typGRP.md` — **v4.2**, `standards.md` — **v3.2**.
3. В таблицу §0.4 добавить строку:

```markdown
| `DEC-30` | **Применимость блока отделки квартир: при доле чистовой отделки 0 ветка снимается, передача — РВЭ + 180 дн** | `typGRP.md`, карта применимости · `standards.md` §15.3 | ⚠ **да** — всегда при доле 0 |
```

4. В §0.1, в таблицу «Класс факта → где искать», добавить строку: `| «Входит ли блок отделки квартир в график» | `typGRP.md`, карта применимости |`.

- [ ] **Step 10: Обновить сводку версий в `README.md`**

Заменить строку сводки версий на `CLAUDE.md` v7.2 · `typGRP.md` v4.2 · `bindings.md` v3.1 · `standards.md` v3.2. В разделе «Открытые позиции» добавить строку:

```markdown
| Промежуточная доля квартир с чистовой отделкой (0 < доля < 1) | ⏸ правилом не покрыта: `standards.md` §15.3 задаёт только полюса. Применяется 270 дн, строка в «Допущения» |
```

- [ ] **Step 11: Полная проверка**

```bash
python -m pytest tests -q
python tools/build_grp.py tests/etalon_project.json out/ГРП_эталон.xlsx
python tools/check_regression.py out/ГРП_эталон.xlsx; echo "regression exit=$?"
python tools/validate_grp.py out/ГРП_эталон.xlsx; echo "validate etalon exit=$?"
python tools/build_grp.py tests/project_69.json out/ГРП_69.xlsx
python tools/validate_grp.py out/ГРП_69.xlsx; echo "validate 69 exit=$?"
```

Ожидаемое: тесты зелёные; регрессия — четыре критерия ✅ и код 0; валидатор на эталоне — код 0 и проверка `DEC-30` в статусе OK (все 8 маркеров на месте); валидатор на проекте 69 — проверка `DEC-30` в статусе OK (0 маркеров) и «передача квартир без отделки = РВЭ + 180 дн» в статусе OK. Ошибки валидатора на проекте 69, относящиеся к паркингу и высотности, на этом шаге допустимы — их чинят задачи 4 и 5.

- [ ] **Step 12: Коммит**

```bash
git add tests/test_dec30.py tools/build_grp.py tools/validate_grp.py \
        instructions/typGRP.md instructions/standards.md CLAUDE.md README.md
git commit -m "Task 3: DEC-30 - primenimost bloka otdelki kvartir

Pri dole chistovoy otdelki 0 vetka otdelki kvartir v grafik ne vhodit:
raboty po korpusam, tender Otdelka Kvartiry Chistovaya, RD AI MOKAP
otdelki tipovogo etazha, kontrolnye vehi i obe stroki Peredacha kvartir
s otdelkoy. Ostaetsya Peredacha kvartir bez otdelki - RVE + 180 dn.

Pravilo zapisano v typGRP.md (karta primenimosti) i standards.md 15.3,
indeks DEC-30 v CLAUDE.md 0.4. Versii komplekta podnyaty soglasovanno:
CLAUDE.md 7.2, typGRP.md 4.2, standards.md 3.2.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 4: Встроенный паркинг — `КЛ_ПАРК`, ИТП и веха пуска тепла

При `паркинг.есть = false` метод `configure_parking()` (`tools/build_grp.py:491`) сохраняет из блока паркинга только четыре ветки — «Нулевой цикл», «Общестроительные работы», «ЦТП, ИТП с УУТЭ», «Технические помещения» — и переносит их в корпус К1. Всё остальное снимается, в том числе веха «П1. Пуск тепла в паркинге», которую `CLAUDE.md` §9 числит обязательной. Задача проверяет и, где нужно, чинит эту ветку.

**Files:**
- Create: `tests/test_project69.py`
- Modify: `tools/build_grp.py` — по результатам проверок
- Modify: `docs/superpowers/reports/task-1-diagnostics.md` — отметить закрытые пункты

**Interfaces:**
- Consumes: `Build.configure_parking()`, `Build.apply_standards()`, `Build.thermal(nodes)`, `Build.parking_prefix: str | None`, константы `Std.ITP`, `Std.TECH_ROOMS`, `Bnd.LAG_ITP`, `Bnd.LAG_TECH_ROOMS`.
- Produces: `tests/test_project69.py` — набор тестов краевых веток, дополняется задачей 5.

- [ ] **Step 1: Написать тесты `tests/test_project69.py`**

```python
# -*- coding: utf-8 -*-
"""Краевые ветки проекта «башня 69 этажей, встроенный паркинг»."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

import build_grp

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def built():
    """Собранный проект 69 этажей: строки и рассчитанная сеть."""
    project = json.loads((ROOT / "tests" / "project_69.json").read_text(encoding="utf-8"))
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
    for corpus in project["корпуса"]:
        b.rebuild_monolith(corpus)
    nodes = b.schedule()
    b.thermal(nodes)
    b.wire_zos()
    nodes = b.schedule()
    return b, nodes


def test_otdelnogo_obekta_parkinga_net(built):
    b, _ = built
    assert all(not r["name"].startswith("П1.") for r in b.rows), \
        "при встроенном паркинге отдельного объекта П1 быть не должно"


def test_itp_perenesen_v_korpus_i_poluchil_normativ(built):
    b, _ = built
    i = b.one("По договору ЦТП, ИТП с УУТЭ", "К1")
    assert i is not None, "ИТП не перенесён в корпус (DEC-06, R-05)"
    assert b.rows[i]["dur"] == build_grp.Std.ITP[0]


def test_tehpomeshcheniya_perenoseny_i_yakoryat_itp(built):
    b, _ = built
    tech = b.one("По договору Технические помещения", "К1")
    itp = b.one("По договору ЦТП, ИТП с УУТЭ", "К1")
    assert tech is not None and itp is not None
    tech_key = b.rows[tech]["key"]
    assert any(t == tech_key for t, _, _ in b.rows[itp]["links"]), \
        "ИТП обязан стартовать от отделки техпомещений (BND-VIS-006)"


def test_vekha_pusk_tepla_korpusa_est(built):
    b, _ = built
    assert b.one("Пуск тепла корпус", "К1") is not None


def test_vekha_pusk_tepla_v_parkinge_libo_snyata_libo_privyazana(built):
    """Обязательная веха §9. При встроенном паркинге объекта нет — веха либо
    отсутствует и это записано в «Допущения», либо привязана к корпусу."""
    b, _ = built
    hits = b.find("Пуск тепла в паркинге", exact=False)
    if not hits:
        assert any("паркинг" in a.text.lower() and "тепл" in a.text.lower()
                   for a in b.assumptions), \
            "веха «Пуск тепла в паркинге» снята молча — нарушение CLAUDE.md §12"
    else:
        assert b.rows[hits[0]]["links"], "веха осталась без предшественников"
```

- [ ] **Step 2: Прогнать тесты и зафиксировать фактическую картину**

```bash
python -m pytest tests/test_project69.py -v
```

Каждый упавший тест разбирается по существу: это дефект генератора или дефект ожидания. Тест правится только если ожидание противоречит файлам контекста — тогда правка сопровождается ссылкой на пункт. Подгонять ожидание под поведение кода запрещено (`CLAUDE.md` §12).

- [ ] **Step 3: Починить то, что упало**

Правки вносятся в `tools/build_grp.py`. Каждое изменение поведения обязано:
- ссылаться на пункт `bindings.md`, `standards.md` или `typGRP.md` в аргументе `source` вызова `self.why(...)`;
- либо, если правила нет, оставлять поведение как есть и фиксировать открытый вопрос через `self.note(...)` — по образцу существующей заметки про «отвал» в `apply_standards()` (`tools/build_grp.py:573`).

Для вехи «Пуск тепла в паркинге» при встроенном паркинге правила в комплекте нет. Значение по умолчанию: веха неприменима (объекта «паркинг» в проекте нет, подземные уровни входят в тепловой контур корпуса), её снятие фиксируется строкой в «Допущения» с текстом, включающим слова «паркинг» и «тепло», — иначе тест из Step 1 не пройдёт.

- [ ] **Step 4: Прогнать тесты — убедиться, что проходят**

```bash
python -m pytest tests/test_project69.py -v
```

Ожидаемое: 5 passed.

- [ ] **Step 5: Полная проверка**

```bash
python -m pytest tests -q
python tools/build_grp.py tests/etalon_project.json out/ГРП_эталон.xlsx
python tools/check_regression.py out/ГРП_эталон.xlsx; echo "regression exit=$?"
python tools/validate_grp.py out/ГРП_эталон.xlsx; echo "validate etalon exit=$?"
python tools/build_grp.py tests/project_69.json out/ГРП_69.xlsx
python tools/validate_grp.py out/ГРП_69.xlsx; echo "validate 69 exit=$?"
```

Ожидаемое: тесты зелёные, регрессия эталона в допуске, валидатор эталона — код 0.

- [ ] **Step 6: Отметить закрытые пункты в отчёте задачи 1**

В `docs/superpowers/reports/task-1-diagnostics.md`, в таблице §4, проставить в каждой закрытой строке пометку `закрыто задачей 4` и краткое «чем именно».

- [ ] **Step 7: Коммит**

```bash
git add tests/test_project69.py tools/build_grp.py docs/superpowers/reports/task-1-diagnostics.md
git commit -m "Task 4: vstroennyy parking - KL_PARK, ITP i vekha puska tepla

Pri parking.est=false nulevoy cikl, ITP i tehpomeshcheniya perenosyatsya
v korpus K1. Proverena privyazka ITP k otdelke tehpomeshcheniy i sudba
obyazatelnoy vehi Pusk tepla v parkinge - ona ne mozhet ischezat molcha
(CLAUDE.md 12).

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 5: Высотность 69 этажей

Эталон содержит корпуса на 26 и 55 этажей. Задача проверяет, что правила, откалиброванные на них, дают осмысленный результат на 69 этажах: монолит при `K = 7`, перепривязки `min(6, N)` и `min(15, N)`, тепловой контур, окно отделки против жёсткого минимума 210 дн и финиш ВИС `ЯКОРЬ + 330`.

**Files:**
- Modify: `tests/test_project69.py` — дописать тесты высотности
- Modify: `tools/build_grp.py` и/или `tools/grp_model.py` — по результатам проверок
- Modify: `docs/superpowers/reports/task-1-diagnostics.md`

**Interfaces:**
- Consumes: фикстура `built` из `tests/test_project69.py` (задача 4); `build_grp.monolith_floor_durations(n_floors, complex_frame) -> tuple[int, list[int], int]` — возвращает `(длительность первого этажа, список длительностей остальных, длительность кровли)`; `Std.FIT_MIN`, `Bnd.MASONRY_FLOOR`, `Bnd.VIS_FLOOR`.
- Produces: закрытые проверки высотности; выводы попадают в лист «Обоснование» выдачи.

- [ ] **Step 1: Дописать тесты в `tests/test_project69.py`**

Добавить в конец файла:

```python
# --- высотность 69 этажей ------------------------------------------------

def test_monolit_razvernut_po_vsem_69_etazham(built):
    b, _ = built
    floors = [r for r in b.rows
              if r["name"].startswith("К1. ") and r["name"].endswith("этаж Монолит")]
    assert len(floors) == 69, f"этажей монолита {len(floors)}, ожидалось 69"


def test_koefficient_K_raven_7_pri_69_etazhah(built):
    first, typical, roof = build_grp.monolith_floor_durations(69, True)
    assert typical and typical[0] == 7, \
        "при сложном конструктиве и N > 60 типовой этаж — 7 дн (standards.md §9)"


def test_kladka_privyazana_k_vekhe_6_etazha_a_ne_k_69(built):
    b, _ = built
    assert b.one("К1. 6 этаж Монолит") is not None, \
        "веха min(6, N) = 6 этаж отсутствует — перепривязка BND-KLD-001 невозможна"


def test_vis_privyazany_k_vekhe_15_etazha(built):
    b, _ = built
    assert b.one("К1. 15 этаж Монолит") is not None, \
        "веха min(15, N) = 15 этаж отсутствует — перепривязка BND-VIS-001 невозможна"


def test_okno_otdelki_ne_nizhe_zhestkogo_minimuma(built):
    b, nodes = built
    fits = [k for k, _ in b.fit_tasks]
    assert fits, "задачи отделки не зарегистрированы в fit_tasks"
    for k in fits:
        n = nodes[k]
        assert n.start and n.finish
        assert (n.finish - n.start).days >= build_grp.Std.FIT_MIN[0], \
            f"окно отделки {(n.finish - n.start).days} дн меньше минимума " \
            f"{build_grp.Std.FIT_MIN[0]} дн — запрещено standards.md §15.1"


def test_kriticheskiy_put_nepreryven(built):
    b, nodes = built
    crit = [n for n in nodes.values() if n.critical]
    assert crit, "критический путь пуст"


def test_pri_oknah_pvh_kladka_naruzhnyh_sten_prisutstvuet(built):
    """Витражей нет → BND-KLD-001 применяется, ЯКОРЬ_ОГР = кладка наружных стен."""
    b, _ = built
    assert b.one("По договору Кладка наружных стен", "К1") is not None, \
        "при остеклении ПВХ наружная кладка обязана быть — она задаёт ЯКОРЬ_ОГР"


def test_porog_DEC_14_po_svayam_otrabotan(built):
    """90 свай БНС на одну установку: превышение порога даёт уведомление, а не тишину."""
    b, _ = built
    days, _ = build_grp.pile_duration(
        [{"тип": "БНС", "количество": 90, "установок": 1}])
    if days > build_grp.Std.PILE_WARN_DAYS[0]:
        assert any("DEC-14" == a.code for a in b.assumptions), \
            "порог DEC-14 превышен, но строки в «Допущения» нет"
        assert b.notices, "порог DEC-14 превышен, но уведомления пользователю нет"
```

- [ ] **Step 2: Прогнать и зафиксировать картину**

```bash
python -m pytest tests/test_project69.py -v
```

- [ ] **Step 3: Починить то, что упало**

Те же требования, что в задаче 4 Step 3: каждое изменение поведения ссылается на пункт файла контекста либо оставляет поведение и фиксирует открытый вопрос через `self.note(...)`.

Отдельно проверить и отразить в листе «Обоснование»: если окно отделки при 69 этажах оказалось меньше 210 дн, порядок мер жёсткий — сначала ВТК, только потом сдвиг финиша (`CLAUDE.md` §6 п. 7), с уведомлением и строкой в «Допущения» по `DEC-02` и `DEC-09`. Сжатие ниже 210 дн запрещено при любом дедлайне.

- [ ] **Step 4: Прогнать тесты — убедиться, что проходят**

```bash
python -m pytest tests/test_project69.py -v
```

Ожидаемое: 13 passed (5 из задачи 4 + 8 новых).

- [ ] **Step 5: Полная проверка**

```bash
python -m pytest tests -q
python tools/build_grp.py tests/etalon_project.json out/ГРП_эталон.xlsx
python tools/check_regression.py out/ГРП_эталон.xlsx; echo "regression exit=$?"
python tools/validate_grp.py out/ГРП_эталон.xlsx; echo "validate etalon exit=$?"
python tools/build_grp.py tests/project_69.json out/ГРП_69.xlsx
python tools/validate_grp.py out/ГРП_69.xlsx; echo "validate 69 exit=$?"
```

- [ ] **Step 6: Отметить закрытые пункты в отчёте задачи 1**

- [ ] **Step 7: Коммит**

```bash
git add tests/test_project69.py tools/build_grp.py tools/grp_model.py \
        docs/superpowers/reports/task-1-diagnostics.md
git commit -m "Task 5: vysotnost 69 etazhey

Provereny monolit pri K=7, perepryvyazki min(6,N) i min(15,N), okno
otdelki protiv zhestkogo minimuma 210 dn, YAKOR_OGR po kladke naruzhnyh
sten pri oknah PVH i porog DEC-14 po svayam. Poryadok mer pri nehvatke
okna: snachala VTK, potom sdvig (CLAUDE.md 6 p.7).

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 6: Остаточные поломки

Задача забирает всё из §5 отчёта `docs/superpowers/reports/task-1-diagnostics.md` — «поломки, не покрытые задачами 2–5». Если раздел пуст, задача закрывается без изменений кода: это допустимый исход, а не повод искать работу.

**Files:**
- Modify: `tools/build_grp.py`, `tools/grp_model.py`, `tools/validate_grp.py` — по списку
- Modify: `tests/test_project69.py` — тест на каждую починенную поломку
- Modify: `docs/superpowers/reports/task-1-diagnostics.md`

**Interfaces:**
- Consumes: §5 отчёта задачи 1; фикстура `built` из `tests/test_project69.py`.
- Produces: `python tools/validate_grp.py out/ГРП_69.xlsx` завершается с кодом 0.

- [ ] **Step 1: Прочитать §5 отчёта и составить список**

```bash
cat docs/superpowers/reports/task-1-diagnostics.md
python tools/build_grp.py tests/project_69.json out/ГРП_69.xlsx
python tools/validate_grp.py out/ГРП_69.xlsx
```

Список — это пересечение §5 отчёта и текущих строк «ОШИБКА» валидатора. Если и то и другое пусто — перейти к Step 5.

- [ ] **Step 2: На каждую поломку — падающий тест**

Тест дописывается в `tests/test_project69.py` и использует фикстуру `built`. Формулировка утверждения — от требования файла контекста, а не от текущего поведения кода.

- [ ] **Step 3: Прогнать — убедиться, что падают**

```bash
python -m pytest tests/test_project69.py -v
```

- [ ] **Step 4: Починить и убедиться, что проходят**

```bash
python -m pytest tests/test_project69.py -v
```

Требование к каждой правке то же: ссылка на пункт файла контекста в `self.why(...)` либо открытый вопрос через `self.note(...)`. Значения «по опыту» запрещены (`CLAUDE.md` §0.3).

- [ ] **Step 5: Полная проверка**

```bash
python -m pytest tests -q
python tools/build_grp.py tests/etalon_project.json out/ГРП_эталон.xlsx
python tools/check_regression.py out/ГРП_эталон.xlsx; echo "regression exit=$?"
python tools/validate_grp.py out/ГРП_эталон.xlsx; echo "validate etalon exit=$?"
python tools/build_grp.py tests/project_69.json out/ГРП_69.xlsx
python tools/validate_grp.py out/ГРП_69.xlsx; echo "validate 69 exit=$?"
```

Ожидаемое: **все четыре кода возврата 0**. Это первая задача, где валидатор на проекте 69 обязан быть чистым.

- [ ] **Step 6: Закрыть отчёт**

В `docs/superpowers/reports/task-1-diagnostics.md` в §4 и §5 не должно остаться строк без пометки: либо `закрыто задачей N`, либо `открытый вопрос владельцу — строка в «Допущения»`.

- [ ] **Step 7: Коммит**

```bash
git add tools tests docs/superpowers/reports/task-1-diagnostics.md
git commit -m "Task 6: ostatochnye polomki progona 69 etazhey

Zakryty punkty otcheta zadachi 1, ne pokrytye zadachami 2-5.
validate_grp na proekte 69 - 0 oshibok.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 7: Итоговая выдача

Задача не меняет логику расчёта. Её результат — сама выдача и обновлённое состояние документации репозитория.

**Files:**
- Create: `out/ГРП_69.xlsx` (не коммитится — каталог `out/` рабочий)
- Create: `docs/superpowers/reports/task-7-vydacha.md`
- Modify: `README.md` — раздел «Состояние генератора»

**Interfaces:**
- Consumes: весь конвейер после задач 2–6.
- Produces: `docs/superpowers/reports/task-7-vydacha.md` — сводка выдачи для планировщика.

- [ ] **Step 1: Собрать выдачу**

```bash
python tools/build_grp.py tests/project_69.json out/ГРП_69.xlsx
python tools/validate_grp.py out/ГРП_69.xlsx; echo "exit=$?"
```

Ожидаемое: код возврата 0.

- [ ] **Step 2: Снять ключевые даты из выдачи**

```bash
python -c "
import sys
sys.stdout.reconfigure(errors='replace')
sys.path.insert(0, 'tools')
from validate_grp import load
from pathlib import Path
tasks = load(Path('out/ГРП_69.xlsx'))
keys = ('разрешение на строительство', 'рвэ по этапу', 'получено рвэ',
        'передача квартир без отделки', 'пуск тепла', 'зос')
for t in tasks:
    nm = t['Название задачи'].lower()
    if any(k in nm for k in keys):
        print(f\"{t['Ид.']:>5} {t['Название задачи'][:56]:56} {t['Начало']:>12} {t['Окончание']:>12}\")
"
```

- [ ] **Step 3: Проверить листы книги**

```bash
python -c "
import openpyxl
wb = openpyxl.load_workbook('out/ГРП_69.xlsx', read_only=True)
print(wb.sheetnames)
for s in wb.sheetnames:
    print(s, wb[s].max_row)
"
```

Ожидаемое: пять листов — `ГРП`, `Обоснование`, `Допущения`, `Чек-лист`, `Сценарии` (`CLAUDE.md` §3.1). Лист «Допущения» непустой и содержит строки по `DEC-30`, `DEC-16`, умолчаниям `standards.md` §3 и неподтверждённым индивидуальным условиям.

- [ ] **Step 4: Написать сводку `docs/superpowers/reports/task-7-vydacha.md`**

```markdown
# Выдача ГРП: башня 69 этажей, без отделки квартир

Вход: `tests/project_69.json` · Файл: `out/ГРП_69.xlsx` · Коммит: <git rev-parse --short HEAD>

## Ключевые даты

| Веха | Дата |
| --- | --- |
| Разрешение на строительство (РС) получено | ... |
| Пуск тепла корпус К1 | ... |
| Получен ЗОС | ... |
| РВЭ по этапу 1 получен | ... |
| Передача квартир без отделки | ... |

Фаза A: ... дн · Фаза РС → РВЭ: ... дн

## Состав выдачи

Строк: ... · вех: ... · обоснований: ... · допущений: ...
Листы: ГРП · Обоснование · Допущения · Чек-лист · Сценарии

## Проверки

| Проверка | Результат |
| --- | --- |
| `validate_grp.py out/ГРП_69.xlsx` | код 0, ошибок 0 |
| `validate_grp.py out/ГРП_эталон.xlsx` | код 0, ошибок 0 |
| `check_regression.py out/ГРП_эталон.xlsx` | четыре критерия §10 в допуске |
| `pytest tests` | ... passed |

## Открытые вопросы владельцу

<строки листа «Допущения», требующие решения: DEC-30 при промежуточной доле,
веха пуска тепла в паркинге при встроенном паркинге, прочее из задач 4–6>
```

- [ ] **Step 5: Обновить `README.md`**

В разделе «Состояние генератора» добавить строку о проверенных краевых ветках: одна башня 69 этажей сложного конструктива со встроенным паркингом, без отделки квартир — собирается и валидируется без ошибок. Таблицу критериев приёмки §10 не трогать, если регрессия не изменилась; если отклонения изменились — обновить числа фактическим выводом `check_regression.py`.

- [ ] **Step 6: Финальная проверка целиком**

```bash
python -m pytest tests -q
python tools/build_grp.py tests/etalon_project.json out/ГРП_эталон.xlsx
python tools/check_regression.py out/ГРП_эталон.xlsx; echo "regression exit=$?"
python tools/validate_grp.py out/ГРП_эталон.xlsx; echo "validate etalon exit=$?"
python tools/build_grp.py tests/project_69.json out/ГРП_69.xlsx
python tools/validate_grp.py out/ГРП_69.xlsx; echo "validate 69 exit=$?"
```

Ожидаемое: все коды возврата 0, тесты зелёные, четыре критерия §10 ✅.

- [ ] **Step 7: Коммит**

```bash
git add docs/superpowers/reports/task-7-vydacha.md README.md
git commit -m "Task 7: itogovaya vydacha GRP dlya proekta 69 etazhey

Svodka klyuchevyh dat, sostava vydachi i otkrytyh voprosov vladeltsu.
README obnovlen: proverennye kraevye vetki.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Что план сознательно не делает

- Не трогает `instructions/bindings.md`, если задачи 4–5 не найдут в нём реального расхождения. Правка файла контекста ради удобства реализации запрещена (Global Constraint 7).
- Не реализует поддержку двух и более этапов ввода: `check_stages()` останавливает расчёт намеренно, правил владельцем не задано (`typGRP.md` §14 п. 7).
- Не считает сценарии и вероятностный расчёт (`DEC-16`).
- Не заполняет колонки «Вид работ» и «Код классификатора»: справочник МДМ не поступил, изобретать коды запрещено (`CLAUDE.md` §12).
- Не коммитит содержимое `out/` — это рабочий каталог.
