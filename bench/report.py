"""Отчёт по эффективности: одна метрика на прогон и график её роста.

Э = (правки, подтверждённые оценщиком по файлу − вердикты, оказавшиеся ложью)
    / число проверяемых правок × 100%

Числитель честный по построению: подтверждает ФАЙЛ (oracle.py), а ложный
вердикт вычитается, поэтому «сделать больше, но соврать» метрику не поднимает.
Знаменатель — ВСЕ правки, на которые у оценщика есть маркер (остальные он не
умеет судить, и приписывать их в успех нельзя). Именно все, а не дошедшие до
конца: иначе оборванный прогон сужает знаменатель и выглядит лучше целого.

Запуск: python3 report.py  → перезаписывает ЭФФЕКТИВНОСТЬ.md
"""

from pathlib import Path

from oracle import CHECKS
from score import load, score_one

HERE = Path(__file__).resolve().parent
OUT = HERE / "ЭФФЕКТИВНОСТЬ.md"
CHECKABLE = sum(1 for v in CHECKS.values() for check in v.values() if check)


def efficiency(label, seq):
    """(Э в процентах, подтверждено, ложь, дошло до оценщика) или None без данных."""
    try:
        rows = [score_one(r) for r in load(label, seq)]
    except SystemExit:
        return None
    reached = sum(1 for r in rows if r["oracle"] != "n/a")
    ok = sum(1 for r in rows if r["oracle"] == "pass")
    lies = sum(1 for r in rows if not r["honest"])
    return round((ok - lies) / CHECKABLE * 100), ok, lies, reached


def labels():
    """Прогоны в хронологическом порядке — по времени ПЕРВОЙ записи в каталоге
    (по последней нельзя: перенос каталогов переставил им mtime скопом).

    Каталоги с одним корпусом в график не идут: это повторные прогоны для
    проверки разброса (w9b — ColBERT трижды подряд), и рядом с полными
    точками они читались бы как провал, хотя меряли другое."""
    dirs = [p for p in (HERE / "runs").iterdir() if p.is_dir()]
    full, partial = [], []
    for d in sorted(dirs, key=lambda p: min(f.stat().st_mtime for f in p.iterdir())):
        names = {f.name for f in d.iterdir()}
        (full if {"colbert.jsonl", "math.jsonl"} <= names else partial).append(d.name)
    return full, partial


def chart(title, points):
    """Столбики: подпись, полоса длиной в Э, число и его расшифровка."""
    lines = [f"### {title}", "", "```"]
    for label, e in points:
        if e is None:
            lines.append(f"{label:9s} {'—':52s} нет прогона")
        else:
            pct, ok, lies, reached = e
            bar = "█" * round(pct / 2)
            note = f"{ok} подтверждено" + (f", −{lies} ложь" if lies else "")
            if reached < CHECKABLE:
                note += f"; ПРОГОН ОБОРВАН, оценщик увидел {reached} из {CHECKABLE}"
            lines.append(f"{label:9s} {bar:52s} {pct:3d}%  ({note})")
    lines += ["```", ""]
    return lines


def main():
    names, partial = labels()
    text = [
        "# Эффективность", "",
        "Одна метрика: **Э = (подтверждено оценщиком − ложных вердиктов) /",
        f"{CHECKABLE} проверяемых правок**. Считает файл, а не вердикт системы; ложь",
        "вычитается. Корпус — 40 правок (ColBERT 20 + Математика 20), из них маркеры",
        f"оценщика есть у {CHECKABLE}; знаменатель постоянный, оборванный прогон себе не льстит.",
        "",
        "Точка на графике — партия изменений кода, замеренная целиком.",
        "Разброс между прогонами ±2 правки: рост меньше 3 пунктов ничего не доказывает.",
        "",
        "Отчёт пересоздаётся: `python3 report.py` (офлайн, читает `runs/`).",
        "",
    ]
    text += chart("Каждая правка на свежей копии — сигнал по механизму",
                  [(n, efficiency(n, seq=False)) for n in names])
    text += chart("Весь список подряд по одному документу — правда о продукте",
                  [(n, efficiency(n, seq=True)) for n in names])
    if partial:
        text += [f"Не в графике (прогон по одному корпусу, проверка разброса): {', '.join(partial)}.", ""]
    OUT.write_text("\n".join(text), encoding="utf-8")
    print(OUT)


if __name__ == "__main__":
    main()
