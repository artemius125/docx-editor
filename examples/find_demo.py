"""Приёмка find.py: 20 якорей (по одному на правку из Правки_ColBERT_20.md)
на настоящем документе Архитектура_ColBERT.docx, плюс outline и fragment.
"""

from docx import Document

from docx_editor.parse import doc_map, index
from docx_editor import find

REAL_DOC = "/home/artem/Загрузки/Архитектура_ColBERT.docx"

# Якорь — самая специфичная ДОСЛОВНАЯ цитата из текста самой правки (не термин
# своими словами). Для правок 6, 9, 15 такой цитаты в тексте правки нет вообще —
# это не промах find.py, а находка (см. отчёт строителя): правка 6 — глобальное
# типографическое правило (идёт мимо Редактора), правка 9 — правка по позиции
# («заголовок» = первый блок, а не текст), правка 15 — уже известный невыполнимый
# случай (BUILD_PLAN «Известные риски»: в документе нет списков).
ANCHORS = {
    1: "структурные маркеры как",
    2: "исчисляемым секундами",
    3: "меняет этот процесс.Исследователи",
    4: "0.4436",
    5: "MS-MARCO",
    6: None,
    7: "Single-Vector Bi-encoders",
    8: "2025/2026 годов",
    9: None,
    10: "Модернизация агрегации",
    11: "Кросс-энкодер",
    12: "на порядок (в 6–10 раз)",
    13: "коллапс производительности стандартного ColBERT (падение до 86–97%)",
    14: "MRR@10",
    15: None,
    16: "матричных умножений MaxSim",
    17: "представляет собой скалярное произведение",
    18: "недифференцируемой природе операции",
    19: "FLOPs",
    20: "MetaEmbed",
}


def _load():
    doc = Document(REAL_DOC)
    idx = index(doc)
    return doc_map(doc, idx)


def test_anchors_hit_every_edit():
    blocks = _load()
    no_anchor = []
    for n, needle in ANCHORS.items():
        if needle is None:
            no_anchor.append(n)
            continue
        hits = find.by_text(blocks, needle)
        assert hits, f"правка {n}: якорь {needle!r} не нашёл ни одного блока"
        assert len(hits) <= 20, f"правка {n}: якорь {needle!r} слишком общий: {len(hits)} блоков"

    # список правок без цитаты зафиксирован явно: если он изменится — это
    # сигнал, что мой выбор якоря был плохим, а не новая находка
    assert no_anchor == [6, 9, 15], f"неожиданный набор правок без якоря: {no_anchor}"
    print(f"find_demo: {20 - len(no_anchor)}/20 якорей нашли блоки (≤20 каждый), "
          f"{len(no_anchor)} без цитаты в тексте правки")


def test_first_mention_picks_earliest():
    blocks = _load()
    # MRR@10 упоминается дважды (p48 и p77) — first_mention обязан вернуть первый
    hits = find.by_text(blocks, "MRR@10")
    assert len(hits) >= 2, "ожидали повторный термин для проверки first_mention"
    assert find.first_mention(blocks, "MRR@10") == "p48"
    assert find.first_mention(blocks, "текста, которого точно нет") is None
    print("find_demo: first_mention возвращает самое раннее упоминание")


def test_by_regex():
    blocks = _load()
    hits = find.by_regex(blocks, r"\d\.\d")
    assert "p67" in hits, "в p67 три десятичных числа (4.2%, 0.4448, 0.4436)"
    print("find_demo: by_regex находит десятичные числа")


def test_outline_is_compact_and_valid():
    blocks = _load()
    o = find.outline(blocks)
    assert len(o) < 3200, f"оглавление слишком длинное для промпта Навигатора: {len(o)} знаков"
    ids = {line.split(" ", 1)[0] for line in o.splitlines()}
    real_ids = {b["id"] for b in blocks}
    assert ids <= real_ids, f"в оглавлении лишние id: {ids - real_ids}"
    assert len(ids) == len(blocks), "оглавление пропустило блоки"
    print(f"find_demo: outline — {len(o)} знаков на {len(blocks)} блоков, все id существуют")


def test_fragment_dedupes_overlap():
    blocks = _load()
    # p10 и p12 с around=1 дают пересекающиеся диапазоны [p9,p11] и [p11,p13]
    frag = find.fragment(blocks, ["p10", "p12"], around=1)
    ids = [b["id"] for b in frag]
    assert ids == ["p9", "p10", "p11", "p12", "p13"], ids
    print("find_demo: fragment схлопывает пересекающиеся диапазоны без дублей")


if __name__ == "__main__":
    test_anchors_hit_every_edit()
    test_first_mention_picks_earliest()
    test_by_regex()
    test_outline_is_compact_and_valid()
    test_fragment_dedupes_overlap()
