"""Проверка parse.py на реальном документе: 116 блоков, 0 таблиц, все Normal.

Ф9: документ размечен ПРЯМЫМ форматированием, а не именованными стилями —
`style` остаётся "Normal" везде, но `level`/`list`, читаемые из pPr напрямую,
обязаны вскрыть реальную структуру (24 заголовка, 31 элемент списка, из них
7 вложенных)."""

from collections import Counter

from docx import Document

from docx_editor.parse import doc_map, index, render

DOC = "/home/artem/Загрузки/Архитектура_ColBERT.docx"


def main():
    doc = Document(DOC)
    idx = index(doc)
    blocks = doc_map(doc, idx)

    assert len(blocks) == 116, f"ожидали 116 блоков, получили {len(blocks)}"
    tables = [b for b in blocks if b["kind"] == "t"]
    assert len(tables) == 0, f"ожидали 0 таблиц, получили {len(tables)}"
    paragraphs = [b for b in blocks if b["kind"] == "p"]
    styles = {b["style"] for b in paragraphs}
    assert styles == {"Normal"}, f"ожидали только стиль Normal, получили {styles}"

    with_level = [b for b in paragraphs if b["level"] is not None]
    assert len(with_level) == 24, f"ожидали 24 блока с level, получили {len(with_level)}"
    by_level = Counter(b["level"] for b in with_level)
    assert by_level == {0: 1, 1: 9, 2: 14}, f"неожиданное распределение по level: {by_level}"

    with_list = [b for b in paragraphs if b["list"] is not None]
    assert len(with_list) == 31, f"ожидали 31 блок с list, получили {len(with_list)}"
    nested = [b for b in with_list if b["list"]["ilvl"] == 1]
    assert len(nested) == 7, f"ожидали 7 вложенных (ilvl=1), получили {len(nested)}"

    text = render(blocks)
    assert text.startswith("p0 [Normal]")

    print(f"parse_demo: ok, {len(blocks)} блоков, стили {styles}, "
          f"level: {dict(by_level)}, list: {len(with_list)} (вложенных {len(nested)})")


if __name__ == "__main__":
    main()
