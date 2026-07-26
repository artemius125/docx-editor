"""Проверка parse.py на реальном документе: 116 блоков, 0 таблиц, все Normal."""

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

    text = render(blocks)
    assert text.startswith("p0 [Normal]")

    print(f"parse_demo: ok, {len(blocks)} блоков, стили {styles}")


if __name__ == "__main__":
    main()
