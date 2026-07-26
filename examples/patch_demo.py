"""Приёмка patch.py: девять операций на маленьком документе + валидатор невалидных патчей."""

from docx import Document
from docx.oxml.ns import qn

from docx_editor.parse import doc_map, index
from docx_editor.patch import apply, validate

REAL_DOC = "/home/artem/Загрузки/Архитектура_ColBERT.docx"


def _build():
    doc = Document()
    doc.add_paragraph("Первый абзац.")
    doc.add_paragraph("Второй абзац для удаления.")
    doc.add_paragraph("Третий абзац.")
    p = doc.add_paragraph()
    p.add_run("Срок ")
    p.add_run("3").bold = True
    p.add_run("0 дн").italic = True
    r = p.add_run("ей")
    r.bold, r.underline = True, True
    p.add_run(" истекает.")
    return doc


def test_ops():
    # idx создаётся один раз и переиспользуется: apply дописывает в него новые
    # элементы (insert_after/create_table), повторный index(doc) звал бы
    # позиционную перенумерацию и разрушил бы стабильность id.
    doc = _build()
    idx = index(doc)
    blocks = doc_map(doc, idx)

    op = {"op": "replace_text", "id": "p0", "old": "Первый", "new": "Изменённый"}
    assert validate(blocks, op, doc) is None
    apply(doc, idx, op)
    blocks = doc_map(doc, idx)
    assert blocks[0]["text"] == "Изменённый абзац."

    # фраза лежит на стыке трёх ранов разного начертания — соседи не тронуты
    apply(doc, idx, {"op": "replace_text", "id": "p3", "old": "30 дней", "new": "60 дней"})
    blocks = doc_map(doc, idx)
    assert blocks[3]["text"] == "Срок 60 дней истекает."
    runs = doc.paragraphs[3].runs
    assert runs[0].text == "Срок " and not runs[0].bold
    assert runs[-1].text == " истекает." and not runs[-1].bold

    apply(doc, idx, {"op": "set_text", "id": "p2", "text": "Совсем другой текст."})
    blocks = doc_map(doc, idx)
    assert blocks[2]["text"] == "Совсем другой текст."

    apply(doc, idx, {"op": "insert_after", "id": "p2", "text": "Новый абзац.", "style": "Normal"})
    blocks = doc_map(doc, idx)
    assert blocks[3]["text"] == "Новый абзац."

    apply(doc, idx, {"op": "set_style", "id": "p0", "style": "Heading 1"})
    blocks = doc_map(doc, idx)
    assert blocks[0]["style"] == "Heading 1"

    before = len(blocks)
    apply(doc, idx, {"op": "delete", "id": "p1"})
    blocks = doc_map(doc, idx)
    assert len(blocks) == before - 1 and all(b["id"] != "p1" for b in blocks)

    last_id = blocks[-1]["id"]
    apply(doc, idx, {"op": "move_after", "id": "p0", "after": last_id})
    blocks = doc_map(doc, idx)
    assert blocks[-1]["id"] == "p0"

    apply(doc, idx, {"op": "create_table", "after": "p0", "rows": [["a", "b"], ["c", "d"]], "header": True})
    blocks = doc_map(doc, idx)
    table = [b for b in blocks if b["kind"] == "t"][0]
    assert table["rows"] == [["a", "b"], ["c", "d"]]

    apply(doc, idx, {"op": "set_cell", "id": table["id"], "row": 1, "col": 1, "text": "абзац в ячейке"})
    blocks = doc_map(doc, idx)
    table = [b for b in blocks if b["kind"] == "t"][0]
    assert table["rows"][1][1] == "абзац в ячейке"

    apply(doc, idx, {"op": "replace_all", "old": "абзац", "new": "блок"})
    blocks = doc_map(doc, idx)
    table = [b for b in blocks if b["kind"] == "t"][0]
    assert table["rows"][1][1] == "блок в ячейке"
    assert not any("абзац" in b["text"] for b in blocks if b["kind"] == "p")

    print("patch_demo: девять операций применены и проверены")


def test_invalid():
    doc = _build()
    idx = index(doc)
    blocks = doc_map(doc, idx)

    assert isinstance(validate(blocks, {"op": "frobnicate"}, doc), str)
    assert isinstance(validate(blocks, {"op": "replace_text", "id": "p99", "old": "a", "new": "b"}, doc), str)
    assert isinstance(validate(blocks, {"op": "move_after", "id": "p0", "after": "p99"}, doc), str)
    assert isinstance(validate(blocks, {"op": "replace_text", "id": "p0", "old": "Первый", "new": "Первый"}, doc), str)
    assert isinstance(validate(blocks, {"op": "replace_all", "old": "абзац", "new": "абзац"}, doc), str)
    assert isinstance(validate(blocks, {"op": "replace_text", "id": "p0", "old": "нет такого текста", "new": "x"}, doc), str)

    err = validate(blocks, {"op": "set_style", "id": "p0", "style": "НесуществующийСтиль"}, doc)
    assert isinstance(err, str) and "Normal" in err

    # insert_after с несуществующим стилем — тот же класс ошибки, что и set_style
    err_ins = validate(blocks, {"op": "insert_after", "id": "p0", "text": "x", "style": "НесуществующийСтиль"}, doc)
    assert isinstance(err_ins, str) and "Normal" in err_ins

    apply(doc, idx, {"op": "create_table", "after": "p0", "rows": [["a", "b"]]})
    blocks = doc_map(doc, idx)
    table_id = [b["id"] for b in blocks if b["kind"] == "t"][0]
    assert isinstance(validate(blocks, {"op": "set_cell", "id": table_id, "row": 5, "col": 0, "text": "x"}, doc), str)

    assert validate(blocks, {"op": "replace_text", "id": "p0", "old": "Первый", "new": "Другой"}, doc) is None

    print("patch_demo: невалидные патчи отбиты валидатором")


def test_hyperlink():
    # текст гиперссылки лежит в w:hyperlink/w:r — не в r_lst абзаца верхнего
    # уровня. p.text (то, что видит validate) его включает, поэтому apply
    # обязан ходить по тем же ранам, иначе offset'ы разъезжаются и
    # _replace_span молча портит абзац (см. находку в BUILD_PLAN).
    doc = Document()
    p = doc.add_paragraph("см. ")
    hl = p._p.makeelement(qn("w:hyperlink"), {})
    r = p._p.makeelement(qn("w:r"), {})
    t = p._p.makeelement(qn("w:t"), {})
    t.text = "документацию"
    r.append(t)
    hl.append(r)
    p._p.append(hl)

    idx = index(doc)
    blocks = doc_map(doc, idx)
    assert blocks[0]["text"] == "см. документацию"

    op = {"op": "replace_text", "id": "p0", "old": "документацию", "new": "инструкцию"}
    assert validate(blocks, op, doc) is None
    apply(doc, idx, op)
    blocks = doc_map(doc, idx)
    assert blocks[0]["text"] == "см. инструкцию", blocks[0]["text"]

    print("patch_demo: replace_text через w:hyperlink не портит абзац")


def test_create_table_has_borders():
    doc = _build()
    idx = index(doc)
    apply(doc, idx, {"op": "create_table", "after": "p0", "rows": [["a", "b"], ["c", "d"]]})
    tbl = next(el for tid, el in idx.items() if tid.startswith("t"))
    assert tbl.tblPr.find(qn("w:tblBorders")) is not None, "нет w:tblBorders — таблица без рамок"
    print("patch_demo: create_table ставит рамки прямо в XML")


def test_real_doc_style_error():
    doc = Document(REAL_DOC)
    idx = index(doc)
    blocks = doc_map(doc, idx)
    err = validate(blocks, {"op": "set_style", "id": "p0", "style": "List Bullet"}, doc)
    assert isinstance(err, str), "ожидали отказ: List Bullet отсутствует в документе"
    available = err.split("доступны стили абзацев:")[1]
    assert "List Bullet" not in available, err
    assert "List Paragraph" in available, err
    print("patch_demo: set_style(List Bullet) на реальном документе честно отказан")


def test_replace_text_tolerates_nbsp():
    # документ реально содержит U+00A0 посреди предложений (находка Ф4: 61 из
    # 116 абзацев, 147 вхождений), а цитата в патче от модели набрана обычным
    # пробелом. old написан обычным пробелом, в абзаце — U+00A0.
    doc = Document()
    doc.add_paragraph("слово\xa0слово синее небо.")
    idx = index(doc)
    blocks = doc_map(doc, idx)

    op = {"op": "replace_text", "id": "p0", "old": "слово слово", "new": "два слова"}
    assert validate(blocks, op, doc) is None
    apply(doc, idx, op)
    blocks = doc_map(doc, idx)
    # результат целиком, не только начало: неверный end() обрежет или откусит
    # лишнее от хвоста абзаца — падать должно громко, а не тихо на префиксе
    assert blocks[0]["text"] == "два слова синее небо.", blocks[0]["text"]

    # точное совпадение (быстрый путь) не сломано
    op2 = {"op": "replace_text", "id": "p0", "old": "синее небо", "new": "ясное небо"}
    assert validate(blocks, op2, doc) is None
    apply(doc, idx, op2)
    blocks = doc_map(doc, idx)
    assert blocks[0]["text"] == "два слова ясное небо.", blocks[0]["text"]

    print("patch_demo: replace_text находит и корректно режет спан сквозь U+00A0")


if __name__ == "__main__":
    test_ops()
    test_invalid()
    test_hyperlink()
    test_create_table_has_borders()
    test_real_doc_style_error()
    test_replace_text_tolerates_nbsp()
