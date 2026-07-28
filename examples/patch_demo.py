"""Приёмка patch.py: девять операций на маленьком документе + валидатор невалидных патчей."""

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx_editor.parse import doc_map, index
from docx_editor.patch import _HANDLERS, _OPS, apply, validate

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


def test_set_text_rejects_ellipsis_truncation():
    # Находка Ф8: живой Редактор ответил set_text, укоротив абзац 456→332
    # знака и оборвав его буквальным «…» — смысловая правка была на месте,
    # поэтому Проверяющий сказал ok, а машинная проверка "текст изменился"
    # такое по устройству не ловит. validate обязан отбить это ДО применения.
    doc = _build()
    idx = index(doc)
    blocks = doc_map(doc, idx)

    bad = {"op": "set_text", "id": "p0", "text": "Первый…"}
    err = validate(blocks, bad, doc)
    assert isinstance(err, str) and err, "обрезанный set_text с многоточием обязан быть отбит"
    assert "p0" in err, err

    # многоточие тремя точками — тот же случай
    bad_dots = {"op": "set_text", "id": "p0", "text": "Первый..."}
    assert isinstance(validate(blocks, bad_dots, doc), str)

    # replace_text — та же защита, new короче old и обрывается многоточием
    bad_rt = {"op": "replace_text", "id": "p3", "old": "Срок 30 дней истекает.", "new": "Срок 30…"}
    assert isinstance(validate(blocks, bad_rt, doc), str)

    # легитимный случай: исходный текст УЖЕ оканчивается многоточием —
    # запрет не должен срабатывать на нём (в приёмочном документе таких
    # абзацев 0, но правило не должно бить по ним, если появятся)
    doc2 = Document()
    doc2.add_paragraph("Абзац с оборванной мыслью…")
    idx2 = index(doc2)
    blocks2 = doc_map(doc2, idx2)
    legit = {"op": "set_text", "id": "p0", "text": "Абзац с другой оборванной мыслью…"}
    assert validate(blocks2, legit, doc2) is None, validate(blocks2, legit, doc2)

    print("patch_demo: обрезание абзаца с многоточием в конце отбито, легитимное многоточие пропущено")


def test_replace_text_rejects_mid_word_cut():
    # Ф12: render() резал абзац ровно на границе в 300 знаков — "old" от
    # модели заканчивался посреди слова, _flex_span его честно находил (он
    # РЕАЛЬНО есть в документе), и хвост слова оставался приклеен к новому
    # тексту («...новых мощностей, херьёзных...»). validate обязан отбивать
    # такой "old" ДО применения: правая граница совпадения должна быть
    # границей слова, а не серединой.
    doc = Document()
    doc.add_paragraph("которые требуют серьёзных вычислительных ресурсов")
    idx = index(doc)
    blocks = doc_map(doc, idx)

    cut = {"op": "replace_text", "id": "p0", "old": "которые требуют с", "new": "которые требуют новых"}
    err = validate(blocks, cut, doc)
    assert isinstance(err, str) and "слова" in err, err

    # replace_all — та же защита (Ф12: тот же класс дефекта у replace_all)
    err_all = validate(blocks, {"op": "replace_all", "old": "которые требуют с", "new": "которые требуют новых"}, doc)
    assert isinstance(err_all, str) and "слова" in err_all, err_all

    # легитимный случай: правка суффикса внутри более длинного слова —
    # ЛЕВАЯ граница совпадения приходится на середину слова (как и в обрыве
    # выше), а ПРАВАЯ — на настоящий конец слова; гвард смотрит только на
    # правую и не должен сработать
    doc2 = Document()
    doc2.add_paragraph("Современные кросс-энкодеры работают быстро.")
    idx2 = index(doc2)
    blocks2 = doc_map(doc2, idx2)
    legit = {"op": "replace_text", "id": "p0", "old": "энкодеры", "new": "энкодера"}
    assert validate(blocks2, legit, doc2) is None, validate(blocks2, legit, doc2)

    # set_format — тот же класс защиты, что и replace_text/replace_all (Ф15)
    doc3 = Document()
    doc3.add_paragraph("которые требуют серьёзных вычислительных ресурсов")
    idx3 = index(doc3)
    blocks3 = doc_map(doc3, idx3)
    err_fmt = validate(blocks3, {"op": "set_format", "id": "p0", "old": "которые требуют с", "i": True}, doc3)
    assert isinstance(err_fmt, str) and "слова" in err_fmt, err_fmt

    print("patch_demo: обрыв «old» посреди слова отбит (replace_text, replace_all, set_format), легитимный суффикс пропущен")


def _add_numpr(p, ilvl, num_id):
    """Помечает абзац как элемент списка: w:pPr/w:numPr/{w:ilvl,w:numId} —
    ровно то, что parse._list читает из XML (Ф15: у _build() списков нет,
    их приходится собирать вручную, как test_hyperlink собирает гиперссылку)."""
    pPr = p._p.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    ilvl_el = OxmlElement("w:ilvl")
    ilvl_el.set(qn("w:val"), str(ilvl))
    numId_el = OxmlElement("w:numId")
    numId_el.set(qn("w:val"), str(num_id))
    numPr.append(ilvl_el)
    numPr.append(numId_el)
    pPr.append(numPr)


def test_set_format():
    # p3: "Срок 30 дней истекает." — "30 дней" лежит ровно на границах трёх
    # ранов разного начертания ("3" bold, "0 дн" italic, "ей" bold+underline);
    # соседи ("Срок " и " истекает.") без форматирования вовсе.
    doc = _build()
    idx = index(doc)
    blocks = doc_map(doc, idx)
    original_text = blocks[3]["text"]

    op = {"op": "set_format", "id": "p3", "old": "30 дней", "i": True}
    assert validate(blocks, op, doc) is None
    apply(doc, idx, op)
    blocks = doc_map(doc, idx)
    assert blocks[3]["text"] == original_text, "set_format не должен менять текст"

    runs = doc.paragraphs[3].runs
    assert runs[0].text == "Срок " and not runs[0].italic, "сосед слева получил курсив по ошибке"
    assert runs[-1].text == " истекает." and not runs[-1].italic, "сосед справа получил курсив по ошибке"
    assert all(r.italic for r in runs[1:-1]), [(r.text, r.italic) for r in runs]
    assert "".join(r.text for r in runs[1:-1]) == "30 дней"
    assert runs[1].bold and runs[3].bold and runs[3].underline, "старое форматирование внутри спана не должно было слететь"

    # false снимает то, что было true: "3" и "ей" были bold
    op2 = {"op": "set_format", "id": "p3", "old": "30 дней", "b": False}
    assert validate(blocks, op2, doc) is None
    apply(doc, idx, op2)
    runs = doc.paragraphs[3].runs
    assert not any(r.bold for r in runs[1:-1]), [(r.text, r.bold) for r in runs]
    assert all(r.italic for r in runs[1:-1]), "i не был указан в этой операции — не должен был тронуться"
    assert runs[3].underline, "u не был указан — не должен был слететь"

    print("patch_demo: set_format ставит и снимает начертание ровно на спане, соседи и текст не тронуты")


def test_set_format_splits_run_boundary():
    # p0: "Первый абзац." — один ран целиком. old="рвый абзац" начинается и
    # заканчивается ВНУТРИ этого рана — set_format обязан физически разрезать
    # ран (не просто переписать текст, как делает _replace_span), иначе
    # префикс/суффикс получат чужое форматирование вместе со спаном.
    doc = _build()
    idx = index(doc)
    blocks = doc_map(doc, idx)

    op = {"op": "set_format", "id": "p0", "old": "рвый абзац", "b": True}
    assert validate(blocks, op, doc) is None
    apply(doc, idx, op)
    blocks = doc_map(doc, idx)
    assert blocks[0]["text"] == "Первый абзац.", blocks[0]["text"]

    runs = doc.paragraphs[0].runs
    assert "".join(r.text for r in runs) == "Первый абзац."
    assert not runs[0].bold and runs[0].text == "Пе"
    assert not runs[-1].bold and runs[-1].text == "."
    middle = runs[1:-1]
    assert "".join(r.text for r in middle) == "рвый абзац"
    assert all(r.bold for r in middle)

    print("patch_demo: set_format физически режет ран по границам спана")


def test_set_format_rejects_no_flags():
    doc = _build()
    idx = index(doc)
    blocks = doc_map(doc, idx)
    err = validate(blocks, {"op": "set_format", "id": "p0", "old": "Первый"}, doc)
    assert isinstance(err, str) and err, "без b/i/u операция не меняет ничего — обязан быть отказ"
    print("patch_demo: set_format без b/i/u отбит валидатором")


def test_set_list_level():
    doc = _build()
    doc.add_paragraph("Пункт списка")
    idx = index(doc)
    blocks = doc_map(doc, idx)
    list_id = blocks[-1]["id"]
    _add_numpr(doc.paragraphs[-1], ilvl=0, num_id=5)
    blocks = doc_map(doc, idx)
    assert blocks[-1]["list"] == {"ilvl": 0, "numId": 5}, blocks[-1]["list"]

    op = {"op": "set_list_level", "id": list_id, "ilvl": 1}
    assert validate(blocks, op, doc) is None
    apply(doc, idx, op)
    blocks = doc_map(doc, idx)
    assert blocks[-1]["list"] == {"ilvl": 1, "numId": 5}, "numId обязан уцелеть, меняется только ilvl"

    # p0 в списке не состоит вовсе — сменить уровень нельзя
    err = validate(blocks, {"op": "set_list_level", "id": "p0", "ilvl": 1}, doc)
    assert isinstance(err, str) and "списк" in err, err

    print("patch_demo: set_list_level меняет ilvl и сохраняет numId, отказан на абзаце вне списка")


def test_all_ops_have_handlers():
    # Находка BUILD_PLAN: коммит однажды завёл set_format в _OPS без записи
    # в _HANDLERS — apply() падал бы KeyError на первом же применении.
    missing = _OPS - set(_HANDLERS)
    assert not missing, f"операции без обработчика в _HANDLERS: {missing}"
    print("patch_demo: у каждой операции из _OPS есть обработчик в _HANDLERS")


if __name__ == "__main__":
    test_all_ops_have_handlers()
    test_ops()
    test_invalid()
    test_hyperlink()
    test_create_table_has_borders()
    test_real_doc_style_error()
    test_replace_text_tolerates_nbsp()
    test_set_text_rejects_ellipsis_truncation()
    test_replace_text_rejects_mid_word_cut()
    test_set_format()
    test_set_format_splits_run_boundary()
    test_set_format_rejects_no_flags()
    test_set_list_level()
