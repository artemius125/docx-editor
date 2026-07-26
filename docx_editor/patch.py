"""Валидация патча (до применения) и применение девяти операций к живому документу.

Патч — список операций из контракта BUILD_PLAN.md (кроме normalize, она в Ф3):
replace_text, set_text, insert_after, delete, move_after, set_style,
create_table, set_cell, replace_all.
"""

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

_OPS = {
    "replace_text", "set_text", "insert_after", "delete", "move_after",
    "set_style", "create_table", "set_cell", "replace_all",
}
_PARAGRAPH_ONLY = {"replace_text", "set_text", "set_style"}


def _by_id(blocks):
    return {b["id"]: b for b in blocks}


def _table_texts(block):
    return [cell for row in block["rows"] for cell in row]


def _paragraph_style_names(doc):
    return [s.name for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH]


def _style_error(style, doc):
    """None — стиль есть в документе; иначе текст ошибки со списком доступных."""
    names = _paragraph_style_names(doc)
    if style in names:
        return None
    return f"стиль {style!r} не найден; доступны стили абзацев: {', '.join(names)}"


def validate(blocks, op, doc):
    """None — патч валиден; иначе русский текст ошибки для повторной попытки модели.

    doc нужен только для set_style: полный список стилей абзацев, реально
    определённых в документе, берётся из doc.styles — в blocks есть только
    стили, УЖЕ применённые к каким-то абзацам, а этого недостаточно (в
    Архитектура_ColBERT.docx все абзацы Normal, но List Paragraph в
    документе определён и доступен).
    """
    name = op.get("op")
    if name not in _OPS:
        return f"неизвестная операция {name!r}; допустимые: {sorted(_OPS)}"

    by_id = _by_id(blocks)

    if name in _PARAGRAPH_ONLY or name in ("insert_after", "delete", "move_after"):
        block_id = op.get("id")
        if block_id not in by_id:
            return f"блок {block_id!r} не найден в документе"
        if name in _PARAGRAPH_ONLY and by_id[block_id]["kind"] != "p":
            return f"{block_id} — это таблица, {name} работает только с абзацами"

    if name == "insert_after" and op.get("style"):
        err = _style_error(op["style"], doc)
        if err:
            return err

    if name == "move_after":
        after = op.get("after")
        if after not in by_id:
            return f"блок {after!r} (after) не найден в документе"

    if name == "create_table":
        after = op.get("after")
        if after not in by_id:
            return f"блок {after!r} (after) не найден в документе"
        rows = op.get("rows")
        if not rows or not rows[0]:
            return "create_table: rows не может быть пустым"
        ncols = len(rows[0])
        if any(len(r) != ncols for r in rows):
            return f"create_table: строки разной длины: {[len(r) for r in rows]}"

    if name == "set_cell":
        block_id = op.get("id")
        if block_id not in by_id:
            return f"блок {block_id!r} не найден в документе"
        block = by_id[block_id]
        if block["kind"] != "t":
            return f"{block_id} — это абзац, set_cell работает только с таблицами"
        nrows = len(block["rows"])
        ncols = len(block["rows"][0]) if nrows else 0
        row, col = op.get("row"), op.get("col")
        if not (isinstance(row, int) and 0 <= row < nrows) or not (isinstance(col, int) and 0 <= col < ncols):
            return f"ячейка ({row},{col}) вне таблицы {block_id}: размер {nrows}x{ncols}"

    if name in ("replace_text", "replace_all"):
        old, new = op.get("old"), op.get("new")
        if old == new:
            return f"old и new совпадают («{old}») — пустая операция, нечего менять"
        if name == "replace_text":
            if old not in by_id[op["id"]]["text"]:
                return f"в {op['id']} нет текста «{old}»"
        else:
            found = any(b["kind"] == "p" and old in b["text"] for b in blocks) or any(
                b["kind"] == "t" and any(old in c for c in _table_texts(b)) for b in blocks
            )
            if not found:
                return f"текст «{old}» не найден нигде в документе"

    if name == "set_style":
        err = _style_error(op.get("style"), doc)
        if err:
            return err

    return None


def _runs(p_el):
    """Все раны абзаца в порядке документа, включая раны внутри w:hyperlink.

    p_el.r_lst — это только раны верхнего уровня; текст гиперссылки в него
    не попадает, а p.text (то, что видит validate и модель) гиперссылки
    включает. Расхождение офсетов между _ptext и p.text — гарантированная
    порча документа (find возвращает -1 или не то смещение). iter() обходит
    поддерево в порядке документа, поэтому раны внутри w:hyperlink встают
    на своё естественное место и офсеты снова совпадают с p.text.
    """
    return list(p_el.iter(qn("w:r")))


def _ptext(p_el):
    """Текст абзаца, собранный ровно из тех же ранов, что и _replace_span."""
    return "".join(r.text for r in _runs(p_el))


def _replace_span(p_el, start, end, new_text):
    """Заменяет [start:end) в тексте абзаца на new_text, не трогая чужие раны.

    Граница замены может резать чужой ран пополам (фраза лежит на стыке
    ранов с разным начертанием) — тогда затронутые крайние раны режутся по
    границе, а раны целиком внутри диапазона стираются, чужие раны снаружи
    не трогаются вовсе.
    """
    if start < 0 or end < start:
        raise ValueError(f"недопустимый диапазон замены [{start}:{end})")
    runs = _runs(p_el)
    if not runs:
        p_el.add_r().text = new_text
        return
    spans, pos = [], 0
    for r in runs:
        t = r.text
        spans.append((pos, pos + len(t)))
        pos += len(t)
    touched = [i for i, (s, e) in enumerate(spans) if e > start and s < end]
    if not touched:
        runs[-1].text = runs[-1].text + new_text
        return
    fi, li = touched[0], touched[-1]
    prefix = runs[fi].text[: start - spans[fi][0]]
    suffix = runs[li].text[end - spans[li][0]:]
    if fi == li:
        runs[fi].text = prefix + new_text + suffix
    else:
        runs[fi].text = prefix + new_text
        runs[li].text = suffix
        for i in range(fi + 1, li):
            runs[i].text = ""


def _register(idx, prefix, el):
    n = 0
    while f"{prefix}{n}" in idx:
        n += 1
    new_id = f"{prefix}{n}"
    idx[new_id] = el
    return new_id


def _op_replace_text(doc, idx, op):
    el = idx[op["id"]]
    full = _ptext(el)
    pos = full.find(op["old"])
    if pos == -1:
        # validate должен был отсечь это раньше; если добрались сюда — громко падаем,
        # а не режем абзац по отрицательному смещению (см. находку про гиперссылки)
        raise ValueError(f"текст {op['old']!r} не найден в {op['id']} на момент применения")
    _replace_span(el, pos, pos + len(op["old"]), op["new"])
    return f"В {op['id']} заменено «{op['old']}» на «{op['new']}»"


def _op_set_text(doc, idx, op):
    el = idx[op["id"]]
    _replace_span(el, 0, len(_ptext(el)), op["text"])
    return f"Текст {op['id']} заменён целиком"


def _op_insert_after(doc, idx, op):
    ref = idx[op["id"]]
    style = op.get("style") or Paragraph(ref, doc).style
    new_p = doc.add_paragraph(op["text"], style=style)
    ref.addnext(new_p._p)
    new_id = _register(idx, "p", new_p._p)
    return f"После {op['id']} вставлен новый абзац {new_id}"


def _op_delete(doc, idx, op):
    el = idx[op["id"]]
    el.getparent().remove(el)
    return f"Блок {op['id']} удалён"


def _op_move_after(doc, idx, op):
    src, dst = idx[op["id"]], idx[op["after"]]
    dst.addnext(src)
    return f"Блок {op['id']} перемещён после {op['after']}"


def _op_set_style(doc, idx, op):
    Paragraph(idx[op["id"]], doc).style = op["style"]
    return f"Стиль {op['id']} изменён на {op['style']!r}"


def _add_borders(tbl):
    """Одиночные границы прямо в tblPr. В контракте create_table нет поля
    style, а из именованных табличных стилей в документе часто есть только
    Normal Table (без рамок) — рассчитывать на style="Table Grid" нельзя,
    его может не быть (KeyError). Ставим границы всегда, без опций."""
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    tbl.tblPr.append(borders)


def _op_create_table(doc, idx, op):
    ref = idx[op["after"]]
    rows = op["rows"]
    ncols = len(rows[0])
    table = doc.add_table(rows=len(rows), cols=ncols)
    _add_borders(table._tbl)
    for r, row_vals in enumerate(rows):
        for c, val in enumerate(row_vals):
            table.cell(r, c).text = val
    if op.get("header"):
        for cell in table.rows[0].cells:
            for run in cell.paragraphs[0].runs:
                run.bold = True
    ref.addnext(table._tbl)
    new_id = _register(idx, "t", table._tbl)
    return f"После {op['after']} создана таблица {new_id} {len(rows)}x{ncols}"


def _op_set_cell(doc, idx, op):
    table = Table(idx[op["id"]], doc)
    p_el = table.cell(op["row"], op["col"]).paragraphs[0]._p
    _replace_span(p_el, 0, len(_ptext(p_el)), op["text"])
    return f"Ячейка {op['id']}[{op['row']}][{op['col']}] заменена на «{op['text']}»"


def _op_replace_all(doc, idx, op):
    old, new = op["old"], op["new"]
    total = blocks_touched = 0
    for p_el in doc.element.body.iter(qn("w:p")):
        full = _ptext(p_el)
        matches, pos = [], 0
        while True:
            i = full.find(old, pos)
            if i == -1:
                break
            matches.append(i)
            pos = i + len(old)
        if not matches:
            continue
        for m_start in reversed(matches):
            _replace_span(p_el, m_start, m_start + len(old), new)
        total += len(matches)
        blocks_touched += 1
    return f"Заменено {total} вхождений «{old}» на «{new}» в {blocks_touched} блоках (включая ячейки таблиц)"


_HANDLERS = {
    "replace_text": _op_replace_text,
    "set_text": _op_set_text,
    "insert_after": _op_insert_after,
    "delete": _op_delete,
    "move_after": _op_move_after,
    "set_style": _op_set_style,
    "create_table": _op_create_table,
    "set_cell": _op_set_cell,
    "replace_all": _op_replace_all,
}


def apply(doc, idx, op):
    """Применяет одну из девяти операций (патч уже прошёл validate), возвращает описание."""
    return _HANDLERS[op["op"]](doc, idx, op)
