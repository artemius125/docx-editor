"""Цикл одной правки: Навигатор → поиск → Редактор → валидатор → применение →
Проверка → вердикт. Плюс split() — детерминированное деление списка правок.

Три роли-модели (Навигатор/Редактор/Проверяющий) — параметры run_edit с
реализацией по умолчанию: тесты и демо подставляют вместо них фейки, не
тратя вызовы модели на то, что не требует живого ответа (см. CLAUDE.md:
демо, закрываемое фикстурой, закрываем фикстурой).
"""

import os
import re
import tempfile

from docx import Document

from docx_editor import find, llm, patch
from docx_editor.parse import doc_map, index, render

_NAV_PROMPT = """Ты — навигатор по документу .docx. Тебе дают компактное
оглавление (id и начало текста абзаца на строку) и текст правки на русском.
Верни ТОЛЬКО JSON {"kind":"local"|"global","rule":null|"typography"|"quotes",
"ids":[...],"anchors":[...]} без пояснений и markdown-обвязки.

- "rule" — только когда правка ГЛОБАЛЬНАЯ типографская по всему документу:
  "typography" — лишний пробел перед двоеточием/точкой с запятой/закрывающей
  скобкой, слипшиеся предложения (нет пробела после точки между ними),
  невидимые символы; "quotes" — прямые кавычки вместо ёлочек. Иначе rule = null.
- "ids" — id блоков из оглавления, которые точно относятся к правке, если
  видишь их явно. Не пиши id внутрь anchors — это отдельное поле.
- "anchors" — дословные цитаты ИЗ ТЕКСТА ПРАВКИ, самые специфичные, какие
  найдёшь (не термин своими словами). Рядом с формулой цитируй окружающую
  прозу, а не саму формулу — формулы в правке и в документе записаны по-разному.
- "kind" = "global" можно ставить, только если ты вообще не можешь назвать ни
  одного id. Если можешь назвать хотя бы один — это "local", а не "global"."""

_EDIT_PROMPT = """Ты — редактор документа .docx. Тебе дают фрагмент блоков
(построчно: "pNN [Стиль] текст" для абзацев, "tNN [table RxC] ..." для таблиц)
и текст правки. Верни ТОЛЬКО JSON {"ops": [...]} — операции над блоками ИЗ
ЭТОГО ФРАГМЕНТА. Если правку нельзя выразить перечисленными операциями — не
подменяй её похожей и не изобретай новую: верни {"ops": [], "note": "почему
нельзя"}, это честный и ожидаемый ответ, а не провал.

Операции (id — только из фрагмента, other — новых не изобретать):
{"op":"replace_text","id":"p12","old":"...","new":"..."} — заменить кусок текста в абзаце
{"op":"set_text","id":"p12","text":"..."} — заменить текст абзаца целиком
{"op":"insert_after","id":"p12","text":"...","style":"Normal"} — вставить абзац после блока
{"op":"delete","id":"p12"} — удалить блок
{"op":"move_after","id":"p12","after":"p19"} — переместить блок после другого
{"op":"set_style","id":"p12","style":"Heading 1"} — сменить именованный стиль абзаца
{"op":"create_table","after":"p57","rows":[["a","b"]],"header":true} — создать таблицу после блока
{"op":"set_cell","id":"t3","row":0,"col":1,"text":"..."} — заменить текст ячейки таблицы
{"op":"replace_all","old":"...","new":"..."} — заменить текст ВЕЗДЕ в документе (без id)

"old" обязан быть текстом, который реально есть в блоке (дословно, из фрагмента)."""

_CHECK_PROMPT = """Ты — проверяющий правок документа .docx. Тебе дают текст
правки и список реально изменившихся блоков вида "id: было N зн. «...»
стало M зн. «...»". Скажи, точно ли эти изменения выполняют именно эту
правку — не наоборот, не мимо цели, не пустышка (было эквивалентно стало).
Если стало заметно короче было (текст обрублен, конец потерян), это провал,
ДАЖЕ ЕСЛИ нужная по смыслу правка внутри текста присутствует — часть текста
пропала, и это надо отклонить. Верни ТОЛЬКО JSON {"ok": true|false,
"reason": "..."} без пояснений."""


def _navigate(outline_text, request):
    messages = [
        {"role": "system", "content": _NAV_PROMPT},
        {"role": "user", "content": f"Оглавление документа:\n{outline_text}\n\nПравка: {request}"},
    ]
    return llm.chat(messages)


def _edit_llm(fragment_text, request, feedback=None):
    user = f"Фрагмент документа:\n{fragment_text}\n\nПравка: {request}"
    if feedback:
        user += f"\n\n{feedback}"
    messages = [
        {"role": "system", "content": _EDIT_PROMPT},
        {"role": "user", "content": user},
    ]
    return llm.chat(messages)


def _check(request, diff):
    # Длины добавлены находкой Ф8: без чисел Проверяющий видит, что нужная
    # правка внутри текста есть, и молчит про обрубленный хвост абзаца —
    # с числами модель хорошо ловит «стало вдвое короче». Форма числа
    # инвариантна («зн.») — читатель сравнивает два числа, не человек.
    diff_text = "\n".join(
        f'{d["id"]}: было {len(d["before"])} зн. «{d["before"]}» стало {len(d["after"])} зн. «{d["after"]}»'
        for d in diff
    )
    messages = [
        {"role": "system", "content": _CHECK_PROMPT},
        {"role": "user", "content": f"Правка: {request}\n\nИзменения:\n{diff_text}"},
    ]
    return llm.chat(messages)


_ID_PREFIX = re.compile(r"^p\d+\s+")


def _strip_anchor(anchor):
    """Защита от находки Ф5: модель иногда кладёт id прямо внутрь цитаты якоря."""
    return _ID_PREFIX.sub("", anchor)


def _clean_ops(ops):
    """Отбрасывает мусорные элементы списка ops — на 26B модель иногда кладёт
    посторонний примитив (например 0) первым элементом рядом с валидными
    операциями. Валидатор всё равно проверит каждую оставшуюся операцию по
    существу — здесь только защита от AttributeError на не-словаре."""
    return [op for op in (ops or []) if isinstance(op, dict)]


def _dedupe(ids):
    seen, out = set(), []
    for i in ids:
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


def _normalize_id(raw):
    """Приводит id от Навигатора к канонической форме блоков ("p12", "t3"):
    модель то возвращает голое число (112 вместо "p112"), то путает регистр
    префикса ("P112") — код, не промпт, отвечает за форму (находка Ф5, тот
    же класс, что и id внутри цитаты якоря). Голое число получает префикс p,
    иначе просто нижний регистр; итог сверяется с реальными id вызывающим."""
    s = str(raw).strip()
    return f"p{s}" if s.isdigit() else s.lower()


def _resolve(blocks, nav, request):
    """id из Навигатора (провалидированные) + попадания по якорям → id блоков.
    Пусто — запасной поиск по дословным цитатам из текста правки (заказчик
    пишет их в «ёлочках»). Пусто и там — [] означает честный отказ выше."""
    real_ids = {b["id"] for b in blocks}
    ids = [n for n in (_normalize_id(i) for i in (nav.get("ids") or [])) if n in real_ids]
    for anchor in nav.get("anchors") or []:
        ids += find.by_text(blocks, _strip_anchor(anchor))
    if ids:
        return _dedupe(ids)
    for quote in re.findall(r"«([^»]+)»", request):
        ids += find.by_text(blocks, quote)
    return _dedupe(ids)


def _snapshot(doc):
    """Сохраняет документ во временный файл — точка отката ПЕРЕД применением."""
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    doc.save(path)
    return path


def _restore(snapshot_path):
    """Перечитывает документ из снимка. Возвращает НОВЫЕ (doc, idx): id из
    старого idx были ссылками на lxml-элементы прежнего дерева, их больше
    нет — вызывающий обязан взять из возврата, а не держать старый doc."""
    doc = Document(snapshot_path)
    idx = index(doc)
    os.remove(snapshot_path)
    return doc, idx


def _btext(b):
    return b["text"] if b["kind"] == "p" else " | ".join(c for row in b["rows"] for c in row)


def _diff(before, after):
    """Блоки, у которых текст реально изменился, плюс исчезнувшие (delete) и
    появившиеся (insert_after/create_table) — только это видит Проверяющий,
    не документ целиком (инвариант 3 из CLAUDE.md)."""
    b_by_id = {b["id"]: _btext(b) for b in before}
    a_by_id = {b["id"]: _btext(b) for b in after}
    changed = [{"id": i, "before": b_by_id.get(i) or "", "after": t}
               for i, t in a_by_id.items() if b_by_id.get(i) != t]
    changed += [{"id": i, "before": t, "after": ""} for i, t in b_by_id.items() if i not in a_by_id]
    return changed


def _apply_ops(doc, idx, ops, fragment_text, request, editor):
    """Применяет ops по очереди; на невалидной операции — один ретрай к
    Редактору с текстом ошибки (editor=None — ретрая нет, путь rule). Второй
    провал (или провал без Редактора) останавливает применение немедленно.
    Возвращает (описания применённых операций, причина отказа-или-None,
    tries) — tries=2, если к Редактору пришлось обращаться повторно
    (для журнала правки, поле "iter"), иначе 1."""
    applied, retried, i = [], False, 0
    while i < len(ops):
        op = ops[i]
        err = patch.validate(doc_map(doc, idx), op, doc)
        if err is None:
            applied.append(patch.apply(doc, idx, op))
            i += 1
            continue
        if retried or editor is None:
            return applied, err, 2 if retried else 1
        retried = True
        feedback = f'Операция {op} не прошла проверку: {err}. Пришли исправленный {{"ops": [...]}}.'
        reply = editor(fragment_text, request, feedback=feedback) or {}
        ops = _clean_ops(reply.get("ops"))
        if not ops:
            return applied, reply.get("note") or err, 2
        i = 0
    return applied, None, 2 if retried else 1


def _failed(reason, nav=None, editor_reply=None, tries=1):
    """reply/iter обязаны быть в каждой записи, включая честные отказы —
    именно они интереснее всего в журнале для разбора «куда целился Навигатор»."""
    return {
        "verdict": "failed", "reason": reason, "applied": [], "ids": [],
        "iter": tries, "reply": {"nav": nav, "editor": editor_reply},
    }


def run_edit(doc, idx, request, navigator=_navigate, editor=_edit_llm, checker=_check):
    """Один цикл правки. Возвращает (result, doc, idx) — после отката
    (verdict != "done" после применения) doc/idx уже перечитаны из снимка,
    старые ссылки вызывающего использовать нельзя.

    result содержит "reply": {"nav": ..., "editor": ...} — РАЗОБРАННЫЙ JSON-
    ответ ролей (llm.chat парсит его сам, сырое тело HTTP нигде не хранится),
    "editor": None, когда Редактора не звали вообще (путь rule), а не {} —
    это разные вещи для разбора; и "iter" — 1 или 2 (был ли ретрай Редактора
    внутри _apply_ops), настоящее число попыток, а не декоративная константа."""
    blocks = doc_map(doc, idx)
    nav = navigator(find.outline(blocks), request) or {}
    rule = nav.get("rule")

    if rule:
        # Правки типа rule идут мимо Редактора — код нормализует сам, LLM не зовём.
        ops, fragment_text, use_editor, editor_reply = [{"op": "normalize", "rule": rule}], "", None, None
    else:
        ids = _resolve(blocks, nav, request)
        if not ids:
            reason = "не нашёл, где это править: ни id и якоря Навигатора, ни дословные цитаты из текста правки не нашли места в документе"
            return _failed(reason, nav), doc, idx
        fragment_text = render(find.fragment(blocks, ids, around=1))
        editor_reply = editor(fragment_text, request) or {}
        ops = _clean_ops(editor_reply.get("ops"))
        if not ops:
            return _failed(editor_reply.get("note") or "редактор не смог предложить операции для этой правки", nav, editor_reply), doc, idx
        use_editor = editor

    snapshot_path = _snapshot(doc)
    blocks_before = doc_map(doc, idx)
    applied, err, tries = _apply_ops(doc, idx, ops, fragment_text, request, use_editor)

    if not applied:
        doc, idx = _restore(snapshot_path)
        return _failed(err or "ни одна операция не применилась", nav, editor_reply, tries), doc, idx

    diff = _diff(blocks_before, doc_map(doc, idx))
    if not diff:
        doc, idx = _restore(snapshot_path)
        return _failed("операции применились, но текст не изменился", nav, editor_reply, tries), doc, idx

    verdict = checker(request, diff) or {}
    ids_touched = [d["id"] for d in diff]
    reply = {"nav": nav, "editor": editor_reply}
    if verdict.get("ok"):
        os.remove(snapshot_path)
        return {"verdict": "done", "reason": verdict.get("reason", ""), "applied": applied, "ids": ids_touched,
                 "iter": tries, "reply": reply}, doc, idx

    doc, idx = _restore(snapshot_path)
    return {"verdict": "rolled_back", "reason": verdict.get("reason", ""), "applied": applied, "ids": ids_touched,
             "iter": tries, "reply": reply}, doc, idx


_ITEM_MARKER = re.compile(r"^(?:\*\*(\d+)\.\*\*|(\d+)[.)]|[-*])\s+")


def split(request):
    """Делит многопунктный запрос на атомарные правки по markdown-нумерации
    (**1.**, 1., 1), -, * в начале строки). Заголовки (#, ##…) и пустые строки
    отбрасываются, текст пункта тянется до следующего маркера — многострочные
    пункты не рвутся. Без структуры — [request] без изменений. Перенесено из
    прошлого проекта (docx-agent/agent.py) — там уже проверено на этом же
    файле правок."""
    items, current = [], None
    for line in request.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        m = _ITEM_MARKER.match(stripped)
        if m:
            if current is not None:
                items.append(current.strip())
            current = stripped[m.end():]
        elif current is not None:
            current += " " + stripped
    if current is not None:
        items.append(current.strip())
    return items if items else [request]
