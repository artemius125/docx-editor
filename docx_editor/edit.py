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

- "rule" — только когда правка ГЛОБАЛЬНАЯ типографская и относится РОВНО к
  одному из двух названных классов, других классов rule не бывает:
  "typography" — лишний пробел перед двоеточием/точкой с запятой/закрывающей
  скобкой, слипшиеся предложения (нет пробела после точки между ними),
  невидимые символы; "quotes" — прямые кавычки вместо ёлочек. Во всех
  остальных случаях rule = null, даже если правка звучит "по всему
  документу". Например, "термин Bi-encoder пишется четырьмя разными
  способами, приведи к одному" — это НЕ rule (это не типографика, а
  единообразие термина), верни local с ids/anchors на найденные варианты —
  это работа Редактора, а не нормализации.
- "ids" — id блоков из оглавления, которые точно относятся к правке, если
  видишь их явно. Не пиши id внутрь anchors — это отдельное поле.
- "anchors" — дословные цитаты ИЗ ТЕКСТА ПРАВКИ, самые специфичные, какие
  найдёшь (не термин своими словами). Рядом с формулой цитируй окружающую
  прозу, а не саму формулу — формулы в правке и в документе записаны по-разному.
- [B] перед текстом строки — абзац целиком жирный, хотя стиля заголовка у
  него нет; полезно для правок вида «заголовок — это просто жирный абзац».
- "kind" = "global" можно ставить, только если ты вообще не можешь назвать ни
  одного id. Если можешь назвать хотя бы один — это "local", а не "global"."""

_EDIT_PROMPT = """Ты — редактор документа .docx. Тебе дают фрагмент блоков
(построчно: "pNN [метаданные] текст" для абзацев, "tNN [table RxC] ..." для
таблиц) и текст правки. Всё внутри [] — метаданные блока: стиль абзаца и,
если есть, уровень заголовка (H1…), номер уровня списка, оформление («весь
жирный/курсивный/подчёркнутый» или, если оформлена только часть, «жирным:
«кусок текста»»). Текст ПОСЛЕ закрывающей ] — дословный текст блока: "old"
обязан быть процитирован буквально из него, метки из [] в "old" не входят.
Верни ТОЛЬКО JSON {"ops": [...]} — операции над блоками ИЗ ЭТОГО ФРАГМЕНТА.
Если правку нельзя выразить перечисленными операциями — не подменяй её
похожей и не изобретай новую: верни {"ops": [], "note": "почему нельзя"},
это честный и ожидаемый ответ, а не провал.

Операции (id — только из фрагмента, other — новых не изобретать):
{"op":"replace_text","id":"p12","old":"...","new":"..."} — заменить кусок текста в абзаце
{"op":"set_text","id":"p12","text":"..."} — заменить текст абзаца целиком
{"op":"insert_after","id":"p12","text":"...","style":"Normal"} — вставить абзац после блока
{"op":"delete","id":"p12"} — удалить блок
{"op":"move_after","id":"p12","after":"p19"} — переместить блок после другого
{"op":"set_style","id":"p12","style":"Heading 1"} — сменить именованный стиль абзаца
{"op":"create_table","after":"p57","rows":[["a","b"]],"header":true} — создать таблицу после блока
{"op":"set_cell","id":"t3","row":0,"col":1,"text":"..."} — заменить текст ячейки таблицы
{"op":"replace_all","old":"...","new":"..."} — заменить текст ВЕЗДЕ в документе (без id)"""

_CHECK_PROMPT = """Ты — проверяющий правок документа .docx. Тебе дают текст
правки и список реально изменившихся блоков вида "id: было N зн. «...»
стало M зн. «...»", а если поменялся стиль абзаца, его уровень, список или
позиция в документе — строкой вида "id: текст не изменился («...»), но
стиль «Normal» → «Heading 1»" (или припиской после длин, если изменился и
текст тоже). Такая структурная правка — смена стиля, уровня или позиции —
реальное изменение, а не пустышка, даже если сам текст не тронут ни на
символ: не отклоняй её только из-за того, что "было" и "стало" совпадают,
если строка называет изменение стиля/уровня/позиции. Скажи, точно ли эти
изменения выполняют именно эту правку — не наоборот, не мимо цели, не
пустышка (было эквивалентно стало и структура тоже не менялась). Если стало
заметно короче было (текст обрублен, конец потерян), это провал, ДАЖЕ ЕСЛИ
нужная по смыслу правка внутри текста присутствует — часть текста пропала,
и это надо отклонить. Верни ТОЛЬКО JSON {"ok": true|false, "reason": "..."}
без пояснений."""


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
    #
    # Change A: запись diff'а может нести "note" (стиль/уровень/список/
    # позиция) без изменения текста — строка "было 200 зн. «X» стало 200 зн.
    # «X»" в этом случае читалась бы моделью как пустышка. Поэтому если текст
    # не изменился, строка строится вокруг note, а не вокруг длин; если
    # изменилось и то и другое — note дописывается к строке с длинами.
    lines = []
    for d in diff:
        note = d.get("note")
        if d["before"] == d["after"]:
            lines.append(f'{d["id"]}: текст не изменился («{d["after"]}»), но {note}')
        else:
            line = f'{d["id"]}: было {len(d["before"])} зн. «{d["before"]}» стало {len(d["after"])} зн. «{d["after"]}»'
            if note:
                line += f'; {note}'
            lines.append(line)
    diff_text = "\n".join(lines)
    messages = [
        {"role": "system", "content": _CHECK_PROMPT},
        {"role": "user", "content": f"Правка: {request}\n\nИзменения:\n{diff_text}"},
    ]
    return llm.chat(messages)


_ID_PREFIX = re.compile(r"^p\d+\s+")


def _strip_anchor(anchor):
    """Защита от находки Ф5: модель иногда кладёт id прямо внутрь цитаты якоря."""
    return _ID_PREFIX.sub("", anchor)


_QUOTE_RX = re.compile(r"«([^»]+)»|`([^`]+)`|\"([^\"]+)\"|“([^”]+)”")


def _quotes(request):
    """Дословные цитаты из текста правки во всех употребимых кавычках —
    «ёлочках», `обратных`, "прямых" и “типографских” — короче 2 знаков не считаем."""
    out = []
    for m in _QUOTE_RX.finditer(request):
        q = next(g for g in m.groups() if g is not None)
        if len(q) >= 2:
            out.append(q)
    return out


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
    """id из Навигатора (провалидированные) + попадания по якорям + попадания
    по дословным цитатам («…», `…`, "…", "…") из текста правки — ВСЕ ТРИ
    источника объединяются и дедуплицируются, а не по очереди с откатом на
    следующий.

    Находка Ф8 (item C): цитата раньше искалась только как запасной путь,
    когда ids и anchors ничего не дали — а правки вида «сделай одинаково
    везде» (например №8: три разных написания года, каждое в «ёлочках»)
    Навигатор называет ОДНИМ id/anchor (первый случай, который увидел), и
    запасной путь никогда не срабатывал: ids были не пусты, а остальные
    случаи из документа терялись, фрагмент давал Редактору только один
    случай, и правка честно, но неверно отказывалась «в документе больше
    случаев нет». Цитата теперь всегда доливает свои совпадения.

    Люди цитируют документ неточно (тире вместо запятой и т.п.) — на двух
    приёмочных корпусах 11 таких якорей/цитат не находит ни точный, ни
    пробельно-гибкий поиск. Поэтому для каждого якоря и каждой цитаты по
    отдельности: если find.by_text не нашёл ни одного блока, в дело
    вступает find.locate (приблизительный поиск по токенам) — только как
    запасной путь для ЭТОГО якоря/цитаты, точные попадания find.by_text
    им не заменяются и остаются первыми.

    Пусто после всех трёх источников — [] означает честный отказ выше."""
    real_ids = {b["id"] for b in blocks}
    ids = [n for n in (_normalize_id(i) for i in (nav.get("ids") or [])) if n in real_ids]
    for anchor in nav.get("anchors") or []:
        anchor = _strip_anchor(anchor)
        ids += find.by_text(blocks, anchor) or find.locate(blocks, anchor)
    for quote in _quotes(request):
        ids += find.by_text(blocks, quote) or find.locate(blocks, quote)
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


def _struct_note(b, a):
    """Текст структурного отличия абзаца b→a (стиль/уровень/список), или None.
    У таблиц структуры нет — там сравнивается только текст ячеек, это уже
    делает текстовое сравнение в _diff."""
    if b["kind"] != "p" or a["kind"] != "p":
        return None
    parts = []
    if b["style"] != a["style"]:
        parts.append(f'стиль «{b["style"]}» → «{a["style"]}»')
    if b.get("level") != a.get("level"):
        parts.append(f'уровень {b.get("level")} → {a.get("level")}')
    if b.get("list") != a.get("list"):
        parts.append(f'список {b.get("list")} → {a.get("list")}')
    return "; ".join(parts) if parts else None


def _move_note(before, after, a_by_id):
    """Запись diff'а на переместившийся блок, если множество id одинаково, а
    порядок отличается (move_after), иначе None. Переместившимся считается id
    с наибольшим сдвигом индекса — соседи, чья позиция сдвинулась заодно с
    ним на +-1, отдельной записи не получают (иначе на один move_after
    приходилось бы N записей вместо одной)."""
    order_before = [b["id"] for b in before]
    order_after = [b["id"] for b in after]
    if order_before == order_after:
        return None
    pos_before = {i: n for n, i in enumerate(order_before)}
    pos_after = {i: n for n, i in enumerate(order_after)}
    moved_id = max(pos_before, key=lambda i: abs(pos_before[i] - pos_after[i]))
    idx_after = pos_after[moved_id]
    after_id = order_after[idx_after - 1] if idx_after > 0 else None
    note = f"перемещён после {after_id}" if after_id else "перемещён в начало документа"
    text = _btext(a_by_id[moved_id])
    return {"id": moved_id, "before": text, "after": text, "note": note}


def _diff(before, after):
    """Блоки, у которых текст, стиль, уровень или список реально изменились,
    плюс исчезнувшие (delete) и появившиеся (insert_after/create_table) —
    только это видит Проверяющий, не документ целиком (инвариант 3 из
    CLAUDE.md).

    Раньше сравнивался только _btext — set_style и move_after были невидимы
    diff'у, Проверяющий получал пустой список изменений и run_edit откатывал
    правку с "текст не изменился", хотя структура реально поменялась. Теперь
    блоки сравниваются целиком (текст + style/level/list), а перестановка
    блоков детектируется отдельно (_move_note) — только когда МНОЖЕСТВО id до
    и после одинаково (иначе это insert_after/delete, не move), чтобы не
    плодить ложные записи о перемещении там, где на деле что-то вставили или
    удалили."""
    b_by_id = {b["id"]: b for b in before}
    a_by_id = {b["id"]: b for b in after}

    changed = []
    for i, blk in a_by_id.items():
        bb = b_by_id.get(i)
        text_before = _btext(bb) if bb else ""
        text_after = _btext(blk)
        note = _struct_note(bb, blk) if bb else None
        if bb is None or text_before != text_after or note:
            rec = {"id": i, "before": text_before, "after": text_after}
            if note:
                rec["note"] = note
            changed.append(rec)
    changed += [{"id": i, "before": _btext(b), "after": ""} for i, b in b_by_id.items() if i not in a_by_id]

    if set(b_by_id) == set(a_by_id):
        move = _move_note(before, after, a_by_id)
        if move:
            existing = next((c for c in changed if c["id"] == move["id"]), None)
            if existing:
                existing["note"] = f'{existing["note"]}; {move["note"]}' if existing.get("note") else move["note"]
            else:
                changed.append(move)
    return changed


_ID_FIELDS = {
    # Поля операции, которые называют id блока (контракт — см. _EDIT_PROMPT).
    # normalize и replace_all сюда не входят: normalize — путь rule, Редактора
    # там нет вовсе; replace_all документ-широкая ПО КОНТРАКТУ и id не несёт.
    "replace_text": ("id",), "set_text": ("id",), "insert_after": ("id",),
    "delete": ("id",), "move_after": ("id", "after"), "set_style": ("id",),
    "create_table": ("after",), "set_cell": ("id",),
}


def _out_of_lane(op, fragment_ids):
    """id первого поля операции, которое ссылается на блок ВНЕ fragment_ids
    (фрагмент, показанный Редактору в этом вызове, плюс id блоков, созданных
    ЭТИМ ЖЕ батчем — см. пополнение fragment_ids в _apply_ops), или None."""
    for field in _ID_FIELDS.get(op.get("op"), ()):
        val = op.get(field)
        if val is not None and val not in fragment_ids:
            return val
    return None


def _apply_ops(doc, idx, ops, fragment_text, request, editor, fragment_ids=None):
    """Применяет ops по очереди; на невалидной операции — один ретрай к
    Редактору с текстом ошибки (editor=None — ретрая нет, путь rule). Второй
    провал (или провал без Редактора) останавливает применение немедленно.
    Возвращает (описания применённых операций, причина отказа-или-None,
    tries) — tries=2, если к Редактору пришлось обращаться повторно
    (для журнала правки, поле "iter"), иначе 1.

    fragment_ids=None — гвард выключен (путь rule: editor=None, id-полей у
    normalize нет). Иначе — set id блоков фрагмента, который видел Редактор;
    операция вне этого множества отклоняется ДО patch.validate, тем же текстом
    ошибки и тем же единственным ретраем, что и любая другая невалидная
    операция (находка: validate звали с картой ВСЕГО документа, поэтому id
    любого существующего блока проходил, даже если Редактор его не видел).
    Успешно применённая операция пополняет fragment_ids id-ами, которые
    patch.apply мог только что создать (_register кладёт их прямо в idx) —
    иначе insert_after/create_table внутри своего же батча не смогли бы
    сослаться на блок, который сами только что создали."""
    applied, retried, i = [], False, 0
    while i < len(ops):
        op = ops[i]
        out = _out_of_lane(op, fragment_ids) if fragment_ids is not None else None
        err = (
            f"блок {out!r} вне фрагмента, показанного в этом вызове — правь только то, что было показано"
            if out is not None else patch.validate(doc_map(doc, idx), op, doc)
        )
        if err is None:
            before_ids = set(idx)
            applied.append(patch.apply(doc, idx, op))
            if fragment_ids is not None:
                fragment_ids |= set(idx) - before_ids
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


def _already(reason, nav, editor_reply=None, tries=1):
    """Находка Ф8 (item A) + Change C: rule-путь, ничего не изменивший
    (normalize дал пустой diff), — это НЕ провал сам по себе, но и не сразу
    "already": rule умеет отличить только "нормализация ничего не нашла" от
    "нормализация нашла и применила", а не "уже чисто" от "я не умею то, что
    попросили" (Математика edit 1 — двойной пробел и висящий пробел остались
    в документе, а rule смолчал, потому что искал не тот класс дефектов).
    Поэтому пустой diff у rule теперь restart'ит цикл обычным локальным путём
    (см. run_edit/_run_local), и already легитимен только если ЭТОТ путь тоже
    не нашёл, что менять (fallthrough=True в _run_local) — тогда rule и
    Редактор согласны, что менять нечего, и "failed" вводил бы в заблуждение
    не меньше ложного "done" (инвариант 5)."""
    return {
        "verdict": "already", "reason": reason, "applied": [], "ids": [],
        "iter": tries, "reply": {"nav": nav, "editor": editor_reply},
    }


def _restore_in_place(doc, idx, snapshot_path):
    """Восстанавливает doc/idx НА МЕСТЕ (мутирует переданные объекты, а не
    создаёт новую пару) и удаляет снимок. Нужна там, откуда нельзя вернуть
    вызывающему новую пару (result, doc, idx) обычным return — управление
    уходит через исключение (item D). Вызывающий держит СТАРУЮ ссылку на doc,
    и после падения Проверяющего она обязана снова стать нетронутой, а не
    домашним объектом с половиной правки внутри."""
    fresh = Document(snapshot_path)
    body = doc.element.body
    for child in list(body):
        body.remove(child)
    for child in list(fresh.element.body):
        body.append(child)
    idx.clear()
    idx.update(index(doc))
    os.remove(snapshot_path)


def _finish(doc, idx, snapshot_path, blocks_before, applied, request, checker, nav, editor_reply, tries):
    """Общий хвост ПОСЛЕ применения операций — общий и для одного вызова
    Редактора, и для нескольких (см. _run_clusters): diff по всему снимку →
    Проверяющий → вердикт → коммит/откат. Ровно один снимок, один diff, один
    вызов Проверяющего на всю правку — граница транзакции не дробится вместе
    с шагами (см. спецификацию задачи про кластеризацию). Возвращает
    (result, doc, idx, tries); result is None ТОЛЬКО когда diff пуст — смысл
    пустого diff разный на разных путях (item A/C), решение оставлено
    вызывающему, а doc/idx в этом случае уже НОВАЯ пара из _restore."""
    diff = _diff(blocks_before, doc_map(doc, idx))
    if not diff:
        doc, idx = _restore(snapshot_path)
        return None, doc, idx, tries

    try:
        verdict = checker(request, diff) or {}
    except Exception:
        # item D: транспорт (или что угодно другое) умер во время вызова Проверяющего —
        # патч уже применён к живому doc, снимок ещё жив. Молчать нельзя (инвариант 6),
        # но и оставлять doc мутированным-непроверенным нельзя (инвариант 5): восстанавливаем
        # ЭТИ ЖЕ объекты doc/idx на месте (см. _restore_in_place — обычный _restore тут не
        # годится, он возвращает НОВУЮ пару, а мы не можем сделать return из except) и
        # поднимаем исходное исключение дальше — громко, как и было.
        _restore_in_place(doc, idx, snapshot_path)
        raise

    ids_touched = [d["id"] for d in diff]
    reply = {"nav": nav, "editor": editor_reply}
    if verdict.get("ok"):
        os.remove(snapshot_path)
        return {"verdict": "done", "reason": verdict.get("reason", ""), "applied": applied, "ids": ids_touched,
                 "iter": tries, "reply": reply}, doc, idx, tries

    doc, idx = _restore(snapshot_path)
    return {"verdict": "rolled_back", "reason": verdict.get("reason", ""), "applied": applied, "ids": ids_touched,
             "iter": tries, "reply": reply}, doc, idx, tries


def _apply_and_check(doc, idx, ops, fragment_text, request, use_editor, nav, editor_reply, checker):
    """Путь одного вызова Редактора (только rule — normalize мимо кластеров,
    см. run_edit): снимок → применение операций → _finish. См. _finish за
    диффом/Проверяющим/вердиктом и _run_clusters за путём с кластерами
    (единственным путём для локальных правок через Редактора)."""
    snapshot_path = _snapshot(doc)
    blocks_before = doc_map(doc, idx)
    applied, err, tries = _apply_ops(doc, idx, ops, fragment_text, request, use_editor)

    if not applied:
        doc, idx = _restore(snapshot_path)
        return _failed(err or "ни одна операция не применилась", nav, editor_reply, tries), doc, idx, tries

    return _finish(doc, idx, snapshot_path, blocks_before, applied, request, checker, nav, editor_reply, tries)


def _cluster(blocks, ids, around=1):
    """Группирует resolved id по соседству в порядке документа: окно цели —
    она сама плюс around соседей с каждой стороны (та же идея, что и у
    find.fragment). Окна, которые пересекаются, сливаются в один кластер —
    два id в соседних абзацах остаются ОДНИМ шагом, а разбросанные по
    документу id расходятся по разным. Кластеры и id внутри них — в порядке
    документа. Позиции берутся из blocks на момент вызова (до начала цикла
    правок в _run_clusters) — порядок кластеров этим и фиксируется, дальше
    внутри цикла меняется только содержимое фрагмента, не разбиение."""
    pos = {b["id"]: i for i, b in enumerate(blocks)}
    ordered = sorted(ids, key=lambda i: pos[i])
    clusters, cur, cur_end = [], [], None
    for i in ordered:
        p = pos[i]
        if cur_end is not None and p - around <= cur_end:
            cur.append(i)
            cur_end = max(cur_end, p + around)
        else:
            if cur:
                clusters.append(cur)
            cur, cur_end = [i], p + around
    if cur:
        clusters.append(cur)
    return clusters


def _run_clusters(doc, idx, blocks, ids, request, editor, checker, nav, fallthrough):
    """Путь с несколькими маленькими шагами вместо одного большого фрагмента
    (измерено на 40 живых правках: Редактор отказывает по ОБЪЁМУ, а не по
    смыслу, когда целей много — edit с 4 целей до 10 упала done→rolled_back).
    Резолвленные id группируются в кластеры (_cluster), и каждый кластер
    получает СВОЙ вызов Редактора на СВЕЖЕМ doc_map (обязательно — предыдущие
    кластеры уже мутировали документ, старая карта дала бы Редактору "old",
    которого больше нет, и validate его отклонит).

    Транзакция и Проверяющий — по-прежнему ОДНИ на правку целиком (_finish):
    шаги дробятся, обязательство — нет. Кластер, где Редактор честно вернул
    пустые ops, — пропускается, не проваливает правку (другим кластерам ещё
    есть что делать); кластер, чьи ops не прошли валидацию даже после ретрая
    внутри _apply_ops, — прерывает ВЕСЬ цикл, откатывая единственный снимок."""
    snapshot_path = _snapshot(doc)
    blocks_before = doc_map(doc, idx)
    clusters = _cluster(blocks, ids, around=1)

    applied_all, editor_replies, tries_total = [], [], 0
    for cluster_ids in clusters:
        frag_blocks = find.fragment(doc_map(doc, idx), cluster_ids, around=1)
        fragment_text = render(frag_blocks)
        fragment_ids = {b["id"] for b in frag_blocks}
        reply = editor(fragment_text, request) or {}
        editor_replies.append(reply)
        tries_total += 1
        ops = _clean_ops(reply.get("ops"))
        if not ops:
            continue
        applied, err, cluster_tries = _apply_ops(doc, idx, ops, fragment_text, request, editor, fragment_ids)
        tries_total += cluster_tries - 1  # первый вызов кластера уже посчитан выше, тут доучитывается только ретрай
        applied_all += applied
        if err is not None:
            doc, idx = _restore(snapshot_path)
            return _failed(err, nav, editor_replies, tries_total), doc, idx

    if not applied_all:
        doc, idx = _restore(snapshot_path)
        if fallthrough:
            reason = "нормализация ничего не нашла, и Редактор ни в одном кластере не предложил операций — оба пути согласны, что менять нечего"
            return _already(reason, nav, editor_replies, tries_total), doc, idx
        return _failed("редактор не предложил операций ни в одном из кластеров", nav, editor_replies, tries_total), doc, idx

    result, doc, idx, tries = _finish(doc, idx, snapshot_path, blocks_before, applied_all, request, checker, nav, editor_replies, tries_total)
    if result is not None:
        return result, doc, idx
    return _failed("операции применились, но текст не изменился", nav, editor_replies, tries), doc, idx


def _run_local(doc, idx, blocks, nav, request, editor, checker, fallthrough=False):
    """Локальный путь через Редактора: резолв id → кластеризация → цикл
    маленьких шагов (_run_clusters, один вызов Редактора на кластер). Нет
    верхней границы на число целей и нет отдельной ветки для одной цели —
    один резолвленный id это тоже кластер, просто единственный, а
    единообразный цикл короче, чем branch на N==1 (владелец: обязательство
    транзакции не меняется от того, сколько шагов внутри). Используется и
    как обычный (non-rule) путь run_edit, и как fallthrough item C — когда
    rule ничего не изменил и цикл перезапускает ТОТ ЖЕ запрос обычным путём.
    fallthrough=True меняет только то, каким вердиктом закрывается "здесь
    нечего резолвить/Редактор честно отказался": already (rule и Редактор
    согласны — менять нечего), а не failed."""
    ids = _resolve(blocks, nav, request)
    if not ids:
        if fallthrough:
            verdict_fn = _already
            reason = "нормализация ничего не нашла, а обычный поиск (id/якоря/цитаты) тоже не нашёл, где это применить — похоже, менять действительно нечего"
        else:
            verdict_fn = _failed
            reason = "не нашёл, где это править: ни id и якоря Навигатора, ни дословные цитаты из текста правки не нашли места в документе"
        return verdict_fn(reason, nav), doc, idx

    return _run_clusters(doc, idx, blocks, ids, request, editor, checker, nav, fallthrough)


def run_edit(doc, idx, request, navigator=_navigate, editor=_edit_llm, checker=_check):
    """Один цикл правки. Возвращает (result, doc, idx) — после отката
    (verdict != "done" после применения) doc/idx уже перечитаны из снимка,
    старые ссылки вызывающего использовать нельзя.

    result содержит "reply": {"nav": ..., "editor": ...} — РАЗОБРАННЫЙ JSON-
    ответ ролей (llm.chat парсит его сам, сырое тело HTTP нигде не хранится).
    "editor": None, когда Редактора не звали вообще (путь rule и он ничего не
    нашёл менять); на локальном пути (_run_local/_run_clusters) "editor" —
    СПИСОК ответов, по одному на кластер (даже если резолвился один id и
    кластер тоже один) — иначе список из N ответов выглядел бы как один
    вызов, а их было N. "iter" — настоящее число обращений к Редактору за
    всю правку: по одному на кластер плюс по одному за каждый ретрай внутри
    _apply_ops (на пути rule, где Редактора не звали, — 1, как и раньше), не
    декоративная константа.

    Change C: пустой diff у rule — не терминал. Код восстанавливает документ
    и продолжает ТЕМ ЖЕ запросом обычным локальным путём (_run_local с
    fallthrough=True) — rule не умеет отличить "уже чисто" от "не умею это
    делать", это умеет только реальная попытка через Редактора."""
    blocks = doc_map(doc, idx)
    nav = navigator(find.outline(blocks), request) or {}
    rule = nav.get("rule")

    if not rule:
        return _run_local(doc, idx, blocks, nav, request, editor, checker)

    # Правки типа rule идут мимо Редактора — код нормализует сам, LLM не зовём.
    ops, fragment_text = [{"op": "normalize", "rule": rule}], ""
    result, doc, idx, _tries = _apply_and_check(doc, idx, ops, fragment_text, request, None, nav, None, checker)
    if result is not None:
        return result, doc, idx
    # doc/idx здесь уже НОВАЯ пара из _restore (см. _apply_and_check) — blocks
    # обязан быть пересчитан на них, старый blocks с прошлого doc использовать нельзя.
    return _run_local(doc, idx, doc_map(doc, idx), nav, request, editor, checker, fallthrough=True)


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
