"""Две детерминированные нормализации текста, без модели: typography и quotes.

Правят живой документ на месте через _ptext/_replace_span из patch.py — ту же
машинерию, что использует replace_all (раны, гиперссылки, ячейки таблиц), чтобы
не заводить второй механизм редактирования текста.
"""

import re

from docx.oxml.ns import qn

from docx_editor.patch import _ptext, _replace_span

# Пробел (включая U+00A0 и другие юникодные пробелы) перед пунктуацией.
_WS_BEFORE_PUNCT = re.compile(r"[^\S\n]+(?=[,.:;!?)\]}»])")

# Пропущенный пробел между предложениями: буква любого регистра + точка + заглавная.
# Защита от ложных срабатываний на десятичных числах (4.2, 0.4448) — не регистр
# слева, а то, что классы букв не включают цифры: цифра по обе стороны от точки
# в принципе не может совпасть с [A-Za-zА-Яа-яЁё] / [A-ZА-ЯЁ].
_GLUED_SENTENCE = re.compile(r"([A-Za-zА-Яа-яЁё])\.([A-ZА-ЯЁ])")

# Невидимые форматирующие символы: word joiner, zero-width space, BOM, мягкий перенос.
_INVISIBLE = re.compile(r"[⁠​﻿­]+")


def _paragraphs(doc):
    """Все w:p документа, включая абзацы внутри ячеек таблиц (как replace_all)."""
    return doc.element.body.iter(qn("w:p"))


def typography(doc):
    """Три правила по всему документу: пробел перед пунктуацией, слипшиеся
    предложения, невидимые символы. Правки внутри абзаца — справа налево,
    иначе более ранние замены сдвигают смещения более поздних."""
    ws = glued = invisible = 0
    for p_el in _paragraphs(doc):
        text = _ptext(p_el)
        matches = list(_WS_BEFORE_PUNCT.finditer(text))
        for m in reversed(matches):
            _replace_span(p_el, m.start(), m.end(), "")
        ws += len(matches)

        text = _ptext(p_el)
        matches = list(_GLUED_SENTENCE.finditer(text))
        for m in reversed(matches):
            _replace_span(p_el, m.start(), m.end(), f"{m.group(1)}. {m.group(2)}")
        glued += len(matches)

        text = _ptext(p_el)
        matches = list(_INVISIBLE.finditer(text))
        for m in reversed(matches):
            _replace_span(p_el, m.start(), m.end(), "")
        invisible += len(matches)

    return (
        f"Типография: убрано пробелов перед пунктуацией — {ws}, "
        f"разделено слипшихся предложений — {glued}, "
        f"вырезано невидимых символов — {invisible}"
    )


def quotes(doc):
    """Прямые " заменяются на ёлочки: чётные по порядку в абзаце — «, нечётные — ».

    В документе вложенных кавычек внутри ёлочек нет, число прямых кавычек в
    каждом абзаце чётно, поэтому простое чередование в порядке появления даёт
    корректную расстановку открывающих/закрывающих.
    """
    replaced = 0
    for p_el in _paragraphs(doc):
        text = _ptext(p_el)
        positions = [i for i, ch in enumerate(text) if ch == '"']
        for order, pos in reversed(list(enumerate(positions))):
            char = "«" if order % 2 == 0 else "»"
            _replace_span(p_el, pos, pos + 1, char)
        replaced += len(positions)

    return f"Кавычки: заменено на ёлочки — {replaced}"
