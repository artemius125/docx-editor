"""Один вызов модели: POST на OpenAI-совместимый /chat/completions, разбор JSON.

Ключ и параметры — в .env в корне репозитория (не коммитится, парсим сами,
без новой зависимости). Абстракции провайдера и резервной модели здесь нет —
решено архитектором (BUILD_PLAN, Ф5): это последнее средство, а не сервис.

Ретрай — РОВНО один, и только на обрыве транспорта (httpx.TransportError:
ConnectError/ReadTimeout и подобные) — архитектор пересмотрел исходный запрет
на Ф7 по измеренным данным: 2 обрыва на ~21 живом вызове при прогоне
`colbert_run.py`, а полный прогон делает 40–60 вызовов — без ретрая он
физически не может закончиться. HTTP-ошибка (httpx.HTTPStatusError) и
не-JSON ответ по-прежнему падают мгновенно и громко — ретраится только
транспорт, никогда не то, что ответила модель.

На не-2xx ответе не используется raise_for_status(): реальный сбой
(400 от litellm, ContextWindowExceededError) показал, что его сообщение —
только статусная строка, тело с настоящей причиной выброшено. Поэтому
исключение собирается вручную с телом ответа (обрезанным) и размером
запроса — тип исключения тот же (httpx.HTTPStatusError), чтобы код,
ловящий его или httpx.TransportError по отдельности, не менялся.
"""

import json
from pathlib import Path

import httpx

_ENV_PATH = Path(__file__).resolve().parent.parent / ".env"


def _env():
    """Построчный разбор .env (KEY=VALUE, # — комментарий)."""
    values = {}
    for line in _ENV_PATH.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        values[key.strip()] = value.strip()
    return values


def chat(messages: list[dict]) -> dict:
    """Отправляет messages модели из .env, возвращает разобранный JSON-объект.

    Всегда JSON-режим и temperature=0.1 — так решено в плане. Падает громко:
    и на HTTP-ошибке (httpx.HTTPStatusError с телом ответа и размером
    запроса в сообщении), и если content не парсится как JSON
    (json.JSONDecodeError) — тихого запасного пути нет.
    """
    env = _env()
    url = env["LLM_BASE_URL"].rstrip("/") + "/chat/completions"
    kwargs = dict(
        headers={"Authorization": f"Bearer {env['AITUNNEL_KEY']}"},
        json={
            "model": env["LLM_MODEL"],
            "messages": messages,
            "temperature": 0.1,
            "response_format": {"type": "json_object"},
        },
        timeout=120,
    )
    try:
        r = httpx.post(url, **kwargs)
    except httpx.TransportError:
        r = httpx.post(url, **kwargs)  # один ретрай, только на обрыве транспорта
    if r.is_error:
        # raise_for_status() даёт только статусную строку и выбрасывает тело
        # ответа — на живом сбое (400, litellm) там лежала ЕДИНСТВЕННАЯ
        # причина (ContextWindowExceededError), а строка "400 Bad Request"
        # сама по себе ни о чём не говорит. 4000 символов — с запасом под
        # вложенный JSON litellm (само сообщение — десятки-сотни символов),
        # но не мегабайт, если сервер вдруг отдаст HTML-страницу ошибки.
        body = r.text
        if len(body) > 4000:
            body = body[:4000] + "…[тело обрезано, было {} симв.]".format(len(body))
        sent_chars = sum(len(m.get("content", "")) for m in messages)
        raise httpx.HTTPStatusError(
            f"{r.status_code} {r.reason_phrase} для {url} "
            f"(отправлено {sent_chars} знаков в messages): {body}",
            request=r.request,
            response=r,
        )
    content = r.json()["choices"][0]["message"]["content"]
    return json.loads(content)
