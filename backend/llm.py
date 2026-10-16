"""Minimal async LLM client.

Replaces CrewAI, which wrapped a single agent call six times per user request.
One HTTP POST, one retry policy, one timeout — no env mutation, no shared state.
"""

import asyncio
import os

import httpx

TIMEOUT = httpx.Timeout(connect=10.0, read=180.0, write=30.0, pool=10.0)
MAX_RETRIES = 2
RETRY_STATUS = {408, 409, 425, 429, 500, 502, 503, 504}

_client: httpx.AsyncClient | None = None


def client() -> httpx.AsyncClient:
    global _client
    if _client is None or _client.is_closed:
        _client = httpx.AsyncClient(timeout=TIMEOUT)
    return _client


async def aclose() -> None:
    global _client
    if _client is not None and not _client.is_closed:
        await _client.aclose()
    _client = None


def _split(system: str, user: str) -> tuple[str, str]:
    return system or "", user or ""


def _payload_openai(model: str, system: str, user: str, max_tokens: int) -> dict:
    return {
        "model": model,
        "messages": (
            ([{"role": "system", "content": system}] if system else [])
            + [{"role": "user", "content": user}]
        ),
        "max_tokens": max_tokens,
        "temperature": 0,
    }


def _payload_anthropic(model: str, system: str, user: str, max_tokens: int) -> dict:
    return {
        "model": model,
        "max_tokens": max_tokens,
        "temperature": 0,
        "system": system or None,
        "messages": [{"role": "user", "content": user}],
    }


def _payload_google(model: str, system: str, user: str, max_tokens: int) -> dict:
    name = model.replace("gemini-", "").replace("models/", "")
    return {
        "contents": [{"role": "user", "parts": [{"text": user}]}],
        "systemInstruction": {"parts": [{"text": system}]} if system else None,
        "generationConfig": {
            "temperature": 0,
            "maxOutputTokens": max_tokens,
        },
    }


def _extract(provider: str, data: dict) -> str:
    if provider == "anthropic":
        return "".join(
            b.get("text", "") for b in data.get("content", []) if b.get("type") == "text"
        )
    if provider == "google":
        parts = (
            data.get("candidates", [{}])[0].get("content", {}).get("parts", [])
        )
        return "".join(p.get("text", "") for p in parts)
    choices = data.get("choices") or []
    if not choices:
        return ""
    return (choices[0].get("message") or {}).get("content") or ""


async def call_llm(
    provider: str,
    model: str,
    api_key: str,
    system: str,
    user: str,
    max_tokens: int = 4096,
) -> str:
    """One LLM call. Returns text, or raises RuntimeError with the provider error."""
    if not api_key:
        raise RuntimeError(f"No API key configured for provider '{provider}'")
    if not model:
        raise RuntimeError(f"No model configured for provider '{provider}'")

    system, user = _split(system, user)
    params: dict | None = None

    if provider == "openai":
        url = "https://api.openai.com/v1/chat/completions"
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
        body = _payload_openai(model, system, user, max_tokens)
    elif provider == "anthropic":
        url = "https://api.anthropic.com/v1/messages"
        headers = {
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        }
        body = _payload_anthropic(model, system, user, max_tokens)
    elif provider == "google":
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model.replace('gemini-', '').replace('models/', '')}:generateContent"
        headers = {"Content-Type": "application/json"}
        body = _payload_google(model, system, user, max_tokens)
        params = {"key": api_key}
    elif provider == "openrouter":
        url = os.environ.get(
            "CEL_OPENROUTER_ENDPOINT", "https://openrouter.ai/api/v1/chat/completions"
        )
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
            "HTTP-Referer": os.environ.get("CEL_APP_URL", "https://cel.excel"),
            "X-Title": "Cel",
        }
        body = _payload_openai(model, system, user, max_tokens)
        params = None
    else:
        # Unknown provider: fall back to OpenAI-compatible, which covers most proxies
        url = os.environ.get(
            "CEL_OPENAI_COMPAT_ENDPOINT", "https://api.openai.com/v1/chat/completions"
        )
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
        body = _payload_openai(model, system, user, max_tokens)
        params = None

    last_error = ""
    for attempt in range(MAX_RETRIES + 1):
        try:
            resp = await client().post(
                url, headers=headers, json=body, params=params
            )
        except httpx.HTTPError as exc:
            last_error = f"{provider} request failed: {exc}"
            if attempt < MAX_RETRIES:
                await asyncio.sleep(1.5 * (attempt + 1))
                continue
            raise RuntimeError(last_error) from exc

        if resp.status_code in RETRY_STATUS and attempt < MAX_RETRIES:
            await asyncio.sleep(1.5 * (attempt + 1))
            continue

        if not resp.is_success:
            raise RuntimeError(
                f"{provider} API error {resp.status_code}: {resp.text[:600]}"
            )

        try:
            text = _extract(provider, resp.json())
        except Exception as exc:
            raise RuntimeError(f"{provider} returned unparseable response: {exc}") from exc

        if not text.strip():
            raise RuntimeError(f"{provider} returned an empty response")
        return text

    raise RuntimeError(last_error or f"{provider} request failed after retries")
