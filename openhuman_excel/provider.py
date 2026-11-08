from __future__ import annotations

import httpx


class ProviderError(RuntimeError):
    pass


class OpenAICompatibleProvider:
    """Small provider adapter for DeepSeek and other OpenAI-compatible endpoints."""

    def __init__(self, base_url: str, api_key: str, model: str, timeout: int = 90):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.model = model
        self.timeout = timeout

    def complete(self, system: str, user: str) -> str:
        if not self.api_key:
            raise ProviderError("No API key configured. Add OPENH_EXCEL_API_KEY to .env.")
        url = f"{self.base_url}/v1/chat/completions"
        try:
            response = httpx.post(url, headers={"Authorization": f"Bearer {self.api_key}"},
                json={"model": self.model, "temperature": 0.1,
                      "messages": [{"role": "system", "content": system},
                                   {"role": "user", "content": user}]},
                timeout=self.timeout)
            response.raise_for_status()
            return response.json()["choices"][0]["message"]["content"]
        except (httpx.HTTPError, KeyError, IndexError, ValueError) as exc:
            raise ProviderError(f"Model request failed: {type(exc).__name__}: {exc}") from exc
