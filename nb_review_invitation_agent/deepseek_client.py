from __future__ import annotations

import json
import urllib.request
from dataclasses import dataclass


@dataclass(frozen=True)
class DeepSeekConfig:
    api_key: str
    base_url: str = "https://api.deepseek.com"
    model: str = "deepseek-v4-pro"
    timeout_seconds: int = 20


class DeepSeekClient:
    """OpenAI-compatible chat completion wrapper for DeepSeek."""

    def __init__(self, config: DeepSeekConfig):
        if not config.api_key:
            raise ValueError("DEEPSEEK_API_KEY is required for live DeepSeek usage.")
        self._config = config

    def generate_json(self, *, system_prompt: str, user_prompt: str) -> dict:
        payload = {
            "model": self._config.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "response_format": {"type": "json_object"},
            "temperature": 0,
        }
        req = urllib.request.Request(
            url=f"{self._config.base_url.rstrip('/')}/chat/completions",
            method="POST",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {self._config.api_key}",
                "Content-Type": "application/json",
            },
        )
        with urllib.request.urlopen(req, timeout=self._config.timeout_seconds) as response:
            data = json.loads(response.read().decode("utf-8"))
        content = data["choices"][0]["message"]["content"]
        return json.loads(content)


class FakeDeepSeekClient:
    def __init__(self, response: dict | Exception):
        self._response = response
        self.last_system_prompt = ""
        self.last_user_prompt = ""

    def generate_json(self, *, system_prompt: str, user_prompt: str) -> dict:
        self.last_system_prompt = system_prompt
        self.last_user_prompt = user_prompt
        if isinstance(self._response, Exception):
            raise self._response
        return dict(self._response)
