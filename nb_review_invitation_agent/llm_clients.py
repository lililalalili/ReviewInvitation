from __future__ import annotations


def generate_json(*, fake_providers: bool, prompt: str, schema: dict) -> dict:
    if fake_providers:
        return {"provider": "fake", "ok": True, "prompt_len": len(prompt), "schema_keys": sorted(schema.keys())}
    raise RuntimeError("Real providers are not wired in baseline refactor. Use --fake-providers for offline runs.")
