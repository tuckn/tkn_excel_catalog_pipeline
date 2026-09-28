"""Packaged generation instructions, separate from connection profiles."""

from __future__ import annotations

import hashlib
from importlib.resources import files

from .context_source import ContextError
from .models import ContextConfig


def profile_name(config: ContextConfig) -> str:
    if config.prompt_profile != "auto":
        return config.prompt_profile
    return "default-en" if config.language.lower() in {"english", "en"} else "default-ja"


def load_prompt(config: ContextConfig, *, stage: str = "sheet") -> str:
    name = profile_name(config)
    if name not in {"default-ja", "default-en"}:
        raise ContextError(f"Unknown context profile: {name}; choose default-ja or default-en")
    if stage not in {"sheet", "workbook"}:
        raise ContextError(f"Unknown generation stage: {stage}")
    filename = "prompt.md" if stage == "sheet" else "workbook-prompt.md"
    text = (
        files("excel_catalog_pipeline")
        .joinpath("context_profiles", name, filename)
        .read_text(encoding="utf-8")
    )
    language = (
        config.language
        if config.prompt_profile == "auto"
        else ("English" if name == "default-en" else "Japanese")
    )
    return text.replace("{{language}}", language)


def prompt_digest(config: ContextConfig) -> str:
    return hashlib.sha256(load_prompt(config).encode()).hexdigest()
