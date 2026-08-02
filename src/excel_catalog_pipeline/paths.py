"""Stable application-owned path resolution."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

APPLICATION_ID = "excel_catalog_pipeline"


def app_root() -> Path:
    return Path.home() / ".tkn" / APPLICATION_ID


def global_config_path() -> Path:
    return app_root() / "config.yaml"


def state_root() -> Path:
    return app_root() / "state"


def state_path() -> Path:
    return state_root() / "sync-state.json"


def runs_root() -> Path:
    return state_root() / "runs"


def cache_root() -> Path:
    if os.name == "nt":
        return Path.home() / ".cache" / APPLICATION_ID
    return Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / APPLICATION_ID


def temporary_root() -> Path:
    return Path(tempfile.gettempdir()) / APPLICATION_ID
