"""VisualTeX private OCR Python isolation guard.

This runtime is application-owned. User-level Python packages must never
participate in dependency resolution or imports.
"""

from __future__ import annotations

import os
import site
import sys

os.environ["PYTHONNOUSERSITE"] = "1"

try:
    configured = site.getusersitepackages()
    user_sites = [configured] if isinstance(configured, str) else list(configured)
except Exception:
    user_sites = []


def _normalized(value: str) -> str:
    return os.path.normcase(os.path.abspath(value))


normalized_user_sites = [_normalized(path) for path in user_sites if path]
sys.path[:] = [
    path
    for path in sys.path
    if not any(
        _normalized(path) == user_site
        or _normalized(path).startswith(user_site + os.sep)
        for user_site in normalized_user_sites
    )
]
site.ENABLE_USER_SITE = False
