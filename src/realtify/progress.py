from __future__ import annotations

from collections.abc import Callable


ProgressCallback = Callable[[str], None]


def emit_progress(callback: ProgressCallback | None, message: str) -> None:
    if callback:
        callback(message)
