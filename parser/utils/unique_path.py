from pathlib import Path


def unique_path(parent: Path, name: str) -> Path:
    """Возвращает уникальный путь, добавляя счётчик при коллизии"""
    candidate = parent / name
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = parent / f"{name} ({suffix})"
        if not candidate.exists():
            return candidate
        suffix += 1
