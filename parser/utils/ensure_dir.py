import sys
from pathlib import Path
from .logs import log_ok, log_err, log_warn


def ensure_dir(path: Path, label: str) -> None:
    """Проверяет существование папки, при отсутствии создаёт"""
    if path.is_dir():
        log_ok(f"Найдена папка          → {label}")
    else:
        log_warn(f"Папка не найдена       → {label}")
        try:
            path.mkdir(parents=True, exist_ok=True)
            log_ok(f"Создана папка          → {label}")
        except OSError as e:
            log_err(f"Не удалось создать папку {label}: {e}")
            sys.exit(1)
