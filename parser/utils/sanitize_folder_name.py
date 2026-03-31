import re
from pathlib import Path


def sanitize_folder_name(name: str) -> str:
    forbidden_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(forbidden_chars, '', name)
    sanitized = sanitized.replace('"', '')
    sanitized = sanitized.replace('«', '')
    sanitized = sanitized.replace('»', '')

    return sanitized
