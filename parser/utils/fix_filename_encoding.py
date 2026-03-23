def fix_filename_encoding(name: str) -> str:
    """Исправляет кодировку имён файлов cp437 → cp866"""
    try:
        raw = name.encode("cp437")
    except UnicodeEncodeError:
        return name
    try:
        decoded = raw.decode("cp866")
        if decoded != name:
            return decoded
    except UnicodeDecodeError:
        pass
    return name
