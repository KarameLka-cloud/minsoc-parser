import zipfile
import shutil
import os
from pathlib import Path

ZIP_PATH = "/Users/valentin/Documents/projects/parser/ахивы/МФЦ_ЦНМСП_25.02.2026.zip"
EXTRACT_TO = "/Users/valentin/Documents/projects/parser/файлы"
PROCESSED_DIR = "/Users/valentin/Documents/projects/parser/обработанные архивы"


def fix_filename_encoding(name: str) -> str:
    """Попытка восстановить читаемое имя, если оно неправильно раскодировано.

    Файлы из ZIP на Windows часто хранятся в cp437, а названия на русском
    могут быть закодированы в cp866. ZipFile декодирует их в Unicode с помощью
    cp437, поэтому мы пробуем перекодировать обратно и декодировать как cp866.
    Если это дает отличающуюся строку, возвращаем её. В противном случае
    возвращаем оригинал.
    """
    try:
        raw = name.encode('cp437')
    except UnicodeEncodeError:
        # Символы вне cp437 -> нечего исправлять
        return name

    try:
        decoded = raw.decode('cp866')
        if decoded != name:
            return decoded
    except UnicodeDecodeError:
        pass

    # можно добавить другие попытки, например cp1251, но cp866 чаще всего
    return name


def unique_path(parent: Path, name: str) -> Path:
    """Собираем уникальный путь, если в каталоге уже есть объект с таким
    именем."""
    candidate = parent / name
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = parent / f"{name} ({suffix})"
        if not candidate.exists():
            return candidate
        suffix += 1


def process_zip(zip_path: str, extract_to: str, processed_dir: str = "обработанные") -> bool:
    # Преобразуем в объекты Path для удобства
    zip_path = Path(zip_path)
    extract_to = Path(extract_to)
    processed_dir = Path(processed_dir)

    # Проверяем, что файл вообще существует
    if not zip_path.is_file():
        print(f"Ошибка: файл {zip_path} не найден")
        return False

    # Создаём папку для распаковки, если её нет
    extract_to.mkdir(parents=True, exist_ok=True)

    # Создаём папку "обработанные", если отсутствует
    processed_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Распаковываем архив
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
        print(f"Распаковано в → {extract_to}")

        # После распаковки иногда имена файлов выглядят как \"ÅπΓ...\" из-за
        # несовпадения кодировок. Постобработка переименует их обратно.
        for root, dirs, files in os.walk(extract_to):
            # переименовываем сначала папки, чтобы не сломать пути к файлам внутри
            for d in list(dirs):
                new_name = fix_filename_encoding(d)
                if new_name != d:
                    src = Path(root) / d
                    dest = unique_path(Path(root), new_name)
                    print(
                        f"Переименовываю директорию: {src.name} -> {dest.name}")
                    src.rename(dest)
                    dirs[dirs.index(d)] = dest.name
            for f in files:
                new_name = fix_filename_encoding(f)
                if new_name != f:
                    src = Path(root) / f
                    dest = unique_path(Path(root), new_name)
                    print(f"Переименовываю файл: {src.name} -> {dest.name}")
                    src.rename(dest)

        # Переносим оригинальный архив в папку обработанных
        destination = processed_dir / zip_path.name
        shutil.move(zip_path, destination)
        print(f"Архив перемещён → {destination}")

        return True

    except zipfile.BadZipFile:
        print("Ошибка: файл повреждён или не является zip-архивом")
    except PermissionError:
        print("Ошибка прав доступа")
    except Exception as e:
        print(f"Неизвестная ошибка: {e}")

    return False


if __name__ == "__main__":
    # Запускаем процесс с указанными путями
    process_zip(ZIP_PATH, EXTRACT_TO, PROCESSED_DIR)
