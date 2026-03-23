import os
import sys
import re
import zipfile
import shutil
from pathlib import Path
import pandas as pd
from utils.logs import log_ok, log_err, log_info, log_warn, section
from utils.fix_filename_encoding import fix_filename_encoding
from utils.unique_path import unique_path


# ═══════════════════════════════════════════════════════════════
#  КОНФИГУРАЦИЯ
# ═══════════════════════════════════════════════════════════════

ZIP_DIR = "МИНСОЦ"
EXTRACT_DIR = "Временные файлы"
PROCESSED_DIR = "Обработанные архивы"
FILES_DIR = "Файлы"

EXCEL_SKIP_ROWS = 6
FOLDER_COL_INDEX = 0
FILE_COL_INDEX = 18


# ═══════════════════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════════


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


# ═══════════════════════════════════════════════════════════════
#  ПУТИ
# ═══════════════════════════════════════════════════════════════

base_path = Path.cwd()
zip_path = base_path.parent / ZIP_DIR
extract_path = base_path / EXTRACT_DIR
processed_path = zip_path / PROCESSED_DIR
files_path = zip_path / FILES_DIR


def main():

    # ═══════════════════════════════════════════════════════════════
    #  1. ПРОВЕРКА / СОЗДАНИЕ ФАЙЛОВОЙ СТРУКТУРЫ
    # ═══════════════════════════════════════════════════════════════

    section("Проверка файловой структуры")

    if not zip_path.is_dir():
        log_err(f"Исходная папка не найдена → {ZIP_DIR}")
        sys.exit(1)
    log_ok(f"Найдена исходная папка → {ZIP_DIR}")

    ensure_dir(extract_path,   EXTRACT_DIR)
    ensure_dir(processed_path, PROCESSED_DIR)
    ensure_dir(files_path,     FILES_DIR)

    # ═══════════════════════════════════════════════════════════════
    #  2. ПОИСК И РАСПАКОВКА АРХИВА
    # ═══════════════════════════════════════════════════════════════

    section("Обработка архива")

    zip_files = list(zip_path.glob("*.zip"))

    if not zip_files:
        log_err("В папке не найдено ни одного .zip архива")
        sys.exit(1)

    if len(zip_files) > 1:
        log_warn(
            f"Найдено несколько архивов ({len(zip_files)}), обрабатывается первый")

    zip_file_path = zip_files[0]
    zip_filename = zip_file_path.name
    log_ok(f"Найден архив           → {zip_filename}")

    # Распаковка
    try:
        with zipfile.ZipFile(zip_file_path, "r") as zf:
            zf.extractall(extract_path)
        log_ok(f"Распакован в           → {extract_path}")
    except zipfile.BadZipFile:
        log_err(f"Файл повреждён или не является ZIP-архивом: {zip_filename}")
        sys.exit(1)
    except PermissionError as e:
        log_err(f"Нет прав для распаковки: {e}")
        sys.exit(1)

    # Исправление кодировки имён
    renamed_count = 0
    for root, dirs, files in os.walk(extract_path):
        root_path = Path(root)

        for d in list(dirs):
            new_name = fix_filename_encoding(d)
            if new_name != d:
                src = root_path / d
                dest = unique_path(root_path, new_name)
                try:
                    src.rename(dest)
                    log_info(f"Переименована папка    → {d}  →  {dest.name}")
                    dirs[dirs.index(d)] = dest.name
                    renamed_count += 1
                except OSError as e:
                    log_warn(f"Не удалось переименовать папку {d}: {e}")

        for f in files:
            new_name = fix_filename_encoding(f)
            if new_name != f:
                src = root_path / f
                dest = unique_path(root_path, new_name)
                try:
                    src.rename(dest)
                    log_info(f"Переименован файл      → {f}  →  {dest.name}")
                    renamed_count += 1
                except OSError as e:
                    log_warn(f"Не удалось переименовать файл {f}: {e}")

    if renamed_count:
        log_ok(f"Исправлено имён        → {renamed_count}")

    # Перемещение архива в «Обработанные»
    try:
        shutil.move(str(zip_file_path), processed_path)
        log_ok(f"Архив перемещён        → {PROCESSED_DIR}")
    except (shutil.Error, OSError) as e:
        log_warn(f"Не удалось переместить архив: {e}")

    # ═══════════════════════════════════════════════════════════════
    #  3. ЧТЕНИЕ EXCEL-ИНСТРУКЦИИ
    # ═══════════════════════════════════════════════════════════════

    section("Чтение Excel-инструкции")

    excel_files = list(extract_path.glob("*.xlsx"))

    if not excel_files:
        log_err(f"В папке '{EXTRACT_DIR}' не найдено ни одного .xlsx файла")
        sys.exit(1)

    if len(excel_files) > 1:
        log_warn(
            f"Найдено несколько .xlsx ({len(excel_files)}), используется первый")

    excel_path = excel_files[0]
    excel_filename = excel_path.name
    log_ok(f"Найден файл инструкций → {excel_filename}")

    # Извлекаем дату из имени ZIP-архива
    m = re.search(r"(\d{2}[.\-_]\d{2}[.\-_]\d{4})", zip_filename)
    if not m:
        log_err(f"Дата не найдена в имени архива: {zip_filename}")
        sys.exit(1)

    date_str = re.sub(r"[._\-]", ".", m.group(1))
    date_dir = files_path / date_str
    date_dir.mkdir(parents=True, exist_ok=True)
    log_ok(f"Папка для файлов       → {date_str}")

    # Читаем таблицу
    try:
        xls = pd.ExcelFile(excel_path)
        sheet_name = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name,
                           skiprows=EXCEL_SKIP_ROWS)
    except FileNotFoundError:
        log_err(f"Файл не найден: {excel_filename}")
        sys.exit(1)
    except Exception as e:
        log_err(f"Ошибка чтения Excel: {e}")
        sys.exit(1)

    required_cols = max(FOLDER_COL_INDEX, FILE_COL_INDEX) + 1
    if df.shape[1] < required_cols:
        log_err(
            f"В таблице {df.shape[1]} столбцов, "
            f"необходимо минимум {required_cols} (A–S)"
        )
        sys.exit(1)

    log_ok(
        f"Таблица загружена      → {df.shape[0]} строк, {df.shape[1]} столбцов")

    # ═══════════════════════════════════════════════════════════════
    #  4. СОЗДАНИЕ ЦЕЛЕВЫХ ПАПОК
    # ═══════════════════════════════════════════════════════════════

    section("Подготовка папок")

    unique_folders: list[str] = [
        str(val).strip()
        for val in df.iloc[:, FOLDER_COL_INDEX].unique()
        if pd.notna(val) and str(val).strip() and str(val).strip().lower() != "nan"
    ]

    if not unique_folders:
        log_err("Не найдено ни одного непустого имени папки в столбце A")
        sys.exit(1)

    for folder in unique_folders:
        try:
            (date_dir / folder).mkdir(exist_ok=True)
        except OSError as e:
            log_warn(f"Не удалось создать папку '{folder}': {e}")

    log_ok(f"Папок подготовлено     → {len(unique_folders)}")

    # ═══════════════════════════════════════════════════════════════
    #  5. ИНДЕКСИРОВАНИЕ И ПЕРЕНОС ФАЙЛОВ
    # ═══════════════════════════════════════════════════════════════

    section("Перенос файлов")

    log_info("Индексирую файлы в папке 'Временные файлы'...")
    file_index: dict[str, Path] = {}
    for f in extract_path.rglob("*"):
        if f.is_file() and f.name not in file_index:
            file_index[f.name] = f

    log_ok(f"Проиндексировано файлов → {len(file_index)}")
    print()

    moved = 0
    not_found = 0
    skipped = 0

    for _, row in df.iterrows():
        raw_folder = row.iloc[FOLDER_COL_INDEX]
        raw_file = row.iloc[FILE_COL_INDEX]

        folder_name = str(raw_folder).strip() if pd.notna(raw_folder) else ""
        file_name = str(raw_file).strip() if pd.notna(raw_file) else ""

        if not folder_name or folder_name.lower() == "nan":
            skipped += 1
            continue
        if not file_name or file_name.lower() == "nan":
            skipped += 1
            continue

        target_folder = date_dir / folder_name

        if not target_folder.exists():
            log_warn(
                f"Папка не существует    → {folder_name}  (пропуск: {file_name})")
            not_found += 1
            continue

        if file_name not in file_index:
            log_err(f"Файл не найден         → {file_name}")
            not_found += 1
            continue

        src_file = file_index[file_name]
        dest_file = unique_path(target_folder, file_name)

        try:
            shutil.move(str(src_file), str(dest_file))
            del file_index[file_name]
            log_ok(
                f"Перенесён              → {file_name:<40} →  {folder_name}/")
            moved += 1
        except (shutil.Error, OSError) as e:
            log_err(f"Ошибка переноса {file_name}: {e}")
            not_found += 1

    # ═══════════════════════════════════════════════════════════════
    #  6. ПЕРЕМЕЩЕНИЕ EXCEL-ФАЙЛА В ПАПКУ С ДАТОЙ
    # ═══════════════════════════════════════════════════════════════

    section("Перемещение")

    # try:
    #     excel_path.unlink()
    #     log_ok(f"Удалён файл инструкций → {excel_filename}")
    # except OSError as e:
    #     log_warn(f"Не удалось удалить {excel_filename}: {e}")

    try:
        # Если Excel-файл ещё не был перемещён (вдруг он совпадает с каким-то файлом из списка)
        if excel_path.exists():
            dest_excel = unique_path(date_dir, excel_filename)
            shutil.move(str(excel_path), str(dest_excel))
            log_ok(f"Excel-файл перемещён    → {date_str}/{excel_filename}")
        else:
            log_warn("Excel-файл уже был перемещён или не найден")
    except (shutil.Error, OSError) as e:
        log_warn(f"Не удалось переместить Excel-файл: {e}")

    # ═══════════════════════════════════════════════════════════════
    #  7. ИТОГ
    # ═══════════════════════════════════════════════════════════════

    print(f"\n{'═' * 55}")
    print(f"  ИТОГ")
    print(f"{'═' * 55}")
    log_ok(f"Перенесено файлов      → {moved}")
    if skipped:
        log_info(f"Пропущено строк        → {skipped}")
    if not_found:
        log_err(f"Не найдено файлов      → {not_found}")
    print(f"{'═' * 55}\n")


if __name__ == '__main__':
    main()
