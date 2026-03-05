import sys
import pandas as pd
import shutil
from pathlib import Path

# ────────────────────────────────────────────────
# Настройки
# ────────────────────────────────────────────────

sheet_name = 'Лист1'                     # ← измени, если лист называется иначе

SOURCE_FOLDER = "файлы"                   # папка со всеми файлами
TARGET_FOLDER_NAME = "обработанные файлы"       # куда складываем результат

# ────────────────────────────────────────────────

base_dir = Path.cwd()
source_dir = base_dir / SOURCE_FOLDER
target_root = base_dir / TARGET_FOLDER_NAME

# Проверка существования папки files
if not source_dir.is_dir():
    print(f"✗ Папка '{SOURCE_FOLDER}' не найдена в текущей директории!")
    print("Ожидаемая структура:")
    print(f"  {base_dir.name}/")
    print(f"  ├── {SOURCE_FOLDER}/")
    print(f"  │   ├── ваш_файл.xlsx          ← здесь")
    print(f"  │   ├── photo1.jpg")
    print(f"  │   ├── document.pdf")
    print(f"  │   └── ... (все файлы, которые нужно рассортировать)")
    print(f"  └── sort.py (или как назван скрипт)")
    exit(1)

# Поиск единственного .xlsx файла внутри папки files
excel_files = list(source_dir.glob("*.xlsx"))

if len(excel_files) == 0:
    print(f"✗ В папке '{SOURCE_FOLDER}' не найден ни один файл .xlsx")
    exit(1)

if len(excel_files) > 1:
    print(f"✗ В папке '{SOURCE_FOLDER}' найдено несколько .xlsx файлов.")
    print("Оставьте только один:")
    for f in excel_files:
        print(f"  • {f.name}")
    exit(1)

excel_path = excel_files[0]
excel_filename = excel_path.name

print(f"Найден файл с инструкциями: {excel_filename}")
print(f"Папка с файлами:              {source_dir}\n")

# Создаём целевую папку
target_root.mkdir(exist_ok=True)

print(f"Результат будет в: {target_root}\n")

# Читаем Excel (пропускаем 6 первых строк)
try:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=6)
except Exception as e:
    print("Ошибка чтения Excel-файла:", e)
    exit(1)

if df.shape[1] <= 18:
    print("Ошибка: в таблице меньше 19 столбцов (нужны хотя бы A–S)")
    exit(1)

# Столбцы A (0) → имя папки, S (18) → имя файла
a_column = df.iloc[:, 0]
s_column = df.iloc[:, 18]

# Уникальные имена папок (убираем NaN и пустые строки)
unique_folders = [
    str(val).strip()
    for val in a_column.unique()
    if pd.notna(val) and str(val).strip()
]

if not unique_folders:
    print("Не найдено ни одного непустого имени папки в столбце A")
    exit(1)

# Создаём папки
for folder in unique_folders:
    (target_root / folder).mkdir(exist_ok=True)

print(f"Подготовлено папок: {len(unique_folders)}\n")

# ─── Сортировка ────────────────────────────────────────────────

moved = 0
not_found = 0
duplicates_handled = 0

for _, row in df.iterrows():
    folder_name = str(row.iloc[0]).strip()
    file_name = str(row.iloc[18]).strip()

    if not folder_name or not file_name:
        continue

    source_path = source_dir / file_name

    if not source_path.is_file():
        if source_path.name == excel_filename:
            # пропускаем сам Excel-файл, его трогать не нужно
            continue
        print(f"✗ Не найден файл: {file_name}")
        not_found += 1
        continue

    dest_folder = target_root / folder_name
    dest_path = dest_folder / file_name

    # Обработка дубликатов (добавляем _1, _2...)
    if dest_path.exists():
        stem, suffix = dest_path.stem, dest_path.suffix
        counter = 1
        while (dest_folder / f"{stem}_{counter}{suffix}").exists():
            counter += 1
        new_name = f"{stem}_{counter}{suffix}"
        dest_path = dest_folder / new_name
        print(f"  (дубликат → {new_name}) ", end="")
        duplicates_handled += 1

    shutil.move(str(source_path), str(dest_path))
    print(f"✓ {file_name} → {folder_name}")
    moved += 1

# ─── Итог ──────────────────────────────────────────────────────

print("\n" + "─" * 70)
print("Результат:")
print(f"  Перемещено файлов     : {moved}")
print(f"  Не найдено файлов     : {not_found}")
print(f"  Обработано дубликатов : {duplicates_handled}")
print(f"  Куда всё сложилось    : {target_root}")
print("─" * 70)

if moved == 0 and not_found > 0:
    print("Возможные причины проблем:")
    print(" • имена файлов в столбце S не совпадают с реальными (регистр, пробелы, расширения)")
    print(" • файлы лежат не в папке 'files'")
    print(" • в столбце S указан сам Excel-файл (его скрипт пропускает)")

# ─── Удаление исходного Excel-файла в любом случае ────────────
try:
    excel_path.unlink()
    print(f"→ Исходный файл {excel_filename} удалён")
except Exception as e:
    print(f"⚠ Не удалось удалить {excel_filename}: {e}")


# """Скрипт удаляет символы ", «, » из имён папок внутри папки sorted_files.

# Запуск в корне проекта:
#     python3 rename_dirs.py
# """

# ROOT = Path(__file__).parent
# TARGET = ROOT / "sorted_files"

# if not TARGET.exists():
#     print(f"Папка не найдена: {TARGET}")
#     sys.exit(1)

# BAD = ['"', '«', '»']


# def clean_name(name: str) -> str:
#     for ch in BAD:
#         name = name.replace(ch, '')
#     return name.strip()


# def unique_path(parent: Path, name: str) -> Path:
#     candidate = parent / name
#     if not candidate.exists():
#         return candidate
#     suffix = 1
#     while True:
#         candidate = parent / f"{name} ({suffix})"
#         if not candidate.exists():
#             return candidate
#         suffix += 1


# def main():
#     changed = 0
#     for p in sorted(TARGET.iterdir()):
#         if not p.is_dir():
#             continue
#         new_name = clean_name(p.name)
#         if new_name == p.name:
#             continue
#         dest = unique_path(TARGET, new_name)
#         print(f"Renaming: {p.name} -> {dest.name}")
#         p.rename(dest)
#         changed += 1
#     print(f"Готово. Переименовано папок: {changed}")


# if __name__ == '__main__':
#     main()
