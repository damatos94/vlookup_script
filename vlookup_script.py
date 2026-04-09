import pandas as pd
import os
import sys

# ================= ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =================

def col_letter_to_index(letter: str) -> int:
    """Преобразует букву столбца (A, B, C, ..., AA, AB...) в индекс (0-based)."""
    letter = letter.upper().strip()
    index = 0
    for ch in letter:
        index = index * 26 + (ord(ch) - ord('A') + 1)
    return index - 1

def get_column_choice(prompt):
    """Запрашивает букву столбца, возвращает индекс (0-based)."""
    while True:
        inp = input(prompt).strip()
        if not inp:
            print("❌ Необходимо ввести букву столбца (A, B, C, AA, AB и т.д.)")
            continue
        try:
            idx = col_letter_to_index(inp)
            return idx
        except Exception:
            print(f"❌ Некорректная буква столбца: {inp}. Используйте A, B, C, AA, AB...")

def repair_excel_file(filepath):
    """Автовосстановление повреждённого xlsx через скрытый Excel (Windows)"""
    try:
        import win32com.client
    except ImportError:
        print("❌ Отсутствует pywin32. Установите: pip install pywin32")
        return False

    print(f"\n🔧 Восстановление структуры: {os.path.basename(filepath)}")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False

    try:
        wb = excel.Workbooks.Open(os.path.abspath(filepath))
        wb.SaveAs(os.path.abspath(filepath), FileFormat=51)
        wb.Close(False)
        print("✅ Файл успешно пересохранён в корректный формат.")
        return True
    except Exception as e:
        print(f"❌ Ошибка восстановления: {e}")
        return False
    finally:
        excel.Quit()

def read_excel_safe(path, sheet_name, **kwargs):
    """Читает Excel. При битом архиве -> чинит -> читает снова"""
    try:
        return pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl', **kwargs)
    except Exception as e:
        err_str = str(e).lower()
        if 'sharedstrings.xml' in err_str or 'bad zip' in err_str:
            if repair_excel_file(path):
                return pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl', **kwargs)
        raise e

def get_sheet_name(filepath, prompt):
    """Запрашивает имя листа. Фолбэк: Лист1 -> TDSheet. Если листа нет, показывает список."""
    try:
        xl = pd.ExcelFile(filepath, engine='openpyxl')
        sheets = xl.sheet_names
    except Exception as e:
        print(f"❌ Не удалось прочитать список листов: {e}")
        sys.exit(1)

    while True:
        inp = input(prompt).strip()
        if not inp:
            # 🔽 Внутренний фолбэк без изменения внешнего интерфейса
            if "Лист1" in sheets:
                inp = "Лист1"
            elif "TDSheet" in sheets:
                inp = "TDSheet"
            else:
                inp = "Лист1"  # Намеренно, чтобы сработало стандартное сообщение об ошибке ниже
        if inp in sheets:
            return inp
        else:
            print(f"❌ Лист '{inp}' не найден. Доступные листы: {', '.join(sheets)}")

def get_file_path(prompt):
    """Запрашивает путь к файлу, добавляет .xlsx при необходимости, проверяет существование."""
    while True:
        raw = input(prompt).strip()
        if not raw:
            print("❌ Имя файла не может быть пустым.")
            continue
        if (raw.startswith('"') and raw.endswith('"')) or (raw.startswith("'") and raw.endswith("'")):
            raw = raw[1:-1]
        if not raw.lower().endswith('.xlsx'):
            raw += '.xlsx'
        if os.path.exists(raw):
            return raw
        else:
            print(f"❌ Файл '{raw}' не найден.\n")

# ================= ОСНОВНАЯ ФУНКЦИЯ =================

def main():
    print("=" * 70)
    print("📊 АВТОМАТИЧЕСКАЯ ПОДСТАНОВКА ДАННЫХ (аналог ВПР/XLOOKUP)")
    print("=" * 70)
    print("\n📋 КАК ЭТО РАБОТАЕТ:")
    print("   Скрипт найдёт совпадения между двумя таблицами и перенесёт")
    print("   нужные значения из файла-справочника в основной файл.")
    print("\n🔢 ВАШИ ДЕЙСТВИЯ:")
    print("   1. Укажите основной файл (куда вставлять данные).")
    print("   2. Укажите файл-справочник (откуда брать данные).")
    print("   3. Выберите листы в каждом файле (по умолчанию: Лист1; TDSheet).")
    print("   4. В основном файле укажите:")
    print("      • Букву столбца с ключами для поиска.")
    print("      • Букву столбца, куда записать результат.")
    print("   5. В справочнике укажите:")
    print("      • Букву столбца, где искать ключи.")
    print("      • Букву столбца, откуда брать значения.")
    print("\n💡 ВАЖНО ЗНАТЬ:")
    print("   • Столбцы указываются латинскими буквами: A, B, C, … AA, AB…")
    print("   • Если совпадение не найдено, в ячейке появится «НД».")
    print("   • Результат сохранится как «<имя_файла>_результат.xlsx».")
    print("   • Перед запуском обязательно закройте оба файла в Excel.")
    print("=" * 70)

    work_dir = os.path.dirname(os.path.abspath(__file__))
    if work_dir:
        os.chdir(work_dir)
        print(f"📁 Рабочая папка: {work_dir}\n")
    else:
        print(f"📁 Рабочая папка: {os.getcwd()}\n")

    # Запрашиваем файлы
    main_file = get_file_path("📂 Введите имя основного файла: ")
    lookup_file = get_file_path("📂 Введите имя файла-справочника: ")

    # Запрашиваем листы
    main_sheet = get_sheet_name(main_file, f"   Введите имя листа в основном файле (по умолчанию: Лист1; TDSheet): ")
    lookup_sheet = get_sheet_name(lookup_file, f"   Введите имя листа в справочнике (по умолчанию: Лист1; TDSheet): ")

    # Запрашиваем столбцы
    print("\n--- Настройка столбцов в основном файле ---")
    search_col_idx = get_column_choice("   Из какого столбца брать данные для поиска (буква): ")
    target_col_idx = get_column_choice("   В какой столбец записать результат (буква): ")

    print("\n--- Настройка столбцов в справочнике ---")
    lookup_key_idx = get_column_choice("   В каком столбце искать совпадение (ключ) (буква): ")
    lookup_val_idx = get_column_choice("   Из какого столбца брать значение (буква): ")

    base_name = os.path.splitext(main_file)[0]
    output_file = f"{base_name}_результат.xlsx"

    print(f"\n✅ Основной файл: {main_file} (лист '{main_sheet}')")
    print(f"✅ Справочник: {lookup_file} (лист '{lookup_sheet}')")
    print(f"✅ Результат: {output_file}")
    print(f"✅ Поиск в основном файле по столбцу с индексом {search_col_idx}")
    print(f"✅ Запись результата в столбец с индексом {target_col_idx}")
    print(f"✅ Справочник: ключ в столбце с индексом {lookup_key_idx}")
    print(f"✅ Справочник: значение в столбце с индексом {lookup_val_idx}")
    print()

    try:
        # Читаем основной файл (нужны все столбцы для сохранения исходной структуры)
        print("📂 Чтение основного файла...")
        df_main = read_excel_safe(main_file, main_sheet)

        max_col_main = len(df_main.columns) - 1
        if search_col_idx > max_col_main:
            raise ValueError(f"В основном файле нет столбца с индексом {search_col_idx} (всего столбцов: {max_col_main+1})")
        if target_col_idx > max_col_main:
            print(f"⚠️ Столбец с индексом {target_col_idx} будет создан (сейчас в файле {max_col_main+1} столбцов).")

        search_col_name = df_main.columns[search_col_idx]
        print(f"🔍 Поиск по столбцу: '{search_col_name}' (индекс {search_col_idx})")

        # 🔥 ОПТИМИЗАЦИЯ: Читаем справочник ТОЛЬКО по нужным индексам
        print("📂 Чтение справочника...")
        df_lookup = read_excel_safe(lookup_file, lookup_sheet, header=None, usecols=[lookup_key_idx, lookup_val_idx])

        # pandas сортирует столбцы по возрастанию индекса при частичной загрузке.
        # Назначаем имена корректно, независимо от порядка ввода.
        if lookup_key_idx < lookup_val_idx:
            df_lookup.columns = ['key', 'value']
        else:
            df_lookup.columns = ['value', 'key']

        # Векторизованная очистка: убираем пустые ключи, приводим типы, удаляем дубли
        df_lookup = df_lookup[df_lookup['key'].notna()].copy()
        df_lookup['key'] = df_lookup['key'].astype(str)
        df_lookup['value'] = df_lookup['value'].astype(str)
        df_lookup = df_lookup.drop_duplicates(subset='key')

        print(f"🔧 Справочник загружен: {len(df_lookup)} уникальных ключей")

        # Подстановка
        print("\n🔄 Выполняется подстановка...")
        lookup_map = df_lookup.set_index('key')['value'].to_dict()

        # Гарантируем наличие целевого столбца
        if target_col_idx >= len(df_main.columns):
            for _ in range(target_col_idx - len(df_main.columns) + 1):
                df_main[f"Unnamed_{len(df_main.columns)}"] = ""

        target_col_name = df_main.columns[target_col_idx]

        # Быстрая векторизованная подстановка через .map()
        df_main[target_col_name] = (
            df_main[search_col_name]
            .astype(str)
            .map(lookup_map)
            .fillna("НД")
        )

        # Сохранение
        print(f"💾 Сохранение результата: {output_file}")
        df_main.to_excel(output_file, index=False, engine='openpyxl')

        print("\n" + "=" * 70)
        print("✅ ГОТОВО!")
        print("=" * 70)
        print(f"📌 ИТОГ:")
        print(f"   - Искал в столбце с индексом {search_col_idx} основного файла (лист '{main_sheet}')")
        print(f"   - Сравнивал со столбцом с индексом {lookup_key_idx} справочника (лист '{lookup_sheet}')")
        print(f"   - Подставлял из столбца с индексом {lookup_val_idx} справочника")
        print(f"   - Результат записан в столбец с индексом {target_col_idx} основного файла")
        print(f"   - Пустые значения заменены на 'НД'")
        print(f"   - Файл сохранён: {output_file}")

    except PermissionError:
        print("\n❌ Ошибка: закройте оба файла в Excel и запустите снова.")
    except Exception as e:
        print(f"\n❌ Критическая ошибка:\n{e}")
    finally:
        input("\nНажмите Enter, чтобы закрыть окно...")

if __name__ == "__main__":
    main()