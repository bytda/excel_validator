import pandas as pd
import sys
import os

def validate_excel(input_file: str, master_file: str, output_file: str = "validated_output.xlsx"):
    """
    Проверяет входной Excel-файл по правилам:
      1. Удаляет строки, где D, N, O, Q — все пустые.
      2. Удаляет строки, где S пустой или равен "0".
      3. Проверяет, что артикул (столбец Q) присутствует в мастер-данных.
      4. Оставляет только столбцы D, N, O, Q, S.
    """

    # Проверяем существование файлов
    if not os.path.exists(input_file):
        print(f"Ошибка: файл '{input_file}' не найден.")
        sys.exit(1)

    if not os.path.exists(master_file):
        print(f"Ошибка: файл '{master_file}' не найден.")
        sys.exit(1)

    # Загружаем данные
    print("Загрузка данных...")
    df = pd.read_excel(input_file)
    master = pd.read_excel(master_file)

    # Определяем столбец с артикулом в мастер-данных
    if "Артикул" in master.columns:
        master_art_col = "Артикул"
    else:
        master_art_col = master.columns[0]

    # Извлекаем нужные столбцы по индексам Excel (D=3, N=13, O=14, Q=16, S=18)
    cols_idx = {"D": 3, "N": 13, "O": 14, "Q": 16, "S": 18}
    df_subset = df.iloc[:, list(cols_idx.values())].copy()
    df_subset.columns = list(cols_idx.keys())

    # 1) Удаляем строки, где D, N, O, Q все null
    df_subset = df_subset[~(df_subset[["D", "N", "O", "Q"]].isnull().all(axis=1))]

    # 2) Удаляем строки, где S пустой или равен "0"
    df_subset = df_subset[~(df_subset["S"].isnull() | (df_subset["S"].astype(str).str.strip() == "0"))]

    # 3) Проверяем артикулы (по Q)
    master_articles = master[master_art_col].astype(str).str.strip().unique()
    df_subset = df_subset[df_subset["Q"].astype(str).str.strip().isin(master_articles)]

    # 4) Сохраняем результат
    df_subset.to_excel(output_file, index=False)
    print(f"Валидация завершена. Результат сохранён в: {output_file}")


if __name__ == "__main__":
    # Использование из командной строки
    # Пример: python excel_validator.py input.xlsx master_data.xlsx output.xlsx
    if len(sys.argv) < 3:
        print("Использование: python excel_validator.py <input.xlsx> <master_data.xlsx> [output.xlsx]")
        sys.exit(0)

    input_path = sys.argv[1]
    master_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else "validated_output.xlsx"

    validate_excel(input_path, master_path, output_path)
