from io import BytesIO
from pathlib import Path
import re

import pandas as pd


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Результат")

        workbook = writer.book
        worksheet = writer.sheets["Результат"]

        # Формат с переносом текста внутри ячейки
        wrap_format = workbook.add_format({
            "text_wrap": True,
            "valign": "top"
        })

        # Применяем перенос текста ко всем столбцам
        for col_num, column_name in enumerate(df.columns):
            if column_name == "Товары заявки":
                worksheet.set_column(col_num, col_num, 60, wrap_format)
            else:
                worksheet.set_column(col_num, col_num, 20)

    return output.getvalue()


def join_products(series: pd.Series) -> str:
    products = (
        series
        .dropna()
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    products = products[products != ""]

    # Убираем повторы, но сохраняем порядок
    products = products.drop_duplicates()

    # Каждый товар с новой строки внутри одной Excel-ячейки
    return "\n".join(products)

def prepare_drivers_file(vod_file) -> pd.DataFrame:
    df_vod = pd.read_excel(vod_file)

    df_vod.columns = df_vod.columns.astype(str).str.strip()

    df_vod = df_vod[["Номер заявки", "ФИО водителя"]].copy()

    df_vod["Номер заявки"] = df_vod["Номер заявки"].astype("string").str.strip()
    df_vod["ФИО водителя"] = df_vod["ФИО водителя"].astype("string").str.strip()

    df_vod = df_vod.dropna(subset=["Номер заявки"])

    # Оставляем только первое вхождение заявки
    df_vod = df_vod.drop_duplicates(
        subset=["Номер заявки"],
        keep="first"
    ).reset_index(drop=True)

    return df_vod


def process_delivery_file(main_file, drivers_file, original_filename: str):
    df = pd.read_excel(
        main_file,
        sheet_name="TDSheet"
    )

    # --------------------------------------------------
    # 1. Добавляем филиал
    # --------------------------------------------------

    current_filial = None
    filial_values = []

    for i in range(len(df)):
        first_col_value = df.iloc[i, 0]
        fifth_col_value = df.iloc[i, 4]

        if pd.notna(first_col_value) and "Филиал" in str(first_col_value):
            if pd.notna(fifth_col_value) and str(fifth_col_value).strip() != "":
                current_filial = str(fifth_col_value).strip()

        filial_values.append(current_filial)

    df["Филиал"] = filial_values

    # --------------------------------------------------
    # 2. Удаляем полностью пустые столбцы
    # --------------------------------------------------

    df = df.dropna(axis=1, how="all")

    # --------------------------------------------------
    # 3. Удаляем лишние строки
    # --------------------------------------------------

    df = df.replace(r"^\s*$", pd.NA, regex=True)

    first_col = df.iloc[:, 0].astype(str)

    mask_delete = (
        first_col.str.contains("Филиал:", na=False) |
        first_col.str.contains("Список заявок на доставку", na=False) |
        first_col.str.contains("Доставочная организация:", na=False) |
        first_col.str.contains("Дата:", na=False) |
        first_col.str.contains("№", na=False)
    )

    # Безопасно ищем ИТОГО по всей строке
    mask_itogo = df.astype("string").apply(
        lambda col: col.str.contains("ИТОГО", case=False, na=False)
    ).any(axis=1)

    mask_delete = mask_delete | mask_itogo

    # Удаляем строки, которые пустые во всех столбцах, кроме "Филиал"
    cols_except_filial = [col for col in df.columns if col != "Филиал"]
    mask_empty_except_filial = df[cols_except_filial].isna().all(axis=1)

    mask_delete = mask_delete | mask_empty_except_filial

    df = df[~mask_delete].reset_index(drop=True)

    # --------------------------------------------------
    # 4. Ещё раз удаляем пустые столбцы после чистки строк
    # --------------------------------------------------

    df = df.dropna(axis=1, how="all")

    # --------------------------------------------------
    # 5. Оставляем нужные столбцы по позициям
    # --------------------------------------------------

    selected_positions = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 21, 23]

    if df.shape[1] <= max(selected_positions):
        raise ValueError(
            "В файле меньше столбцов, чем ожидалось. "
            "Проверь структуру исходного Excel-файла."
        )

    df = df.iloc[:, selected_positions].copy()

    # --------------------------------------------------
    # 6. Заполняем пропуски вниз, кроме комментария
    # --------------------------------------------------

    exclude_col = "Unnamed: 51"

    cols_to_fill = [col for col in df.columns if col != exclude_col]

    df[cols_to_fill] = df[cols_to_fill].replace(r"^\s*$", pd.NA, regex=True)
    df[cols_to_fill] = df[cols_to_fill].ffill()

    # --------------------------------------------------
    # 7. Переименовываем столбцы
    # --------------------------------------------------

    df = df.rename(columns={
        "Unnamed: 3": "Номер заявки",
        "Unnamed: 5": "Тип контрагента",
        "Unnamed: 8": "Дата доставки",
        "Unnamed: 11": "Время доставки",
        "Unnamed: 13": "Список товаров",
        "Unnamed: 20": "Кол-во товара",
        "Unnamed: 21": "Объем заказа",
        "Unnamed: 23": "Вес заказа",
        "Unnamed: 25": "Адрес доставки",
        "Unnamed: 27": "Телефон клиента",
        "Unnamed: 29": "Способ оплаты",
        "Unnamed: 35": "Стоимость заказа, руб.",
        "Unnamed: 51": "Комментарий"
    })

    required_columns = [
        "Номер заявки",
        "Тип контрагента",
        "Дата доставки",
        "Время доставки",
        "Список товаров",
        "Кол-во товара",
        "Объем заказа",
        "Вес заказа",
        "Адрес доставки",
        "Телефон клиента",
        "Способ оплаты",
        "Стоимость заказа, руб.",
        "Комментарий",
        "Филиал"
    ]

    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        raise ValueError(
            "Не найдены нужные столбцы после обработки: "
            + ", ".join(missing_columns)
        )

    df = df[required_columns].copy()

    # --------------------------------------------------
    # 8. Разбиваем время доставки
    # --------------------------------------------------

    time_text = (
        df["Время доставки"]
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    df[["Время С", "Время ДО"]] = time_text.str.extract(
        r"[сc]\s*(\d{1,2}[:.]\d{2})\s*(?:по|до)\s*(\d{1,2}[:.]\d{2})",
        flags=re.IGNORECASE,
        expand=True
    )

    df["Время С"] = df["Время С"].str.replace(".", ":", regex=False)
    df["Время ДО"] = df["Время ДО"].str.replace(".", ":", regex=False)

    # --------------------------------------------------
    # 9. Полный файл
    # --------------------------------------------------

    full_df = df.copy()

    # --------------------------------------------------
    # 10. Подготавливаем вес для группировки
    # --------------------------------------------------

    df["Вес заказа"] = (
        df["Вес заказа"]
        .astype(str)
        .str.replace(",", ".", regex=False)
    )

    df["Вес заказа"] = pd.to_numeric(df["Вес заказа"], errors="coerce")

    # --------------------------------------------------
    # 11. Сгруппированный файл
    # --------------------------------------------------


    grouped_df = (
        df.groupby(
            [
                "Дата доставки",
                "Номер заявки",
                "Филиал",
                "Адрес доставки"
            ],
            as_index=False
        )
        .agg({
            "Стоимость заказа, руб.": "first",
            "Вес заказа": "sum",
            "Список товаров": join_products,
            "Комментарий": "first"
        })
        .rename(columns={
            "Список товаров": "Товары заявки"
        })
    )

    # --------------------------------------------------
    # 12. Добавляем зону
    # --------------------------------------------------

    grouped_df["Зона"] = pd.NA

    address_text = (
        grouped_df["Адрес доставки"]
        .astype("string")
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    mask_zone_0 = address_text.str.contains(
        r"Санкт-Петербург,|Псков,|Великий Новгород,",
        case=False,
        na=False,
        regex=True
    )

    grouped_df.loc[mask_zone_0, "Зона"] = 0

    # --------------------------------------------------
    # 13. Добавляем водителей
    # --------------------------------------------------

    df_vod = prepare_drivers_file(drivers_file)

    grouped_df["Номер заявки"] = grouped_df["Номер заявки"].astype("string").str.strip()
    df_vod["Номер заявки"] = df_vod["Номер заявки"].astype("string").str.strip()

    grouped_df = grouped_df.merge(
        df_vod,
        on="Номер заявки",
        how="left"
    )

    # --------------------------------------------------
    # 14. Названия файлов
    # --------------------------------------------------

    original_stem = Path(original_filename).stem

    full_filename = f"Полный_{original_stem}.xlsx"
    grouped_filename = f"Водители_Зона_СГруппированный_{original_stem}.xlsx"

    return full_df, grouped_df, full_filename, grouped_filename