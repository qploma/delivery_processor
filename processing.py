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

        wrap_format = workbook.add_format({
            "text_wrap": True,
            "valign": "top"
        })

        default_format = workbook.add_format({
            "valign": "top"
        })

        for col_num, column_name in enumerate(df.columns):
            if column_name in ["Товары заявки", "Комментарий", "Адрес доставки"]:
                worksheet.set_column(col_num, col_num, 60, wrap_format)
            else:
                worksheet.set_column(col_num, col_num, 22, default_format)

    return output.getvalue()


def format_quantity(value) -> str:
    if pd.isna(value):
        return ""

    value_str = str(value).replace("\xa0", " ").strip()

    if value_str == "":
        return ""

    value_str = value_str.replace(",", ".")

    try:
        number = float(value_str)

        if number.is_integer():
            return str(int(number))

        return str(number).rstrip("0").rstrip(".")

    except ValueError:
        return value_str


def join_products_with_quantity(group: pd.DataFrame) -> str:
    product_lines = []

    for _, row in group.iterrows():
        product = row.get("Список товаров")
        quantity = row.get("Кол-во товара")

        if pd.isna(product):
            continue

        product_text = str(product).replace("\xa0", " ").strip()

        if product_text == "":
            continue

        quantity_text = format_quantity(quantity)

        if quantity_text != "":
            product_lines.append(f"{product_text} - {quantity_text}шт.")
        else:
            product_lines.append(product_text)

    return "\n".join(product_lines)


def prepare_drivers_file(vod_file) -> pd.DataFrame:
    df_vod = pd.read_excel(vod_file)

    df_vod.columns = df_vod.columns.astype(str).str.strip()

    df_vod = df_vod[["Номер заявки", "ФИО водителя"]].copy()

    df_vod["Номер заявки"] = df_vod["Номер заявки"].astype("string").str.strip()
    df_vod["ФИО водителя"] = df_vod["ФИО водителя"].astype("string").str.strip()

    df_vod = df_vod.dropna(subset=["Номер заявки"])

    df_vod = df_vod.drop_duplicates(
        subset=["Номер заявки"],
        keep="first"
    ).reset_index(drop=True)

    df_vod = df_vod.rename(columns={
        "ФИО водителя": "Водитель"
    })

    return df_vod


def make_short_df(grouped_df: pd.DataFrame) -> pd.DataFrame:
    if "Водитель" in grouped_df.columns:
        short_columns = [
            "Дата доставки",
            "Номер заявки",
            "Водитель",
            "Филиал",
            "Вес заказа",
            "Адрес доставки",
            "Стоимость заказа, руб.",
            "Зона"
        ]
    else:
        short_columns = [
            "Дата доставки",
            "Номер заявки",
            "Филиал",
            "Вес заказа",
            "Адрес доставки",
            "Стоимость заказа, руб.",
            "Зона"
        ]

    return grouped_df[short_columns].copy()


def process_delivery_file(main_file, drivers_file=None, original_filename: str = "file.xlsx"):
    df = pd.read_excel(
        main_file,
        sheet_name="TDSheet"
    )

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

    df = df.dropna(axis=1, how="all")

    df = df.replace(r"^\s*$", pd.NA, regex=True)

    first_col = df.iloc[:, 0].astype(str)

    mask_delete = (
        first_col.str.contains("Филиал:", na=False) |
        first_col.str.contains("Список заявок на доставку", na=False) |
        first_col.str.contains("Доставочная организация:", na=False) |
        first_col.str.contains("Дата:", na=False) |
        first_col.str.contains("№", na=False)
    )

    mask_itogo = df.astype("string").apply(
        lambda col: col.str.contains("ИТОГО", case=False, na=False)
    ).any(axis=1)

    mask_delete = mask_delete | mask_itogo

    cols_except_filial = [col for col in df.columns if col != "Филиал"]
    mask_empty_except_filial = df[cols_except_filial].isna().all(axis=1)

    mask_delete = mask_delete | mask_empty_except_filial

    df = df[~mask_delete].reset_index(drop=True)

    df = df.dropna(axis=1, how="all")

    selected_positions = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 21, 23]

    if df.shape[1] <= max(selected_positions):
        raise ValueError(
            "В файле меньше столбцов, чем ожидалось. "
            "Проверь структуру исходного Excel-файла."
        )

    df = df.iloc[:, selected_positions].copy()

    exclude_col = "Unnamed: 51"

    cols_to_fill = [col for col in df.columns if col != exclude_col]

    df[cols_to_fill] = df[cols_to_fill].replace(r"^\s*$", pd.NA, regex=True)
    df[cols_to_fill] = df[cols_to_fill].ffill()

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

    time_text = (
        df["Время доставки"]
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    df[["Время С", "Время ПО"]] = time_text.str.extract(
        r"[сc]\s*(\d{1,2}[:.]\d{2})\s*(?:по|до)\s*(\d{1,2}[:.]\d{2})",
        flags=re.IGNORECASE,
        expand=True
    )

    df["Время С"] = df["Время С"].str.replace(".", ":", regex=False)
    df["Время ПО"] = df["Время ПО"].str.replace(".", ":", regex=False)

    full_df = df.copy()

    df["Вес заказа"] = (
        df["Вес заказа"]
        .astype(str)
        .str.replace(",", ".", regex=False)
    )

    df["Вес заказа"] = pd.to_numeric(df["Вес заказа"], errors="coerce")

    group_columns = [
        "Дата доставки",
        "Номер заявки",
        "Филиал",
        "Адрес доставки"
    ]

    grouped_base = (
        df.groupby(
            group_columns,
            as_index=False
        )
        .agg({
            "Тип контрагента": "first",
            "Время С": "first",
            "Время ПО": "first",
            "Телефон клиента": "first",
            "Способ оплаты": "first",
            "Стоимость заказа, руб.": "first",
            "Вес заказа": "sum",
            "Комментарий": "first"
        })
        .rename(columns={
            "Тип контрагента": "Тип клиента"
        })
    )

    products_df = (
        df.groupby(group_columns)
        .apply(join_products_with_quantity)
        .reset_index(name="Товары заявки")
    )

    grouped_df = grouped_base.merge(
        products_df,
        on=group_columns,
        how="left"
    )

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

    if drivers_file is not None:
        df_vod = prepare_drivers_file(drivers_file)

        grouped_df["Номер заявки"] = grouped_df["Номер заявки"].astype("string").str.strip()
        df_vod["Номер заявки"] = df_vod["Номер заявки"].astype("string").str.strip()

        grouped_df = grouped_df.merge(
            df_vod,
            on="Номер заявки",
            how="left"
        )

    final_columns = [
        "Дата доставки",
        "Номер заявки",
        "Филиал",
        "Тип клиента",
        "Адрес доставки",
        "Телефон клиента",
        "Время С",
        "Время ПО",
        "Способ оплаты",
        "Стоимость заказа, руб.",
        "Вес заказа",
        "Товары заявки",
        "Комментарий",
        "Зона"
    ]

    if "Водитель" in grouped_df.columns:
        final_columns.append("Водитель")

    grouped_df = grouped_df[final_columns].copy()

    short_df = make_short_df(grouped_df)

    original_stem = Path(original_filename).stem

    full_filename = f"Полный_{original_stem}.xlsx"

    if drivers_file is not None:
        grouped_filename = f"Водители_Зона_СГруппированный_{original_stem}.xlsx"
        short_filename = f"Урезанный_Водители_Зона_{original_stem}.xlsx"
    else:
        grouped_filename = f"Зона_СГруппированный_{original_stem}.xlsx"
        short_filename = f"Урезанный_Зона_{original_stem}.xlsx"

    return full_df, grouped_df, short_df, full_filename, grouped_filename, short_filename
