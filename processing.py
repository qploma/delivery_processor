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
            if column_name in ["Товары заявки", "Список товаров", "Комментарий", "Адрес доставки"]:
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


def clean_phone(value):
    if pd.isna(value):
        return pd.NA

    value_str = str(value).replace("\xa0", " ").strip()

    if value_str == "":
        return pd.NA

    value_str = value_str.replace(" ", "")

    try:
        number = float(value_str)

        if number.is_integer():
            return str(int(number))

    except ValueError:
        pass

    if value_str.endswith(".0"):
        value_str = value_str[:-2]

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

    # Чистим названия столбцов: убираем переносы, неразрывные пробелы и лишние пробелы
    df_vod.columns = (
        df_vod.columns
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    required_columns = ["ФИО водителя", "Номер заявки"]

    missing_columns = [
        col for col in required_columns
        if col not in df_vod.columns
    ]

    if missing_columns:
        found_columns = ", ".join(map(str, df_vod.columns))

        raise ValueError(
            "Файл с водителями должен содержать ровно два столбца: "
            "«ФИО водителя» и «Номер заявки». "
            f"Не найдены столбцы: {', '.join(missing_columns)}. "
            f"Найденные столбцы в файле: {found_columns}"
        )

    # Берём только нужные два столбца.
    # Порядок в Excel не важен: может быть сначала ФИО, потом номер заявки.
    df_vod = df_vod[["Номер заявки", "ФИО водителя"]].copy()

    df_vod["Номер заявки"] = (
        df_vod["Номер заявки"]
        .astype("string")
        .str.replace("\xa0", " ", regex=False)
        .str.replace(r"\s+", "", regex=True)
        .str.strip()
    )

    df_vod["ФИО водителя"] = (
        df_vod["ФИО водителя"]
        .astype("string")
        .str.replace("\xa0", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    df_vod = df_vod.dropna(subset=["Номер заявки"])
    df_vod = df_vod[df_vod["Номер заявки"] != ""]

    df_vod = df_vod.drop_duplicates(
        subset=["Номер заявки"],
        keep="first"
    ).reset_index(drop=True)

    df_vod = df_vod.rename(columns={
        "ФИО водителя": "Водитель"
    })

    return df_vod


def make_products_df(df: pd.DataFrame, group_columns: list[str]) -> pd.DataFrame:
    rows = []

    for group_key, group in df.groupby(group_columns, dropna=False, sort=False):
        if not isinstance(group_key, tuple):
            group_key = (group_key,)

        row = dict(zip(group_columns, group_key))
        row["Товары заявки"] = join_products_with_quantity(group)
        rows.append(row)

    return pd.DataFrame(rows)


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

    # --------------------------------------------------
    # 1. Добавляем филиал
    # --------------------------------------------------

    current_filial = None
    filial_values = []

    for i in range(len(df)):
        first_col_value = df.iloc[i, 0]

        if df.shape[1] > 4:
            fifth_col_value = df.iloc[i, 4]
        else:
            fifth_col_value = None

        if pd.notna(first_col_value) and "Филиал" in str(first_col_value):
            if pd.notna(fifth_col_value) and str(fifth_col_value).strip() != "":
                current_filial = str(fifth_col_value).strip()

        filial_values.append(current_filial)

    df["Филиал"] = filial_values

    # --------------------------------------------------
    # 2. ВАЖНО: не удаляем пустые столбцы
    # --------------------------------------------------
    # Раньше здесь был dropna(axis=1, how="all").
    # Теперь его нет, чтобы нужные, но пустые столбцы не пропадали.

    df = df.replace(r"^\s*$", pd.NA, regex=True)

    # --------------------------------------------------
    # 3. Удаляем лишние строки
    # --------------------------------------------------

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

    # --------------------------------------------------
    # 4. Забираем нужные исходные столбцы
    # --------------------------------------------------
    # Если нужный столбец оказался полностью пустым и pandas его не прочитал,
    # создаём его пустым, чтобы на выходе структура всегда была одинаковой.

    source_columns = [
        "Unnamed: 3",
        "Unnamed: 5",
        "Unnamed: 8",
        "Unnamed: 11",
        "Unnamed: 13",
        "Unnamed: 20",
        "Unnamed: 21",
        "Unnamed: 23",
        "Unnamed: 25",
        "Unnamed: 27",
        "Unnamed: 29",
        "Unnamed: 35",
        "Unnamed: 51",
        "Филиал"
    ]

    for col in source_columns:
        if col not in df.columns:
            df[col] = pd.NA

    df = df[source_columns].copy()

    # --------------------------------------------------
    # 5. Заполняем пропуски вниз, кроме комментария
    # --------------------------------------------------

    exclude_col = "Unnamed: 51"

    cols_to_fill = [col for col in df.columns if col != exclude_col]

    df[cols_to_fill] = df[cols_to_fill].replace(r"^\s*$", pd.NA, regex=True)
    df[cols_to_fill] = df[cols_to_fill].ffill()

    # --------------------------------------------------
    # 6. Переименовываем столбцы
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

    df = df[required_columns].copy()

    df["Номер заявки"] = df["Номер заявки"].astype("string").str.strip()
    df["Телефон клиента"] = df["Телефон клиента"].apply(clean_phone)

    # --------------------------------------------------
    # 7. Разбиваем время доставки
    # --------------------------------------------------

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

    # --------------------------------------------------
    # 8. Полный обработанный файл
    # --------------------------------------------------

    full_df = df.copy()

    # --------------------------------------------------
    # 9. Подготавливаем вес
    # --------------------------------------------------

    df["Вес заказа"] = (
        df["Вес заказа"]
        .astype(str)
        .str.replace(",", ".", regex=False)
    )

    df["Вес заказа"] = pd.to_numeric(df["Вес заказа"], errors="coerce")

    # Дополнительная гарантия является услугой, а не физическим товаром.
    # Строка остаётся в полном файле и в перечне товаров, но её вес
    # не участвует в агрегировании сгруппированного и урезанного файлов.
    additional_warranty_mask = (
        df["Список товаров"]
        .astype("string")
        .str.replace("\xa0", " ", regex=False)
        .str.contains(r"доп\.?\s*гарантия", case=False, na=False, regex=True)
    )

    df.loc[additional_warranty_mask, "Вес заказа"] = 0

    # --------------------------------------------------
    # 10. Сгруппированный файл
    # --------------------------------------------------

    group_columns = [
        "Дата доставки",
        "Номер заявки",
        "Филиал",
        "Адрес доставки"
    ]

    grouped_base = (
        df.groupby(
            group_columns,
            as_index=False,
            dropna=False,
            sort=False
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

    products_df = make_products_df(df, group_columns)

    grouped_df = grouped_base.merge(
        products_df,
        on=group_columns,
        how="left"
    )

    # --------------------------------------------------
    # 11. Добавляем зону
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
    # 12. Добавляем водителей, если файл загружен
    # --------------------------------------------------

    if drivers_file is not None:
        df_vod = prepare_drivers_file(drivers_file)

        grouped_df["Номер заявки"] = grouped_df["Номер заявки"].astype("string").str.strip()
        df_vod["Номер заявки"] = df_vod["Номер заявки"].astype("string").str.strip()

        grouped_df = grouped_df.merge(
            df_vod,
            on="Номер заявки",
            how="left"
        )

    # --------------------------------------------------
    # 13. Финальный порядок столбцов сгруппированного файла
    # --------------------------------------------------

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

    for col in final_columns:
        if col not in grouped_df.columns:
            grouped_df[col] = pd.NA

    grouped_df = grouped_df[final_columns].copy()

    # --------------------------------------------------
    # 14. Урезанный файл
    # --------------------------------------------------

    short_df = make_short_df(grouped_df)

    # --------------------------------------------------
    # 15. Названия файлов
    # --------------------------------------------------

    original_stem = Path(original_filename).stem

    full_filename = f"Полный_{original_stem}.xlsx"

    if drivers_file is not None:
        grouped_filename = f"Водители_Зона_СГруппированный_{original_stem}.xlsx"
        short_filename = f"Урезанный_Водители_Зона_{original_stem}.xlsx"
    else:
        grouped_filename = f"Зона_СГруппированный_{original_stem}.xlsx"
        short_filename = f"Урезанный_Зона_{original_stem}.xlsx"

    return full_df, grouped_df, short_df, full_filename, grouped_filename, short_filename


def normalize_compact_text(value) -> str:
    if pd.isna(value):
        return ""

    value_text = str(value).replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", value_text)


def first_non_empty(series: pd.Series):
    for value in series:
        if pd.isna(value):
            continue

        if isinstance(value, str):
            cleaned = normalize_compact_text(value)

            if cleaned != "":
                return cleaned
        else:
            return value

    return pd.NA


def split_phone_values(value) -> list[str]:
    if pd.isna(value):
        return []

    value_text = str(value).replace("\xa0", " ").strip()

    if value_text == "":
        return []

    parts = re.split(r"[;\n\r]+", value_text)
    phones = []

    for part in parts:
        cleaned_phone = clean_phone(part)

        if pd.isna(cleaned_phone):
            continue

        cleaned_phone = str(cleaned_phone).strip()

        if cleaned_phone != "":
            phones.append(cleaned_phone)

    return phones


def join_unique_phones(series: pd.Series) -> str:
    phones = []
    seen = set()

    for value in series:
        for phone in split_phone_values(value):
            if phone not in seen:
                seen.add(phone)
                phones.append(phone)

    return "; ".join(phones)


def join_products_for_bitrix24(group: pd.DataFrame) -> str:
    product_lines = []

    for _, row in group.iterrows():
        product = row.get("Список товаров")
        quantity = row.get("Кол-во товара")

        if pd.isna(product):
            continue

        product_text = normalize_compact_text(product)

        if product_text == "":
            continue

        quantity_text = format_quantity(quantity)

        if quantity_text != "":
            product_lines.append(f"{product_text} - {quantity_text}шт.")
        else:
            product_lines.append(product_text)

    return "; ".join(product_lines)


def make_bitrix24_export_df(prepared_df: pd.DataFrame) -> pd.DataFrame:
    df = prepared_df.copy()

    df["Номер заявки"] = (
        df["Номер заявки"]
        .astype("string")
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    df = df[
        df["Номер заявки"].notna()
        & (df["Номер заявки"] != "")
    ].copy()

    df["Телефон клиента"] = df["Телефон клиента"].apply(clean_phone)

    df["Вес заказа"] = (
        df["Вес заказа"]
        .astype("string")
        .str.replace("\xa0", " ", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    df["Вес заказа"] = pd.to_numeric(df["Вес заказа"], errors="coerce")

    normalized_products = (
        df["Список товаров"]
        .astype("string")
        .str.replace("\xa0", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .str.lower()
    )

    delivery_service_mask = normalized_products.eq("доставка товара клиенту")
    df.loc[delivery_service_mask, "Вес заказа"] = 0

    rows = []

    for order_number, group in df.groupby(
        "Номер заявки",
        sort=False,
        dropna=False
    ):
        order_weight = group["Вес заказа"].sum(min_count=1)

        rows.append({
            "Название": order_number,
            "Номер заявки": order_number,
            "Дата доставки": first_non_empty(group["Дата доставки"]),
            "Время доставки": first_non_empty(group["Время доставки"]),
            "Список товаров": join_products_for_bitrix24(group),
            "Вес заказа, кг": order_weight,
            "Адрес доставки": first_non_empty(group["Адрес доставки"]),
            "Адрес на карте": "",
            "Телефоны клиента": join_unique_phones(group["Телефон клиента"]),
            "Способ оплаты": first_non_empty(group["Способ оплаты"]),
            "Стоимость заказа": first_non_empty(group["Стоимость заказа, руб."]),
            "Комментарий": first_non_empty(group["Комментарий"]),
            "Водитель": "",
            "Номер маршрута": "",
            "Логист": "",
        })

    result_columns = [
        "Название",
        "Номер заявки",
        "Дата доставки",
        "Время доставки",
        "Список товаров",
        "Вес заказа, кг",
        "Адрес доставки",
        "Адрес на карте",
        "Телефоны клиента",
        "Способ оплаты",
        "Стоимость заказа",
        "Комментарий",
        "Водитель",
        "Номер маршрута",
        "Логист"
    ]

    return pd.DataFrame(rows, columns=result_columns)


def process_bitrix24_file(
    main_file,
    original_filename: str = "file.xlsx"
) -> tuple[pd.DataFrame, str]:
    df = pd.read_excel(
        main_file,
        sheet_name="TDSheet"
    )

    current_filial = None
    filial_values = []

    for row_index in range(len(df)):
        first_col_value = df.iloc[row_index, 0]

        if df.shape[1] > 4:
            fifth_col_value = df.iloc[row_index, 4]
        else:
            fifth_col_value = None

        if pd.notna(first_col_value) and "Филиал" in str(first_col_value):
            if pd.notna(fifth_col_value) and str(fifth_col_value).strip() != "":
                current_filial = str(fifth_col_value).strip()

        filial_values.append(current_filial)

    df["Филиал"] = filial_values
    df = df.replace(r"^\s*$", pd.NA, regex=True)

    first_col = df.iloc[:, 0].astype(str)

    mask_delete = (
        first_col.str.contains("Филиал:", na=False)
        | first_col.str.contains("Список заявок на доставку", na=False)
        | first_col.str.contains("Доставочная организация:", na=False)
        | first_col.str.contains("Дата:", na=False)
        | first_col.str.contains("№", na=False)
    )

    mask_itogo = df.astype("string").apply(
        lambda column: column.str.contains("ИТОГО", case=False, na=False)
    ).any(axis=1)

    cols_except_filial = [
        column
        for column in df.columns
        if column != "Филиал"
    ]
    mask_empty_except_filial = df[cols_except_filial].isna().all(axis=1)

    df = df[
        ~(mask_delete | mask_itogo | mask_empty_except_filial)
    ].reset_index(drop=True)

    source_columns = [
        "Unnamed: 3",
        "Unnamed: 8",
        "Unnamed: 11",
        "Unnamed: 13",
        "Unnamed: 20",
        "Unnamed: 23",
        "Unnamed: 25",
        "Unnamed: 27",
        "Unnamed: 29",
        "Unnamed: 35",
        "Unnamed: 51"
    ]

    for column in source_columns:
        if column not in df.columns:
            df[column] = pd.NA

    df = df[source_columns].copy()

    comment_column = "Unnamed: 51"
    columns_to_fill = [
        column
        for column in df.columns
        if column != comment_column
    ]

    df[columns_to_fill] = df[columns_to_fill].replace(
        r"^\s*$",
        pd.NA,
        regex=True
    )
    df[columns_to_fill] = df[columns_to_fill].ffill()

    df = df.rename(columns={
        "Unnamed: 3": "Номер заявки",
        "Unnamed: 8": "Дата доставки",
        "Unnamed: 11": "Время доставки",
        "Unnamed: 13": "Список товаров",
        "Unnamed: 20": "Кол-во товара",
        "Unnamed: 23": "Вес заказа",
        "Unnamed: 25": "Адрес доставки",
        "Unnamed: 27": "Телефон клиента",
        "Unnamed: 29": "Способ оплаты",
        "Unnamed: 35": "Стоимость заказа, руб.",
        "Unnamed: 51": "Комментарий"
    })

    bitrix_df = make_bitrix24_export_df(df)

    original_stem = Path(original_filename).stem
    bitrix_filename = f"{original_stem}_Битрикс24.csv"

    return bitrix_df, bitrix_filename




def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    export_df = df.copy()

    for column in export_df.columns:
        if export_df[column].dtype == "object":
            export_df[column] = export_df[column].apply(
                lambda value: value.replace(";", "&")
                if isinstance(value, str)
                else value
            )

    return export_df.to_csv(
        index=False,
        sep=";",
        encoding="utf-8-sig"
    ).encode("utf-8-sig")
