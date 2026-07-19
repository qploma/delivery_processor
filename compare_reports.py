from io import BytesIO
import re

import pandas as pd


ORDER_NUMBER_ALIASES = {
    "номер заказа",
    "номер заявки",
    "заявка",
}

ORDER_COST_ALIASES = {
    "стоимость заказа",
}

DELIVERY_COST_ALIASES = {
    "стоимость доставки",
}


DUPLICATE_COLUMNS = [
    "Номер заказа",
    "Стоимость заказа",
    "Стоимость доставки",
    "Строка в исходном файле",
    "Количество строк с этим номером",
    "Тип дубля",
]


def normalize_header(value) -> str:
    text = str(value).replace("\xa0", " ").replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def normalize_order_number(value):
    if pd.isna(value):
        return pd.NA

    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", "", text)

    if text == "":
        return pd.NA

    if text.endswith(".0"):
        text = text[:-2]

    return text


def parse_money(value):
    if pd.isna(value):
        return pd.NA

    text = str(value).replace("\xa0", " ").strip()

    if text == "":
        return pd.NA

    text = re.sub(r"\s+", "", text)
    text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return pd.NA


def find_column(columns, aliases: set[str]):
    normalized_columns = {
        normalize_header(column): column
        for column in columns
    }

    for alias in aliases:
        if alias in normalized_columns:
            return normalized_columns[alias]

    return None


def series_has_different_values(series: pd.Series) -> bool:
    normalized_values = []

    for value in series:
        if pd.isna(value):
            normalized_values.append(("empty", None))
        else:
            normalized_values.append(("value", float(value)))

    return len(set(normalized_values)) > 1


def find_duplicates(result: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    duplicate_mask = result.duplicated(
        subset=["Номер заказа"],
        keep=False
    )

    duplicates = result[duplicate_mask].copy()

    if duplicates.empty:
        return pd.DataFrame(columns=DUPLICATE_COLUMNS), []

    conflicting_numbers = []
    duplicate_types = {}
    duplicate_counts = {}

    for order_number, group in duplicates.groupby(
        "Номер заказа",
        sort=False,
        dropna=False
    ):
        duplicate_counts[order_number] = len(group)

        has_order_cost_conflict = series_has_different_values(
            group["Стоимость заказа"]
        )
        has_delivery_cost_conflict = series_has_different_values(
            group["Стоимость доставки"]
        )

        if has_order_cost_conflict or has_delivery_cost_conflict:
            conflicting_numbers.append(str(order_number))
            duplicate_types[order_number] = "Разные значения стоимости"
        else:
            duplicate_types[order_number] = "Одинаковые строки"

    duplicates["Количество строк с этим номером"] = (
        duplicates["Номер заказа"].map(duplicate_counts)
    )
    duplicates["Тип дубля"] = duplicates["Номер заказа"].map(duplicate_types)

    duplicates = duplicates[DUPLICATE_COLUMNS].sort_values(
        ["Номер заказа", "Строка в исходном файле"],
        kind="stable"
    ).reset_index(drop=True)

    return duplicates, conflicting_numbers


def read_and_prepare_report(
    uploaded_file,
    report_name: str
) -> tuple[pd.DataFrame, bool, pd.DataFrame, list[str]]:
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)

    order_column = find_column(df.columns, ORDER_NUMBER_ALIASES)
    order_cost_column = find_column(df.columns, ORDER_COST_ALIASES)
    delivery_cost_column = find_column(df.columns, DELIVERY_COST_ALIASES)

    missing = []

    if order_column is None:
        missing.append("номер заказа")

    if order_cost_column is None:
        missing.append("стоимость заказа")

    if missing:
        found_columns = ", ".join(map(str, df.columns))
        raise ValueError(
            f"В файле «{report_name}» не найдены обязательные столбцы: "
            f"{', '.join(missing)}. Найденные столбцы: {found_columns}"
        )

    result = pd.DataFrame({
        "Номер заказа": df[order_column].apply(normalize_order_number),
        "Стоимость заказа": df[order_cost_column].apply(parse_money),
        "Строка в исходном файле": df.index + 2,
    })

    has_delivery_cost = delivery_cost_column is not None

    if has_delivery_cost:
        result["Стоимость доставки"] = df[delivery_cost_column].apply(parse_money)
    else:
        result["Стоимость доставки"] = pd.NA

    result = result.dropna(subset=["Номер заказа"]).copy()

    duplicates, conflicting_duplicates = find_duplicates(result)

    # Для самого сравнения оставляем одну строку на номер заказа.
    # Полностью одинаковые дубли не мешают сравнению.
    # Дубли с разными значениями отдельно возвращаются в интерфейс,
    # а основное сравнение для таких файлов будет остановлено.
    prepared_result = (
        result
        .drop_duplicates(subset=["Номер заказа"], keep="first")
        [["Номер заказа", "Стоимость заказа", "Стоимость доставки"]]
        .reset_index(drop=True)
    )

    return (
        prepared_result,
        has_delivery_cost,
        duplicates,
        conflicting_duplicates
    )


def money_equal(left, right, tolerance: float = 0.01) -> bool:
    if pd.isna(left) and pd.isna(right):
        return True

    if pd.isna(left) or pd.isna(right):
        return False

    return abs(float(left) - float(right)) <= tolerance


def empty_comparison_result(
    client_df: pd.DataFrame,
    our_df: pd.DataFrame,
    client_has_delivery: bool,
    our_has_delivery: bool,
    client_duplicates: pd.DataFrame,
    our_duplicates: pd.DataFrame,
    client_conflicting_duplicates: list[str],
    our_conflicting_duplicates: list[str],
) -> dict:
    return {
        "client_only": pd.DataFrame(
            columns=["Номер заказа", "Стоимость заказа"]
        ),
        "our_only": pd.DataFrame(
            columns=["Номер заказа", "Стоимость заказа"]
        ),
        "order_cost_mismatches": pd.DataFrame(
            columns=[
                "Номер заказа",
                "Стоимость заказа клиента",
                "Стоимость заказа нашего отчёта",
            ]
        ),
        "delivery_cost_mismatches": pd.DataFrame(
            columns=[
                "Номер заказа",
                "Стоимость доставки клиента",
                "Стоимость доставки нашего отчёта",
            ]
        ),
        "delivery_comparison_available": client_has_delivery and our_has_delivery,
        "client_has_delivery": client_has_delivery,
        "our_has_delivery": our_has_delivery,
        "client_rows": len(client_df),
        "our_rows": len(our_df),
        "common_rows": 0,
        "client_duplicates": client_duplicates,
        "our_duplicates": our_duplicates,
        "client_duplicate_orders": client_duplicates["Номер заказа"].nunique(),
        "our_duplicate_orders": our_duplicates["Номер заказа"].nunique(),
        "client_conflicting_duplicates": client_conflicting_duplicates,
        "our_conflicting_duplicates": our_conflicting_duplicates,
        "comparison_blocked": True,
    }


def compare_reports(client_file, our_file) -> dict:
    (
        client_df,
        client_has_delivery,
        client_duplicates,
        client_conflicting_duplicates
    ) = read_and_prepare_report(
        client_file,
        "Отчёт клиента"
    )

    (
        our_df,
        our_has_delivery,
        our_duplicates,
        our_conflicting_duplicates
    ) = read_and_prepare_report(
        our_file,
        "Наш отчёт"
    )

    if client_conflicting_duplicates or our_conflicting_duplicates:
        return empty_comparison_result(
            client_df=client_df,
            our_df=our_df,
            client_has_delivery=client_has_delivery,
            our_has_delivery=our_has_delivery,
            client_duplicates=client_duplicates,
            our_duplicates=our_duplicates,
            client_conflicting_duplicates=client_conflicting_duplicates,
            our_conflicting_duplicates=our_conflicting_duplicates,
        )

    client_indexed = client_df.set_index("Номер заказа", drop=False)
    our_indexed = our_df.set_index("Номер заказа", drop=False)

    client_numbers = set(client_indexed.index)
    our_numbers = set(our_indexed.index)

    client_only_numbers = sorted(client_numbers - our_numbers)
    our_only_numbers = sorted(our_numbers - client_numbers)
    common_numbers = sorted(client_numbers & our_numbers)

    client_only = client_indexed.loc[
        client_only_numbers,
        ["Номер заказа", "Стоимость заказа"]
    ].reset_index(drop=True) if client_only_numbers else pd.DataFrame(
        columns=["Номер заказа", "Стоимость заказа"]
    )

    our_only = our_indexed.loc[
        our_only_numbers,
        ["Номер заказа", "Стоимость заказа"]
    ].reset_index(drop=True) if our_only_numbers else pd.DataFrame(
        columns=["Номер заказа", "Стоимость заказа"]
    )

    order_cost_mismatches = []
    delivery_cost_mismatches = []

    for order_number in common_numbers:
        client_row = client_indexed.loc[order_number]
        our_row = our_indexed.loc[order_number]

        if not money_equal(
            client_row["Стоимость заказа"],
            our_row["Стоимость заказа"]
        ):
            order_cost_mismatches.append({
                "Номер заказа": order_number,
                "Стоимость заказа клиента": client_row["Стоимость заказа"],
                "Стоимость заказа нашего отчёта": our_row["Стоимость заказа"],
            })

        if client_has_delivery and our_has_delivery:
            if not money_equal(
                client_row["Стоимость доставки"],
                our_row["Стоимость доставки"]
            ):
                delivery_cost_mismatches.append({
                    "Номер заказа": order_number,
                    "Стоимость доставки клиента": client_row["Стоимость доставки"],
                    "Стоимость доставки нашего отчёта": our_row["Стоимость доставки"],
                })

    return {
        "client_only": client_only,
        "our_only": our_only,
        "order_cost_mismatches": pd.DataFrame(
            order_cost_mismatches,
            columns=[
                "Номер заказа",
                "Стоимость заказа клиента",
                "Стоимость заказа нашего отчёта",
            ]
        ),
        "delivery_cost_mismatches": pd.DataFrame(
            delivery_cost_mismatches,
            columns=[
                "Номер заказа",
                "Стоимость доставки клиента",
                "Стоимость доставки нашего отчёта",
            ]
        ),
        "delivery_comparison_available": client_has_delivery and our_has_delivery,
        "client_has_delivery": client_has_delivery,
        "our_has_delivery": our_has_delivery,
        "client_rows": len(client_df),
        "our_rows": len(our_df),
        "common_rows": len(common_numbers),
        "client_duplicates": client_duplicates,
        "our_duplicates": our_duplicates,
        "client_duplicate_orders": client_duplicates["Номер заказа"].nunique(),
        "our_duplicate_orders": our_duplicates["Номер заказа"].nunique(),
        "client_conflicting_duplicates": client_conflicting_duplicates,
        "our_conflicting_duplicates": our_conflicting_duplicates,
        "comparison_blocked": False,
    }


def unique_orders_to_excel_bytes(client_only: pd.DataFrame, our_only: pd.DataFrame) -> bytes:
    output = BytesIO()

    max_rows = max(len(client_only), len(our_only))

    export_df = pd.DataFrame(index=range(max_rows))
    export_df["Номер заказа клиента"] = client_only["Номер заказа"].reindex(range(max_rows))
    export_df["Стоимость заказа клиента"] = client_only["Стоимость заказа"].reindex(range(max_rows))
    export_df[""] = ""
    export_df["Номер заказа нашего отчёта"] = our_only["Номер заказа"].reindex(range(max_rows))
    export_df["Стоимость заказа нашего отчёта"] = our_only["Стоимость заказа"].reindex(range(max_rows))

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Уникальные заказы")

        workbook = writer.book
        worksheet = writer.sheets["Уникальные заказы"]

        header_format = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#D9EAF7",
            "border": 1,
        })
        money_format = workbook.add_format({
            "num_format": "#,##0.00",
            "align": "center",
        })
        centered_format = workbook.add_format({"align": "center"})

        for column_number, column_name in enumerate(export_df.columns):
            worksheet.write(0, column_number, column_name, header_format)

        worksheet.set_column("A:A", 24, centered_format)
        worksheet.set_column("B:B", 24, money_format)
        worksheet.set_column("C:C", 4)
        worksheet.set_column("D:D", 28, centered_format)
        worksheet.set_column("E:E", 30, money_format)
        worksheet.freeze_panes(1, 0)

    return output.getvalue()


def format_duplicate_sheet(
    workbook,
    worksheet,
    duplicates_df: pd.DataFrame
) -> None:
    header_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "bg_color": "#FCE5CD",
        "border": 1,
    })
    money_format = workbook.add_format({
        "num_format": "#,##0.00",
        "align": "center",
    })
    centered_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
    })

    for column_number, column_name in enumerate(duplicates_df.columns):
        worksheet.write(0, column_number, column_name, header_format)

        if "Стоимость" in column_name:
            worksheet.set_column(
                column_number,
                column_number,
                24,
                money_format
            )
        elif column_name == "Тип дубля":
            worksheet.set_column(
                column_number,
                column_number,
                30,
                centered_format
            )
        else:
            worksheet.set_column(
                column_number,
                column_number,
                25,
                centered_format
            )

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(
        0,
        0,
        max(len(duplicates_df), 1),
        max(len(duplicates_df.columns) - 1, 0)
    )


def duplicates_to_excel_bytes(
    client_duplicates: pd.DataFrame,
    our_duplicates: pd.DataFrame
) -> bytes:
    output = BytesIO()

    client_export = client_duplicates.reindex(columns=DUPLICATE_COLUMNS)
    our_export = our_duplicates.reindex(columns=DUPLICATE_COLUMNS)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        client_export.to_excel(
            writer,
            index=False,
            sheet_name="Дубли клиента"
        )
        our_export.to_excel(
            writer,
            index=False,
            sheet_name="Дубли нашего отчёта"
        )

        workbook = writer.book

        format_duplicate_sheet(
            workbook,
            writer.sheets["Дубли клиента"],
            client_export
        )
        format_duplicate_sheet(
            workbook,
            writer.sheets["Дубли нашего отчёта"],
            our_export
        )

    return output.getvalue()


def dataframe_to_comparison_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        header_format = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#D9EAF7",
            "border": 1,
        })
        money_format = workbook.add_format({
            "num_format": "#,##0.00",
            "align": "center",
        })
        centered_format = workbook.add_format({"align": "center"})

        for column_number, column_name in enumerate(df.columns):
            worksheet.write(0, column_number, column_name, header_format)

            if "Стоимость" in column_name:
                worksheet.set_column(column_number, column_number, 32, money_format)
            else:
                worksheet.set_column(column_number, column_number, 24, centered_format)

        worksheet.freeze_panes(1, 0)

    return output.getvalue()
