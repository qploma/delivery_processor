import hashlib
import re

import pandas as pd
from sqlalchemy import create_engine, text


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""

    value = str(value).replace("\xa0", " ").strip()
    value = re.sub(r"\s+", " ", value)

    return value


def normalize_comment_for_hash(value) -> str:
    return normalize_text(value).lower()


def make_comment_hash(comment) -> str:
    normalized_comment = normalize_comment_for_hash(comment)
    return hashlib.sha256(normalized_comment.encode("utf-8")).hexdigest()


def parse_date(value):
    if pd.isna(value):
        return None

    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")

    if pd.isna(parsed):
        return None

    return parsed.date()


def parse_decimal(value):
    if pd.isna(value):
        return None

    value_str = str(value).replace("\xa0", " ").strip()

    if value_str == "":
        return None

    value_str = value_str.replace(" ", "").replace(",", ".")

    try:
        return float(value_str)
    except ValueError:
        return None


def make_display_rows(df: pd.DataFrame) -> list[dict]:
    if df.empty:
        return []

    display_df = df.copy()

    display_df = display_df.rename(columns={
        "delivery_date": "Дата доставки",
        "order_number": "Номер заявки",
        "branch": "Филиал",
        "client_type": "Тип клиента",
        "delivery_address": "Адрес доставки",
        "client_phone": "Телефон клиента",
        "time_from": "Время С",
        "time_to": "Время ПО",
        "payment_method": "Способ оплаты",
        "order_price": "Стоимость заказа, руб.",
        "order_weight": "Вес заказа",
        "products": "Товары заявки",
        "order_comment": "Комментарий",
        "zone": "Зона"
    })

    columns = [
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

    columns = [col for col in columns if col in display_df.columns]

    return display_df[columns].to_dict("records")


def prepare_grouped_df_for_db(grouped_df: pd.DataFrame, source_file: str) -> pd.DataFrame:
    df = grouped_df.copy()

    # Водителя специально НЕ переносим в deliveries.
    # Водители живут отдельно в drivers и delivery_assignments.
    df = df.rename(columns={
        "Дата доставки": "delivery_date",
        "Номер заявки": "order_number",
        "Филиал": "branch",
        "Тип клиента": "client_type",
        "Адрес доставки": "delivery_address",
        "Телефон клиента": "client_phone",
        "Время С": "time_from",
        "Время ПО": "time_to",
        "Способ оплаты": "payment_method",
        "Стоимость заказа, руб.": "order_price",
        "Вес заказа": "order_weight",
        "Товары заявки": "products",
        "Комментарий": "order_comment",
        "Зона": "zone"
    })

    required_columns = [
        "delivery_date",
        "order_number",
        "branch",
        "client_type",
        "delivery_address",
        "client_phone",
        "time_from",
        "time_to",
        "payment_method",
        "order_price",
        "order_weight",
        "products",
        "order_comment",
        "zone"
    ]

    for col in required_columns:
        if col not in df.columns:
            df[col] = None

    df = df[required_columns].copy()

    df["delivery_date"] = df["delivery_date"].apply(parse_date)
    df["order_price"] = df["order_price"].apply(parse_decimal)
    df["order_weight"] = df["order_weight"].apply(parse_decimal)

    text_columns = [
        "order_number",
        "branch",
        "client_type",
        "delivery_address",
        "client_phone",
        "time_from",
        "time_to",
        "payment_method",
        "products",
        "order_comment",
        "zone"
    ]

    for col in text_columns:
        df[col] = df[col].apply(normalize_text)

    df = df[df["order_number"] != ""].copy()

    df["order_comment_hash"] = df["order_comment"].apply(make_comment_hash)
    df["source_file"] = source_file

    df = df.astype(object).where(pd.notna(df), None)

    return df


def get_engine(database_url: str):
    return create_engine(
        database_url,
        pool_pre_ping=True,
        pool_recycle=50
    )


def check_grouped_df_for_db(grouped_df: pd.DataFrame, database_url: str, source_file: str) -> dict:
    df_to_db = prepare_grouped_df_for_db(grouped_df, source_file)

    incoming_duplicates_df = df_to_db[
        df_to_db.duplicated(
            subset=["order_number", "order_comment_hash"],
            keep=False
        )
    ].copy()

    incoming_duplicates = incoming_duplicates_df[
        ["order_number", "order_comment"]
    ].rename(columns={
        "order_number": "Номер заявки",
        "order_comment": "Комментарий"
    }).to_dict("records")

    incoming_duplicate_pairs = set(
        zip(
            incoming_duplicates_df["order_number"],
            incoming_duplicates_df["order_comment_hash"]
        )
    )

    engine = get_engine(database_url)

    exact_duplicates = []
    exact_duplicate_pairs = set()

    same_order_different_comment = []
    same_order_different_comment_pairs = set()

    with engine.begin() as conn:
        for _, row in df_to_db.iterrows():
            exact_duplicate = conn.execute(
                text("""
                    SELECT
                        id,
                        order_number,
                        order_comment,
                        source_file,
                        created_at
                    FROM deliveries
                    WHERE order_number = :order_number
                      AND order_comment_hash = :order_comment_hash
                    LIMIT 1;
                """),
                {
                    "order_number": row["order_number"],
                    "order_comment_hash": row["order_comment_hash"]
                }
            ).mappings().first()

            if exact_duplicate is not None:
                exact_duplicates.append({
                    "Номер заявки": row["order_number"],
                    "Комментарий в новом файле": row["order_comment"],
                    "ID в БД": exact_duplicate["id"],
                    "Файл в БД": exact_duplicate["source_file"],
                    "Дата записи в БД": exact_duplicate["created_at"]
                })

                exact_duplicate_pairs.add(
                    (row["order_number"], row["order_comment_hash"])
                )

                continue

            different_comment_rows = conn.execute(
                text("""
                    SELECT
                        id,
                        order_number,
                        order_comment,
                        source_file,
                        created_at
                    FROM deliveries
                    WHERE order_number = :order_number
                      AND order_comment_hash <> :order_comment_hash;
                """),
                {
                    "order_number": row["order_number"],
                    "order_comment_hash": row["order_comment_hash"]
                }
            ).mappings().all()

            for old_row in different_comment_rows:
                same_order_different_comment.append({
                    "Номер заявки": row["order_number"],
                    "Новый комментарий": row["order_comment"],
                    "Старый комментарий в БД": old_row["order_comment"],
                    "ID в БД": old_row["id"],
                    "Файл в БД": old_row["source_file"],
                    "Дата записи в БД": old_row["created_at"]
                })

                same_order_different_comment_pairs.add(
                    (row["order_number"], row["order_comment_hash"])
                )

    problem_pairs = (
        incoming_duplicate_pairs
        | exact_duplicate_pairs
        | same_order_different_comment_pairs
    )

    if problem_pairs:
        passed_df = df_to_db[
            ~df_to_db.apply(
                lambda row: (row["order_number"], row["order_comment_hash"]) in problem_pairs,
                axis=1
            )
        ].copy()
    else:
        passed_df = df_to_db.copy()

    passed_rows = make_display_rows(passed_df)

    has_blocking_errors = len(incoming_duplicates) > 0 or len(exact_duplicates) > 0
    needs_confirmation = len(same_order_different_comment) > 0 and not has_blocking_errors

    if has_blocking_errors:
        status = "blocked"
    elif needs_confirmation:
        status = "needs_confirmation"
    else:
        status = "ok"

    return {
        "status": status,
        "rows_count": len(df_to_db),
        "passed_rows": passed_rows,
        "incoming_duplicates": incoming_duplicates,
        "exact_duplicates": exact_duplicates,
        "same_order_different_comment": same_order_different_comment
    }


def save_grouped_df_to_mysql(
    grouped_df: pd.DataFrame,
    database_url: str,
    source_file: str,
    confirm_repeated_orders: bool = False
) -> dict:
    check_result = check_grouped_df_for_db(
        grouped_df=grouped_df,
        database_url=database_url,
        source_file=source_file
    )

    if check_result["status"] == "blocked":
        return {
            "saved": False,
            "reason": "blocked",
            "check_result": check_result,
            "inserted_rows": 0
        }

    if check_result["status"] == "needs_confirmation" and not confirm_repeated_orders:
        return {
            "saved": False,
            "reason": "needs_confirmation",
            "check_result": check_result,
            "inserted_rows": 0
        }

    df_to_db = prepare_grouped_df_for_db(grouped_df, source_file)

    repeated_order_numbers = {
        item["Номер заявки"]
        for item in check_result["same_order_different_comment"]
    }

    engine = get_engine(database_url)

    with engine.begin() as conn:
        upload_result = conn.execute(
            text("""
                INSERT INTO upload_batches (
                    source_file,
                    rows_count,
                    created_by
                )
                VALUES (
                    :source_file,
                    :rows_count,
                    :created_by
                );
            """),
            {
                "source_file": source_file,
                "rows_count": len(df_to_db),
                "created_by": "site"
            }
        )

        upload_id = upload_result.lastrowid

        inserted_rows = 0

        for _, row in df_to_db.iterrows():
            is_confirmed_duplicate = 1 if row["order_number"] in repeated_order_numbers else 0

            delivery_result = conn.execute(
                text("""
                    INSERT INTO deliveries (
                        upload_id,
                        order_number,
                        delivery_date,
                        branch,
                        client_type,
                        delivery_address,
                        client_phone,
                        time_from,
                        time_to,
                        payment_method,
                        order_price,
                        order_weight,
                        products,
                        order_comment,
                        order_comment_hash,
                        is_confirmed_duplicate,
                        zone,
                        source_file
                    )
                    VALUES (
                        :upload_id,
                        :order_number,
                        :delivery_date,
                        :branch,
                        :client_type,
                        :delivery_address,
                        :client_phone,
                        :time_from,
                        :time_to,
                        :payment_method,
                        :order_price,
                        :order_weight,
                        :products,
                        :order_comment,
                        :order_comment_hash,
                        :is_confirmed_duplicate,
                        :zone,
                        :source_file
                    );
                """),
                {
                    "upload_id": upload_id,
                    "order_number": row["order_number"],
                    "delivery_date": row["delivery_date"],
                    "branch": row["branch"],
                    "client_type": row["client_type"],
                    "delivery_address": row["delivery_address"],
                    "client_phone": row["client_phone"],
                    "time_from": row["time_from"],
                    "time_to": row["time_to"],
                    "payment_method": row["payment_method"],
                    "order_price": row["order_price"],
                    "order_weight": row["order_weight"],
                    "products": row["products"],
                    "order_comment": row["order_comment"],
                    "order_comment_hash": row["order_comment_hash"],
                    "is_confirmed_duplicate": is_confirmed_duplicate,
                    "zone": row["zone"],
                    "source_file": source_file
                }
            )

            delivery_id = delivery_result.lastrowid

            conn.execute(
                text("""
                    INSERT INTO delivery_assignments (
                        delivery_id,
                        order_number,
                        delivery_date,
                        driver_id,
                        driver_name,
                        assignment_date,
                        status,
                        status_comment
                    )
                    VALUES (
                        :delivery_id,
                        :order_number,
                        :delivery_date,
                        NULL,
                        NULL,
                        NULL,
                        0,
                        NULL
                    );
                """),
                {
                    "delivery_id": delivery_id,
                    "order_number": row["order_number"],
                    "delivery_date": row["delivery_date"]
                }
            )

            inserted_rows += 1

    return {
        "saved": True,
        "reason": "saved",
        "upload_id": upload_id,
        "inserted_rows": inserted_rows,
        "check_result": check_result
    }


def get_active_drivers(database_url: str) -> pd.DataFrame:
    engine = get_engine(database_url)

    with engine.begin() as conn:
        rows = conn.execute(
            text("""
                SELECT
                    id,
                    full_name,
                    phone,
                    is_active
                FROM drivers
                WHERE is_active = 1
                ORDER BY full_name;
            """)
        ).mappings().all()

    return pd.DataFrame(rows)


def get_assignments_by_date(database_url: str, selected_date) -> pd.DataFrame:
    delivery_date = parse_date(selected_date)

    columns = [
        "assignment_id",
        "delivery_id",
        "delivery_date",
        "order_number",
        "driver_id",
        "driver_name",
        "status",
        "status_comment"
    ]

    if delivery_date is None:
        return pd.DataFrame(columns=columns)

    engine = get_engine(database_url)

    with engine.begin() as conn:
        rows = conn.execute(
            text("""
                SELECT
                    id AS assignment_id,
                    delivery_id,
                    delivery_date,
                    order_number,
                    driver_id,
                    driver_name,
                    status,
                    status_comment
                FROM delivery_assignments
                WHERE delivery_date = :delivery_date
                ORDER BY
                    status ASC,
                    driver_name ASC,
                    order_number ASC;
            """),
            {
                "delivery_date": delivery_date
            }
        ).mappings().all()

    if not rows:
        return pd.DataFrame(columns=columns)

    return pd.DataFrame(rows)


def update_assignment_driver_and_status(
    database_url: str,
    assignment_id: int,
    driver_id=None,
    status=None
) -> None:
    engine = get_engine(database_url)

    with engine.begin() as conn:
        if driver_id is not None:
            driver = conn.execute(
                text("""
                    SELECT
                        id,
                        full_name
                    FROM drivers
                    WHERE id = :driver_id
                      AND is_active = 1
                    LIMIT 1;
                """),
                {
                    "driver_id": int(driver_id)
                }
            ).mappings().first()

            if driver is None:
                raise ValueError("Водитель не найден или неактивен.")

            new_status = 5 if status is None else int(status)

            conn.execute(
                text("""
                    UPDATE delivery_assignments
                    SET
                        driver_id = :driver_id,
                        driver_name = :driver_name,
                        status = :status,
                        assignment_date = CURRENT_DATE
                    WHERE id = :assignment_id;
                """),
                {
                    "driver_id": driver["id"],
                    "driver_name": driver["full_name"],
                    "status": new_status,
                    "assignment_id": int(assignment_id)
                }
            )

        else:
            if status is None:
                return

            conn.execute(
                text("""
                    UPDATE delivery_assignments
                    SET
                        status = :status
                    WHERE id = :assignment_id;
                """),
                {
                    "status": int(status),
                    "assignment_id": int(assignment_id)
                }
            )
