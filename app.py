import os
from datetime import date

import pandas as pd
import streamlit as st

from database import (
    check_grouped_df_for_db,
    save_grouped_df_to_mysql,
    get_active_drivers,
    get_assignments_by_date,
    update_assignment_driver_and_status,
    bulk_assign_orders_to_driver
)

from processing import (
    process_delivery_file,
    dataframe_to_excel_bytes
)

from compare_reports import (
    compare_reports,
    unique_orders_to_excel_bytes,
    dataframe_to_comparison_excel_bytes
)


STATUS_LABELS = {
    0: "Не назначено",
    1: "Доставлено",
    2: "Не доставлен, но отгружен",
    3: "ТС подано, но не отгружено",
    4: "Отменена",
    5: "На водителе"
}

STATUS_COLORS = {
    0: "background-color: #eeeeee",
    1: "background-color: #d9ead3",
    2: "background-color: #fce5cd",
    3: "background-color: #ccffcc",
    4: "background-color: #f4cccc",
    5: "background-color: #ffffff"
}


st.set_page_config(
    page_title="Обработка доставок",
    layout="wide"
)

def require_site_password():
    site_password = os.getenv("SITE_PASSWORD")

    if not site_password:
        st.error("Пароль сайта не настроен. Добавьте SITE_PASSWORD в .env на сервере.")
        st.stop()

    if "site_authenticated" not in st.session_state:
        st.session_state["site_authenticated"] = False

    if st.session_state["site_authenticated"]:
        return

    st.title("Вход на сайт")

    password = st.text_input(
        "Введите пароль",
        type="password"
    )

    if st.button("Войти"):
        if password == site_password:
            st.session_state["site_authenticated"] = True
            st.rerun()
        else:
            st.error("Неверный пароль.")

    st.stop()


require_site_password()

st.markdown(
    """
    <style>
        div[data-testid="stDataFrame"] div[role="columnheader"] {
            font-size: 16px !important;
            font-weight: 800 !important;
            text-align: center !important;
            justify-content: center !important;
        }

        div[data-testid="stDataFrame"] div[role="gridcell"] {
            text-align: center !important;
            justify-content: center !important;
        }

        div[data-testid="stDataFrame"] th {
            font-size: 16px !important;
            font-weight: 800 !important;
            text-align: center !important;
        }

        div[data-testid="stDataFrame"] td {
            text-align: center !important;
        }
    </style>
    """,
    unsafe_allow_html=True
)


def center_styler(styler):
    return (
        styler
        .set_properties(**{
            "text-align": "center",
            "vertical-align": "middle"
        })
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("text-align", "center"),
                    ("font-weight", "800"),
                    ("font-size", "16px")
                ]
            },
            {
                "selector": "td",
                "props": [
                    ("text-align", "center"),
                    ("vertical-align", "middle")
                ]
            }
        ])
    )


def centered_dataframe(df: pd.DataFrame):
    return center_styler(df.style)


def add_status_label(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()

    if "status" in result.columns:
        result["status"] = pd.to_numeric(
            result["status"],
            errors="coerce"
        ).fillna(0).astype(int)

        result["Статус"] = result["status"].map(STATUS_LABELS).fillna("Неизвестный статус")

    return result


def style_status_rows(row):
    status = row.get("status", None)

    try:
        status = int(status)
    except (TypeError, ValueError):
        status = None

    css = STATUS_COLORS.get(status, "")

    if css:
        css = f"{css}; text-align: center; vertical-align: middle"
    else:
        css = "text-align: center; vertical-align: middle"

    return [css for _ in row]


def make_display_assignments_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Дата", "Номер заявки", "Водитель", "Статус"])

    display_df = add_status_label(df)

    display_df["Водитель"] = (
        display_df["driver_name"]
        .fillna("Не назначен")
        .replace("", "Не назначен")
    )

    display_df = display_df.rename(columns={
        "delivery_date": "Дата",
        "order_number": "Номер заявки"
    })

    return display_df[
        [
            "Дата",
            "Номер заявки",
            "Водитель",
            "Статус",
            "status"
        ]
    ].copy()


def styled_assignments_table(display_df: pd.DataFrame):
    styler = display_df.style.apply(
        style_status_rows,
        axis=1
    )

    styler = center_styler(styler)

    if "status" in display_df.columns:
        try:
            styler = styler.hide(axis="columns", subset=["status"])
        except (TypeError, AttributeError):
            pass

    return styler


def make_pool_table(assignments_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    assigned_df = assignments_df[
        (assignments_df["status"] != 0)
        & assignments_df["driver_name"].notna()
        & (assignments_df["driver_name"].astype(str).str.strip() != "")
    ].copy()

    if assigned_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    driver_names = sorted(assigned_df["driver_name"].unique())

    max_orders = 0

    for driver_name in driver_names:
        max_orders = max(
            max_orders,
            len(assigned_df[assigned_df["driver_name"] == driver_name])
        )

    pool_data = {}
    status_data = {}

    for driver_name in driver_names:
        driver_orders = assigned_df[assigned_df["driver_name"] == driver_name].copy()
        driver_orders = driver_orders.sort_values(["status", "order_number"])

        column_name = f"{driver_name} - {len(driver_orders)}"

        values = []
        statuses = []

        for _, row in driver_orders.iterrows():
            values.append(row["order_number"])
            statuses.append(int(row["status"]))

        while len(values) < max_orders:
            values.append("")
            statuses.append("empty")

        pool_data[column_name] = values
        status_data[column_name] = statuses

    return pd.DataFrame(pool_data), pd.DataFrame(status_data)


def style_pool_table(pool_df: pd.DataFrame, status_df: pd.DataFrame):
    styles = pd.DataFrame("", index=pool_df.index, columns=pool_df.columns)

    for row_index in pool_df.index:
        for col in pool_df.columns:
            status = status_df.loc[row_index, col]

            if status == "empty":
                styles.loc[row_index, col] = "text-align: center; vertical-align: middle"
            else:
                color = STATUS_COLORS.get(int(status), "")
                styles.loc[row_index, col] = f"{color}; text-align: center; vertical-align: middle"

    return styles


def parse_pasted_orders(text: str) -> list[str]:
    order_numbers = []

    for line in text.splitlines():
        line = line.strip()

        if line == "":
            continue

        parts = line.split("\t")

        for part in parts:
            cleaned = part.strip()

            if cleaned != "":
                order_numbers.append(cleaned)

    return order_numbers


def show_centered_table(df: pd.DataFrame, use_container_width=True, hide_index=True):
    if df.empty:
        st.dataframe(
            df,
            use_container_width=use_container_width,
            hide_index=hide_index
        )
    else:
        st.dataframe(
            centered_dataframe(df),
            use_container_width=use_container_width,
            hide_index=hide_index
        )


def render_assignments_page():
    st.title("Заявки")

    database_url = os.getenv("DATABASE_URL")

    if not database_url:
        st.error(
            "Переменная DATABASE_URL не настроена. "
            "Невозможно получить заявки из базы данных."
        )
        return

    selected_date = st.date_input(
        "Выберите дату доставки",
        value=date.today(),
        format="DD.MM.YYYY"
    )

    try:
        assignments_df = get_assignments_by_date(
            database_url=database_url,
            selected_date=selected_date
        )

        drivers_df = get_active_drivers(database_url=database_url)

    except Exception as error:
        st.error("Не удалось загрузить данные из БД.")
        st.exception(error)
        return

    if "bulk_assign_result" in st.session_state:
        result = st.session_state.pop("bulk_assign_result")

        if result.get("assigned"):
            st.success(f"Назначено заявок: {len(result['assigned'])}")
            show_centered_table(pd.DataFrame(result["assigned"]))

        if result.get("skipped"):
            st.warning(f"Пропущено заявок: {len(result['skipped'])}")
            show_centered_table(pd.DataFrame(result["skipped"]))

    if assignments_df.empty:
        st.info("На выбранную дату заявок в БД нет.")
        return

    assignments_df["status"] = pd.to_numeric(
        assignments_df["status"],
        errors="coerce"
    ).fillna(0).astype(int)

    st.subheader("Неназначенные заявки")

    unassigned_df = assignments_df[assignments_df["status"] == 0].copy()

    left_col, right_col = st.columns([2, 1])

    with left_col:
        if unassigned_df.empty:
            st.success("Неназначенных заявок на выбранную дату нет.")
        else:
            unassigned_display_df = make_display_assignments_df(unassigned_df)

            st.dataframe(
                styled_assignments_table(unassigned_display_df),
                use_container_width=True,
                hide_index=True
            )

    with right_col:
        st.write("Назначить водителя")

        if unassigned_df.empty:
            st.info("Нет заявок для назначения.")
        elif drivers_df.empty:
            st.warning("В таблице drivers нет активных водителей.")
        else:
            assignment_options = {}

            for _, row in unassigned_df.iterrows():
                label = f"{row['order_number']} | ID назначения: {row['assignment_id']}"
                assignment_options[label] = int(row["assignment_id"])

            driver_options = {}

            for _, row in drivers_df.iterrows():
                label = f"{row['full_name']} | ID: {row['id']}"
                driver_options[label] = int(row["id"])

            selected_assignment_label = st.selectbox(
                "Одна заявка",
                options=list(assignment_options.keys()),
                key="assign_assignment_select"
            )

            selected_driver_label = st.selectbox(
                "Водитель",
                options=list(driver_options.keys()),
                key="assign_driver_select"
            )

            if st.button("Назначить одну заявку", key="assign_driver_button"):
                try:
                    update_assignment_driver_and_status(
                        database_url=database_url,
                        assignment_id=assignment_options[selected_assignment_label],
                        driver_id=driver_options[selected_driver_label],
                        status=5
                    )

                    st.success("Водитель назначен. Статус изменён на «На водителе».")
                    st.rerun()

                except Exception as error:
                    st.error("Не удалось назначить водителя.")
                    st.exception(error)

            st.divider()

            st.write("Массовое назначение")

            selected_bulk_driver_label = st.selectbox(
                "Водитель для массового назначения",
                options=list(driver_options.keys()),
                key="bulk_driver_select"
            )

            multi_assignment_labels = st.multiselect(
                "Выбрать несколько заявок из списка",
                options=list(assignment_options.keys()),
                key="bulk_assignment_multiselect"
            )

            pasted_orders_text = st.text_area(
                "Или вставить номера заявок из Excel",
                placeholder="Г-00531403\nГ-00539497\nГ-00530250",
                height=150,
                key="bulk_orders_text_area"
            )

            if st.button("Назначить выбранные заявки", key="bulk_assign_button"):
                try:
                    order_numbers = []

                    for label in multi_assignment_labels:
                        order_number = label.split("|")[0].strip()
                        order_numbers.append(order_number)

                    pasted_order_numbers = parse_pasted_orders(pasted_orders_text)

                    order_numbers.extend(pasted_order_numbers)

                    result = bulk_assign_orders_to_driver(
                        database_url=database_url,
                        selected_date=selected_date,
                        driver_id=driver_options[selected_bulk_driver_label],
                        order_numbers=order_numbers
                    )

                    st.session_state["bulk_assign_result"] = result
                    st.rerun()

                except Exception as error:
                    st.error("Не удалось выполнить массовое назначение.")
                    st.exception(error)

    st.divider()

    st.subheader("Пул заявок по водителям")

    pool_df, pool_status_df = make_pool_table(assignments_df)

    if pool_df.empty:
        st.info("На выбранную дату пока нет заявок, назначенных на водителей.")
    else:
        styled_pool = pool_df.style.apply(
            lambda _: style_pool_table(pool_df, pool_status_df),
            axis=None
        )

        styled_pool = center_styler(styled_pool)

        st.dataframe(
            styled_pool,
            use_container_width=True,
            hide_index=True
        )

        pool_excel = dataframe_to_excel_bytes(pool_df)

        st.download_button(
            label="Скачать пул заявок по водителям",
            data=pool_excel,
            file_name=f"Пул_заявок_{selected_date.strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_pool_by_drivers"
        )

    st.divider()

    st.subheader("Все заявки за выбранную дату")

    all_assignments_display_df = make_display_assignments_df(assignments_df)

    st.dataframe(
        styled_assignments_table(all_assignments_display_df),
        use_container_width=True,
        hide_index=True
    )

    all_assignments_export_df = all_assignments_display_df.drop(
        columns=["status"],
        errors="ignore"
    )

    all_assignments_excel = dataframe_to_excel_bytes(all_assignments_export_df)

    st.download_button(
        label="Скачать заявки за выбранную дату",
        data=all_assignments_excel,
        file_name=f"Заявки_{selected_date.strftime('%d_%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_assignments_by_date"
    )

    st.divider()

    st.subheader("Изменить водителя или статус заявки")

    changeable_assignments_df = assignments_df[assignments_df["status"] != 0].copy()

    if changeable_assignments_df.empty:
        st.info("На выбранную дату нет заявок со статусом, отличным от 0.")
        return

    assignment_change_options = {}

    for _, row in changeable_assignments_df.iterrows():
        current_driver = row["driver_name"]

        if pd.isna(current_driver) or str(current_driver).strip() == "":
            current_driver = "Не назначен"

        status_label = STATUS_LABELS.get(int(row["status"]), "Неизвестный статус")

        label = (
            f"{row['order_number']} | {current_driver} | "
            f"{status_label} | ID назначения: {row['assignment_id']}"
        )

        assignment_change_options[label] = int(row["assignment_id"])

    selected_change_assignment_label = st.selectbox(
        "Заявка для изменения",
        options=list(assignment_change_options.keys()),
        key="change_assignment_select"
    )

    selected_assignment_id = assignment_change_options[selected_change_assignment_label]

    selected_assignment_row = changeable_assignments_df[
        changeable_assignments_df["assignment_id"] == selected_assignment_id
    ].iloc[0]

    current_status = int(selected_assignment_row["status"])

    status_options = {
        label: status_code
        for status_code, label in STATUS_LABELS.items()
        if status_code != 0
    }

    status_labels = list(status_options.keys())

    default_status_label = STATUS_LABELS.get(current_status, "На водителе")

    if default_status_label not in status_labels:
        default_status_label = "На водителе"

    selected_status_label = st.selectbox(
        "Новый статус",
        options=status_labels,
        index=status_labels.index(default_status_label),
        key="change_status_select"
    )

    driver_change_options = {
        "Не менять водителя": None
    }

    if not drivers_df.empty:
        for _, row in drivers_df.iterrows():
            label = f"{row['full_name']} | ID: {row['id']}"
            driver_change_options[label] = int(row["id"])

    selected_driver_change_label = st.selectbox(
        "Новый водитель",
        options=list(driver_change_options.keys()),
        key="change_driver_select"
    )

    if st.button("Сохранить изменения", key="save_assignment_changes"):
        try:
            selected_driver_id = driver_change_options[selected_driver_change_label]
            selected_status = status_options[selected_status_label]

            update_assignment_driver_and_status(
                database_url=database_url,
                assignment_id=selected_assignment_id,
                driver_id=selected_driver_id,
                status=selected_status
            )

            st.success("Изменения сохранены.")
            st.rerun()

        except Exception as error:
            st.error("Не удалось сохранить изменения.")
            st.exception(error)


def render_registry_page():
    st.title("Работа с реестром")

    st.write(
        "Загрузите основной файл доставок. "
        "Файл с водителями можно загрузить дополнительно, если нужно добавить водителей в сгруппированный файл."
    )

    main_file = st.file_uploader(
        "Основной файл доставок",
        type=["xlsx"],
        key="main_file"
    )

    drivers_file = st.file_uploader(
        "Файл с водителями, необязательно",
        type=["xlsx"],
        key="drivers_file"
    )

    st.info(
        "Файл с водителями должен содержать ровно два столбца: "
        "«ФИО водителя» и «Номер заявки». "
        "Названия столбцов должны быть именно такими. "
        "В номерах заявок не должно быть лишних пробелов и прочих символов."
    )

    if main_file is not None:
        st.success(f"Основной файл загружен: {main_file.name}")

    if drivers_file is not None:
        st.success(f"Файл с водителями загружен: {drivers_file.name}")
    else:
        st.warning(
            "Файл с водителями не загружен. "
            "Сгруппированный и урезанный файлы будут сформированы без колонки «Водитель»."
        )

    if st.button("Обработать файл"):
        if main_file is None:
            st.error("Сначала загрузите основной файл доставок.")
        else:
            try:
                main_file.seek(0)

                if drivers_file is not None:
                    drivers_file.seek(0)

                (
                    full_df,
                    grouped_df,
                    short_df,
                    full_filename,
                    grouped_filename,
                    short_filename
                ) = process_delivery_file(
                    main_file=main_file,
                    drivers_file=drivers_file,
                    original_filename=main_file.name
                )

                st.session_state["full_df"] = full_df
                st.session_state["grouped_df"] = grouped_df
                st.session_state["short_df"] = short_df
                st.session_state["full_filename"] = full_filename
                st.session_state["grouped_filename"] = grouped_filename
                st.session_state["short_filename"] = short_filename
                st.session_state["source_filename"] = main_file.name

                st.session_state.pop("db_check_result", None)
                st.session_state.pop("db_save_result", None)

                st.success("Файл успешно обработан.")

            except Exception as error:
                st.error("Произошла ошибка при обработке файла.")
                st.exception(error)

    if "grouped_df" in st.session_state:
        full_df = st.session_state["full_df"]
        grouped_df = st.session_state["grouped_df"]
        short_df = st.session_state["short_df"]

        full_filename = st.session_state["full_filename"]
        grouped_filename = st.session_state["grouped_filename"]
        short_filename = st.session_state["short_filename"]
        source_filename = st.session_state["source_filename"]

        full_excel = dataframe_to_excel_bytes(full_df)
        grouped_excel = dataframe_to_excel_bytes(grouped_df)
        short_excel = dataframe_to_excel_bytes(short_df)

        st.subheader("Скачать полный обработанный файл")

        st.download_button(
            label="Скачать полный файл",
            data=full_excel,
            file_name=full_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Сгруппированный файл")
        st.write(f"Строк: {len(grouped_df)}")
        show_centered_table(grouped_df)

        st.download_button(
            label="Скачать сгруппированный файл",
            data=grouped_excel,
            file_name=grouped_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Урезанный файл")
        st.write(f"Строк: {len(short_df)}")
        show_centered_table(short_df)

        st.download_button(
            label="Скачать урезанный файл",
            data=short_excel,
            file_name=short_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if "Водитель" in grouped_df.columns:
            no_driver_count = grouped_df["Водитель"].isna().sum()

            if no_driver_count > 0:
                st.warning(
                    f"Для {no_driver_count} заявок не найден водитель. "
                    "Проверьте номера заявок в файле водителей."
                )

        st.divider()

        st.subheader("Запись сгруппированных данных в базу данных")

        st.write(
            "В базу записывается сгруппированный файл без столбца «Водитель». "
            "Перед записью выполняется проверка уникальности по паре «Номер заявки + Комментарий»."
        )

        database_url = os.getenv("DATABASE_URL")

        if not database_url:
            st.error(
                "Переменная DATABASE_URL не настроена. "
                "Запись в базу данных сейчас недоступна."
            )
        else:
            if st.button("Проверить и записать в БД"):
                try:
                    db_check_result = check_grouped_df_for_db(
                        grouped_df=grouped_df,
                        database_url=database_url,
                        source_file=source_filename
                    )

                    st.session_state["db_check_result"] = db_check_result
                    st.session_state.pop("db_save_result", None)

                    if db_check_result["status"] == "ok":
                        db_save_result = save_grouped_df_to_mysql(
                            grouped_df=grouped_df,
                            database_url=database_url,
                            source_file=source_filename,
                            confirm_repeated_orders=False
                        )

                        st.session_state["db_save_result"] = db_save_result

                except Exception as error:
                    st.error("Ошибка при проверке или записи данных в БД.")
                    st.exception(error)

            if "db_check_result" in st.session_state:
                db_check_result = st.session_state["db_check_result"]

                st.write(f"Строк к проверке: {db_check_result['rows_count']}")

                if db_check_result["status"] == "ok":
                    st.success("Проверка пройдена. Все строки записаны в БД.")

                if db_check_result["passed_rows"]:
                    st.success("Строки, которые прошли проверку")
                    show_centered_table(pd.DataFrame(db_check_result["passed_rows"]))

                if db_check_result["incoming_duplicates"]:
                    st.error(
                        "Внутри загружаемого файла есть дубли по паре "
                        "«Номер заявки + Комментарий». Запись в БД заблокирована."
                    )
                    show_centered_table(pd.DataFrame(db_check_result["incoming_duplicates"]))

                if db_check_result["exact_duplicates"]:
                    st.error(
                        "В базе уже есть строки с такой же парой "
                        "«Номер заявки + Комментарий». Запись в БД заблокирована."
                    )
                    show_centered_table(pd.DataFrame(db_check_result["exact_duplicates"]))

                if db_check_result["same_order_different_comment"]:
                    st.warning(
                        "В базе уже есть заявки с таким же номером, "
                        "но с другим комментарием. Это может быть возврат, повторная доставка "
                        "или исправленная заявка."
                    )
                    show_centered_table(pd.DataFrame(db_check_result["same_order_different_comment"]))

                if db_check_result["status"] == "blocked":
                    st.error(
                        "Файл не был записан в БД. "
                        "Нужно удалить дубли или исправить комментарии."
                    )

                elif db_check_result["status"] == "needs_confirmation":
                    st.warning(
                        "Файл пока не записан в БД, потому что есть заявки с уже существующим "
                        "номером, но другим комментарием."
                    )

                    confirm_repeated_orders = st.checkbox(
                        "Подтверждаю, что заявки с таким же номером, но другим комментарием нужно записать как отдельные повторные заявки."
                    )

                    if st.button("Записать в БД с подтверждением"):
                        if not confirm_repeated_orders:
                            st.error("Перед записью нужно поставить подтверждение.")
                        else:
                            try:
                                db_save_result = save_grouped_df_to_mysql(
                                    grouped_df=grouped_df,
                                    database_url=database_url,
                                    source_file=source_filename,
                                    confirm_repeated_orders=True
                                )

                                st.session_state["db_save_result"] = db_save_result

                            except Exception as error:
                                st.error("Ошибка при записи данных в БД.")
                                st.exception(error)

            if "db_save_result" in st.session_state:
                db_save_result = st.session_state["db_save_result"]

                if db_save_result.get("saved"):
                    st.success(
                        f"Данные успешно записаны в БД. "
                        f"ID загрузки: {db_save_result['upload_id']}. "
                        f"Добавлено строк: {db_save_result['inserted_rows']}."
                    )
                else:
                    st.error("Данные не были записаны в БД.")
                    st.write(db_save_result)


def render_report_comparison_page():
    st.title("Сравнение отчётов")

    st.write(
        "Загрузите отчёт клиента и наш отчёт. "
        "В каждом файле обязательны столбцы с номером заказа и стоимостью заказа. "
        "Столбец «Стоимость доставки» используется дополнительно, если он есть в обоих файлах."
    )

    st.info(
        "Для номера заказа поддерживаются заголовки: «номер заказа», «номер заявки» и «заявка». "
        "Стоимость заказа должна находиться в столбце «стоимость заказа»."
    )

    left_column, right_column = st.columns(2)

    with left_column:
        client_report_file = st.file_uploader(
            "Отчёт клиента",
            type=["xlsx", "xls"],
            key="client_report_file"
        )

        if client_report_file is not None:
            st.success(f"Отчёт клиента загружен: {client_report_file.name}")

    with right_column:
        our_report_file = st.file_uploader(
            "Наш отчёт",
            type=["xlsx", "xls"],
            key="our_report_file"
        )

        if our_report_file is not None:
            st.success(f"Наш отчёт загружен: {our_report_file.name}")

    if st.button("Сравнить отчёты", key="compare_reports_button"):
        if client_report_file is None or our_report_file is None:
            st.error("Загрузите оба файла перед сравнением.")
        else:
            try:
                result = compare_reports(
                    client_file=client_report_file,
                    our_file=our_report_file
                )

                st.session_state["report_comparison_result"] = result
                st.success("Сравнение завершено.")

            except Exception as error:
                st.session_state.pop("report_comparison_result", None)
                st.error("Отчёты не удалось сравнить.")
                st.exception(error)

    if "report_comparison_result" not in st.session_state:
        return

    result = st.session_state["report_comparison_result"]

    metrics_columns = st.columns(4)
    metrics_columns[0].metric("Строк у клиента", result["client_rows"])
    metrics_columns[1].metric("Строк в нашем отчёте", result["our_rows"])
    metrics_columns[2].metric("Совпавших номеров", result["common_rows"])
    metrics_columns[3].metric(
        "Расхождений по стоимости заказа",
        len(result["order_cost_mismatches"])
    )

    st.divider()
    st.subheader("Уникальные номера заказов")

    unique_left, unique_right = st.columns(2)

    with unique_left:
        st.write(f"Только в отчёте клиента: {len(result['client_only'])}")
        show_centered_table(result["client_only"])

    with unique_right:
        st.write(f"Только в нашем отчёте: {len(result['our_only'])}")
        show_centered_table(result["our_only"])

    unique_excel = unique_orders_to_excel_bytes(
        client_only=result["client_only"],
        our_only=result["our_only"]
    )

    st.download_button(
        label="Скачать уникальные номера заказов",
        data=unique_excel,
        file_name="Уникальные_номера_заказов.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_unique_orders"
    )

    st.divider()
    st.subheader("Расхождения по стоимости заказа")

    if result["order_cost_mismatches"].empty:
        st.success("Расхождений по стоимости заказа не найдено.")
    else:
        st.warning(
            f"Найдено расхождений: {len(result['order_cost_mismatches'])}. "
            "Эталонным считается значение из отчёта клиента."
        )
        show_centered_table(result["order_cost_mismatches"])

    order_cost_excel = dataframe_to_comparison_excel_bytes(
        result["order_cost_mismatches"],
        "Стоимость заказа"
    )

    st.download_button(
        label="Скачать расхождения по стоимости заказа",
        data=order_cost_excel,
        file_name="Расхождения_стоимости_заказа.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_order_cost_mismatches"
    )

    st.divider()
    st.subheader("Расхождения по стоимости доставки")

    if not result["delivery_comparison_available"]:
        missing_files = []

        if not result["client_has_delivery"]:
            missing_files.append("отчёте клиента")

        if not result["our_has_delivery"]:
            missing_files.append("нашем отчёте")

        st.info(
            "Сравнение стоимости доставки не выполнено: столбец «стоимость доставки» "
            f"отсутствует в {' и '.join(missing_files)}."
        )
    else:
        if result["delivery_cost_mismatches"].empty:
            st.success("Расхождений по стоимости доставки не найдено.")
        else:
            st.warning(
                f"Найдено расхождений: {len(result['delivery_cost_mismatches'])}."
            )
            show_centered_table(result["delivery_cost_mismatches"])

        delivery_cost_excel = dataframe_to_comparison_excel_bytes(
            result["delivery_cost_mismatches"],
            "Стоимость доставки"
        )

        st.download_button(
            label="Скачать расхождения по стоимости доставки",
            data=delivery_cost_excel,
            file_name="Расхождения_стоимости_доставки.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_delivery_cost_mismatches"
        )


assignments_tab, registry_tab, comparison_tab = st.tabs([
    "Заявки",
    "Работа с реестром",
    "Сравнение отчётов"
])

with assignments_tab:
    render_assignments_page()

with registry_tab:
    render_registry_page()

with comparison_tab:
    render_report_comparison_page()
