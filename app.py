import os

import pandas as pd
import streamlit as st

from database import (
    check_grouped_df_for_db,
    save_grouped_df_to_mysql
)

from processing import (
    process_delivery_file,
    dataframe_to_excel_bytes
)


st.set_page_config(
    page_title="Обработка доставок",
    layout="wide"
)

st.title("Обработка Excel-файлов доставок")

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
    st.dataframe(grouped_df, use_container_width=True)

    st.download_button(
        label="Скачать сгруппированный файл",
        data=grouped_excel,
        file_name=grouped_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("Урезанный файл")
    st.write(f"Строк: {len(short_df)}")
    st.dataframe(short_df, use_container_width=True)

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
                st.dataframe(
                    pd.DataFrame(db_check_result["passed_rows"]),
                    use_container_width=True
                )

            if db_check_result["incoming_duplicates"]:
                st.error(
                    "Внутри загружаемого файла есть дубли по паре "
                    "«Номер заявки + Комментарий». Запись в БД заблокирована."
                )
                st.dataframe(
                    pd.DataFrame(db_check_result["incoming_duplicates"]),
                    use_container_width=True
                )

            if db_check_result["exact_duplicates"]:
                st.error(
                    "В базе уже есть строки с такой же парой "
                    "«Номер заявки + Комментарий». Запись в БД заблокирована."
                )
                st.dataframe(
                    pd.DataFrame(db_check_result["exact_duplicates"]),
                    use_container_width=True
                )

            if db_check_result["same_order_different_comment"]:
                st.warning(
                    "В базе уже есть заявки с таким же номером, "
                    "но с другим комментарием. Это может быть возврат, повторная доставка "
                    "или исправленная заявка."
                )
                st.dataframe(
                    pd.DataFrame(db_check_result["same_order_different_comment"]),
                    use_container_width=True
                )

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
