import pandas as pd
import streamlit as st

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
    "Загрузите основной файл доставок и файл с водителями. "
    "На выходе сайт сформирует полный обработанный файл и сгруппированный файл с зонами, товарами и водителями."
)

main_file = st.file_uploader(
    "Основной файл доставок",
    type=["xlsx"],
    key="main_file"
)

drivers_file = st.file_uploader(
    "Файл с водителями",
    type=["xlsx"],
    key="drivers_file"
)

# --------------------------------------------------
# Предпросмотр исходного файла с товарами
# --------------------------------------------------

if main_file is not None:
    st.info(f"Основной файл загружен: {main_file.name}")

    try:
        main_file.seek(0)

        raw_main_df = pd.read_excel(
            main_file,
            sheet_name="TDSheet"
        )

        st.subheader("Исходный файл с товарами")
        st.write(f"Строк: {raw_main_df.shape[0]}, столбцов: {raw_main_df.shape[1]}")
        st.dataframe(raw_main_df, use_container_width=True)

        main_file.seek(0)

    except Exception as error:
        st.error("Не удалось показать исходный файл с товарами.")
        st.exception(error)

# --------------------------------------------------
# Предпросмотр исходного файла с водителями
# --------------------------------------------------

if drivers_file is not None:
    st.info(f"Файл с водителями загружен: {drivers_file.name}")

    try:
        drivers_file.seek(0)

        raw_drivers_df = pd.read_excel(drivers_file)

        st.subheader("Исходный файл с водителями")
        st.write(f"Строк: {raw_drivers_df.shape[0]}, столбцов: {raw_drivers_df.shape[1]}")
        st.dataframe(raw_drivers_df, use_container_width=True)

        drivers_file.seek(0)

    except Exception as error:
        st.error("Не удалось показать исходный файл с водителями.")
        st.exception(error)

# --------------------------------------------------
# Обработка файлов
# --------------------------------------------------

if st.button("Обработать файлы"):
    if main_file is None:
        st.error("Сначала загрузите основной файл доставок.")
    elif drivers_file is None:
        st.error("Сначала загрузите файл с водителями.")
    else:
        try:
            main_file.seek(0)
            drivers_file.seek(0)

            full_df, grouped_df, full_filename, grouped_filename = process_delivery_file(
                main_file=main_file,
                drivers_file=drivers_file,
                original_filename=main_file.name
            )

            full_excel = dataframe_to_excel_bytes(full_df)
            grouped_excel = dataframe_to_excel_bytes(grouped_df)

            st.success("Файлы успешно обработаны.")

            st.subheader("Полный обработанный файл")
            st.write(f"Строк: {len(full_df)}")
            st.dataframe(full_df, use_container_width=True)

            st.download_button(
                label="Скачать полный файл",
                data=full_excel,
                file_name=full_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.subheader("Сгруппированный файл с зонами, товарами и водителями")
            st.write(f"Строк: {len(grouped_df)}")
            st.dataframe(grouped_df, use_container_width=True)

            st.download_button(
                label="Скачать сгруппированный файл",
                data=grouped_excel,
                file_name=grouped_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if "ФИО водителя" in grouped_df.columns:
                no_driver_count = grouped_df["ФИО водителя"].isna().sum()

                if no_driver_count > 0:
                    st.warning(
                        f"Для {no_driver_count} заявок не найден водитель. "
                        "Проверьте номера заявок в файле водителей."
                    )

        except Exception as error:
            st.error("Произошла ошибка при обработке файлов.")
            st.exception(error)