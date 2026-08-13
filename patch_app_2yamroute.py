from pathlib import Path

APP_PATH = Path("app.py")

IMPORT_MARKER = "from processing import ("
NEW_IMPORT = "from yamroute2 import process_yamroute2_file\n"

TAB_OLD = '''assignments_tab, registry_tab, bitrix24_tab, yamroute_tab, comparison_tab = st.tabs([
    "Заявки",
    "Работа с реестром",
    "Работа с Битрикс24",
    "ЯМаршрут",
    "Сравнение отчётов"
])'''

TAB_NEW = '''def render_yamroute2_page():
    st.title("2ЯМаршрут")

    st.write(
        "Загрузите реестр доставок. Сервис сформирует Excel-файл "
        "для ЯМаршрут по новому шаблону с привязкой каждой заявки к складу."
    )

    st.info(
        "Склад определяется по адресу филиала: Сименса, дом 3 → 100; "
        "Выборгское шоссе, дом 503 → 200; остальные адреса → 300. "
        "Вес и количество мест не включают строку «Доставка товара клиенту»."
    )

    source_file = st.file_uploader(
        "Основной файл доставок для 2ЯМаршрут",
        type=["xlsx"],
        key="yamroute2_source_file"
    )

    if source_file is not None:
        st.success(f"Файл загружен: {source_file.name}")

    if st.button(
        "Подготовить файл 2ЯМаршрут",
        key="process_yamroute2_button"
    ):
        if source_file is None:
            st.error("Сначала загрузите основной файл доставок.")
        else:
            try:
                source_file.seek(0)

                (
                    result_df,
                    result_bytes,
                    result_filename
                ) = process_yamroute2_file(
                    main_file=source_file,
                    original_filename=source_file.name
                )

                st.session_state["yamroute2_df"] = result_df
                st.session_state["yamroute2_file_bytes"] = result_bytes
                st.session_state["yamroute2_filename"] = result_filename

                st.success("Файл 2ЯМаршрут успешно подготовлен.")

            except Exception as error:
                st.session_state.pop("yamroute2_df", None)
                st.session_state.pop("yamroute2_file_bytes", None)
                st.session_state.pop("yamroute2_filename", None)
                st.error("Не удалось подготовить файл 2ЯМаршрут.")
                st.exception(error)

    if "yamroute2_df" not in st.session_state:
        return

    result_df = st.session_state["yamroute2_df"]
    result_bytes = st.session_state["yamroute2_file_bytes"]
    result_filename = st.session_state["yamroute2_filename"]

    st.subheader("Результат")
    st.write(f"Заявок: {len(result_df)}")
    show_centered_table(result_df)

    st.download_button(
        label="Скачать файл 2ЯМаршрут",
        data=result_bytes,
        file_name=result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_yamroute2_file"
    )


assignments_tab, registry_tab, bitrix24_tab, yamroute_tab, yamroute2_tab, comparison_tab = st.tabs([
    "Заявки",
    "Работа с реестром",
    "Работа с Битрикс24",
    "ЯМаршрут",
    "2ЯМаршрут",
    "Сравнение отчётов"
])'''

WITH_OLD = '''with yamroute_tab:
    render_yamroute_page()

with comparison_tab:
    render_report_comparison_page()'''

WITH_NEW = '''with yamroute_tab:
    render_yamroute_page()

with yamroute2_tab:
    render_yamroute2_page()

with comparison_tab:
    render_report_comparison_page()'''


def main():
    if not APP_PATH.exists():
        raise FileNotFoundError(
            "app.py не найден. Запустите скрипт из корня проекта."
        )

    text = APP_PATH.read_text(encoding="utf-8")

    if "from yamroute2 import process_yamroute2_file" not in text:
        if IMPORT_MARKER not in text:
            raise RuntimeError(
                "Не найден блок импорта processing в app.py."
            )

        text = text.replace(
            IMPORT_MARKER,
            NEW_IMPORT + "\n" + IMPORT_MARKER,
            1
        )

    if "def render_yamroute2_page():" not in text:
        if TAB_OLD not in text:
            raise RuntimeError(
                "Не найден ожидаемый блок st.tabs в app.py. "
                "Проверьте, что используется последняя версия проекта."
            )

        function_marker = "def render_report_comparison_page():"
        if function_marker not in text:
            raise RuntimeError(
                "Не найдена функция render_report_comparison_page."
            )

        function_code = TAB_NEW.split(
            'assignments_tab, registry_tab, bitrix24_tab, yamroute_tab, yamroute2_tab, comparison_tab = st.tabs(['
        )[0].rstrip()

        text = text.replace(
            function_marker,
            function_code + "\n\n\n" + function_marker,
            1
        )

        text = text.replace(TAB_OLD, TAB_NEW, 1)

        if WITH_OLD not in text:
            raise RuntimeError(
                "Не найден блок with yamroute_tab / comparison_tab."
            )

        text = text.replace(WITH_OLD, WITH_NEW, 1)

    APP_PATH.write_text(text, encoding="utf-8")
    print("Готово: app.py обновлён для вкладки «2ЯМаршрут».")


if __name__ == "__main__":
    main()
