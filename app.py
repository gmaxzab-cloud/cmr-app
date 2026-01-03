import streamlit as st
import pandas as pd
from io import BytesIO


@st.cache_data
def load_workbook(path: str) -> dict:
    """Load all sheets from the provided Excel file into a dict of DataFrames.

    Caching is enabled because the workbook only needs to be read once per session.
    """
    excel_file = pd.ExcelFile(path)
    return {name: excel_file.parse(name) for name in excel_file.sheet_names}


def to_excel_bytes(dfs: dict) -> bytes:
    """Serialize a dict of DataFrames into an Excel file stored in memory.

    Args:
        dfs: A dictionary where keys are sheet names and values are DataFrames.

    Returns:
        A bytes object representing the Excel file.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output.read()


def prepare_lists(data: dict) -> dict:
    """Extract lists used for dropdowns from the workbook.

    Args:
        data: The loaded workbook as a dict of DataFrames.

    Returns:
        A dictionary containing lists and lookup dictionaries for senders, receivers, etc.
    """
    lists = {}
    # Отправители
    senders_df = data.get('Отправитель')
    if senders_df is not None and 'Наименование компании' in senders_df.columns:
        lists['senders'] = senders_df['Наименование компании'].dropna().unique().tolist()
        # Map sender to address
        lists['sender_address'] = senders_df.set_index('Наименование компании')['Адрес компании'].to_dict()
    else:
        lists['senders'] = []
        lists['sender_address'] = {}
    # Получатели и перевозчики
    receivers_df = data.get('Получатель и перевозчик')
    if receivers_df is not None and 'Наименование' in receivers_df.columns:
        lists['receivers'] = receivers_df['Наименование'].dropna().unique().tolist()
        lists['receiver_address'] = receivers_df.set_index('Наименование')['Адрес'].to_dict() if 'Адрес' in receivers_df.columns else {}
        lists['receiver_ogrn'] = receivers_df.set_index('Наименование')['ОГРН.1'].to_dict() if 'ОГРН.1' in receivers_df.columns else {}
    else:
        lists['receivers'] = []
        lists['receiver_address'] = {}
        lists['receiver_ogrn'] = {}
    # Условия поставки from "доп данные"
    extra_df = data.get('доп данные')
    if extra_df is not None and 'Условия поставки' in extra_df.columns:
        lists['incoterms'] = extra_df['Условия поставки'].dropna().unique().tolist()
    else:
        lists['incoterms'] = []
    return lists


def main():
    st.set_page_config(page_title="Генератор документов CMR/Specification/Invoice")
    st.title("Онлайн‑генератор документов CMR, Спецификация и Инвойс")

    # Load workbook
    # The file path could be adjusted; by default, we load the sample provided with this app.
    workbook_path = 'СМАПП спец — копия.xlsx'
    try:
        data = load_workbook(workbook_path)
    except FileNotFoundError:
        st.error(f"Не удалось найти файл: {workbook_path}. Загрузите корректный файл Excel.")
        return

    lists = prepare_lists(data)

    # Containers to hold generated outputs for each document
    cmr_rows = []
    spec_rows = []
    invoice_rows = []

    tabs = st.tabs(["CMR", "Спецификация", "Инвойс"])

    # CMR tab
    with tabs[0]:
        st.header("Форма CMR")
        with st.form(key="cmr_form"):
            sender = st.selectbox("Отправитель", options=lists['senders'], help="Выберите отправителя из списка")
            sender_addr = lists['sender_address'].get(sender, '')
            st.text_input("Адрес отправителя", value=sender_addr, key="cmr_sender_address")
            receiver = st.selectbox("Получатель", options=lists['receivers'], help="Выберите получателя из списка")
            receiver_addr = lists['receiver_address'].get(receiver, '')
            st.text_input("Адрес получателя", value=receiver_addr, key="cmr_receiver_address")
            carrier = st.selectbox("Перевозчик", options=lists['receivers'], help="Выберите перевозчика из списка")
            carrier_addr = lists['receiver_address'].get(carrier, '')
            st.text_input("Адрес перевозчика", value=carrier_addr, key="cmr_carrier_address")
            incoterm = st.selectbox("Условия поставки (Инкотермс)", options=lists['incoterms'])
            contract_no = st.text_input("Номер контракта")
            invoice_no = st.text_input("Номер инвойса")
            cmr_no = st.text_input("Номер CMR")
            # Additional fields can be added as needed
            if st.form_submit_button("Добавить в список CMR"):
                cmr_rows.append({
                    'Отправитель': sender,
                    'Адрес отправителя': sender_addr,
                    'Получатель': receiver,
                    'Адрес получателя': receiver_addr,
                    'Перевозчик': carrier,
                    'Адрес перевозчика': carrier_addr,
                    'Условия поставки': incoterm,
                    'Номер контракта': contract_no,
                    'Номер инвойса': invoice_no,
                    'Номер CMR': cmr_no,
                })
                st.success("Запись CMR добавлена.")

        if cmr_rows:
            st.subheader("Добавленные записи CMR")
            cmr_df = pd.DataFrame(cmr_rows)
            st.dataframe(cmr_df, use_container_width=True)
            excel_bytes = to_excel_bytes({'CMR': cmr_df})
            st.download_button("Скачать CMR (Excel)", data=excel_bytes, file_name="cmr.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Specification tab
    with tabs[1]:
        st.header("Форма спецификации")
        with st.form(key="spec_form"):
            seller = st.selectbox("Продавец", options=lists['senders'], help="Выберите продавца из списка")
            seller_addr = lists['sender_address'].get(seller, '')
            st.text_input("Адрес продавца", value=seller_addr, key="spec_seller_address")
            consignee = st.selectbox("Получатель", options=lists['receivers'])
            consignee_addr = lists['receiver_address'].get(consignee, '')
            st.text_input("Адрес получателя", value=consignee_addr, key="spec_receiver_address")
            contract_no_spec = st.text_input("Номер контракта")
            spec_no = st.text_input("Номер спецификации")
            goods_desc = st.text_area("Описание товаров (каждая строка — отдельная позиция)")
            currency = st.selectbox("Валюта", options=["USD", "EUR", "RUB", "CNY"])
            if st.form_submit_button("Добавить в список спецификаций"):
                spec_rows.append({
                    'Продавец': seller,
                    'Адрес продавца': seller_addr,
                    'Получатель': consignee,
                    'Адрес получателя': consignee_addr,
                    'Номер контракта': contract_no_spec,
                    'Номер спецификации': spec_no,
                    'Описание товаров': goods_desc,
                    'Валюта': currency,
                })
                st.success("Запись спецификации добавлена.")

        if spec_rows:
            st.subheader("Добавленные спецификации")
            spec_df = pd.DataFrame(spec_rows)
            st.dataframe(spec_df, use_container_width=True)
            excel_bytes_spec = to_excel_bytes({'Specification': spec_df})
            st.download_button("Скачать спецификации (Excel)", data=excel_bytes_spec, file_name="specifications.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Invoice tab
    with tabs[2]:
        st.header("Форма инвойса")
        with st.form(key="invoice_form"):
            exporter = st.selectbox("Экспортер", options=lists['senders'])
            exporter_addr = lists['sender_address'].get(exporter, '')
            st.text_input("Адрес экспортера", value=exporter_addr, key="invoice_exporter_address")
            buyer = st.selectbox("Покупатель", options=lists['receivers'])
            buyer_addr = lists['receiver_address'].get(buyer, '')
            st.text_input("Адрес покупателя", value=buyer_addr, key="invoice_buyer_address")
            invoice_number = st.text_input("Номер инвойса")
            invoice_date = st.date_input("Дата инвойса")
            contract_no_inv = st.text_input("Номер контракта")
            incoterm_inv = st.selectbox("Условия поставки (Инкотермс)", options=lists['incoterms'])
            goods_list_inv = st.text_area("Описание товаров (каждая строка — отдельная позиция)")
            if st.form_submit_button("Добавить в список инвойсов"):
                invoice_rows.append({
                    'Экспортер': exporter,
                    'Адрес экспортера': exporter_addr,
                    'Покупатель': buyer,
                    'Адрес покупателя': buyer_addr,
                    'Номер инвойса': invoice_number,
                    'Дата инвойса': invoice_date,
                    'Номер контракта': contract_no_inv,
                    'Условия поставки': incoterm_inv,
                    'Описание товаров': goods_list_inv,
                })
                st.success("Запись инвойса добавлена.")

        if invoice_rows:
            st.subheader("Добавленные инвойсы")
            invoice_df = pd.DataFrame(invoice_rows)
            st.dataframe(invoice_df, use_container_width=True)
            excel_bytes_inv = to_excel_bytes({'Invoices': invoice_df})
            st.download_button("Скачать инвойсы (Excel)", data=excel_bytes_inv, file_name="invoices.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    main()