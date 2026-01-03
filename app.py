import os
from io import BytesIO
import streamlit as st
import pandas as pd

st.set_page_config(page_title="CMR / Спецификация / Инвойс", layout="wide")
st.title("Онлайн-генератор документов\nCMR, Спецификация и Инвойс")

def list_excel_files():
    files = []
    for f in os.listdir("."):
        if os.path.isfile(f) and f.lower().endswith((".xlsx", ".xlsm", ".xlsb")):
            files.append(f)
    return sorted(files)

def read_excel_file(source, name_hint: str | None = None):
    # source: filepath (str) OR Streamlit UploadedFile
    if isinstance(source, str):
        ext = os.path.splitext(source)[1].lower()
        if ext == ".xlsb":
            return pd.ExcelFile(source, engine="pyxlsb")
        return pd.ExcelFile(source, engine="openpyxl")

    data = source.read()
    bio = BytesIO(data)
    name = (getattr(source, "name", "") or (name_hint or "")).lower()
    if name.endswith(".xlsb"):
        return pd.ExcelFile(bio, engine="pyxlsb")
    return pd.ExcelFile(bio, engine="openpyxl")

with st.expander("1) Источник данных (Excel)", expanded=True):
    repo_files = list_excel_files()

    col1, col2 = st.columns([2, 1])
    with col1:
        st.caption("Excel файлы в репозитории (корень):")
        st.code("\n".join(repo_files) if repo_files else "Не найдено", language="text")

    with col2:
        uploaded = st.file_uploader("Загрузить Excel (xlsx/xlsm/xlsb)", type=["xlsx", "xlsm", "xlsb"])
        chosen = st.selectbox("Или выбрать файл из репозитория", options=repo_files) if repo_files else None

excel_source = uploaded if uploaded is not None else chosen
if not excel_source:
    st.error("Не найден Excel-файл. Загрузите файл или добавьте его в репозиторий.")
    st.stop()

try:
    xl = read_excel_file(excel_source, name_hint=chosen if isinstance(excel_source, str) else None)
except Exception as e:
    st.error(f"Не удалось открыть Excel. Ошибка: {e}")
    st.stop()

st.success("Excel открыт.")
st.write("Листы:", xl.sheet_names)

st.header("Предпросмотр (проверка)")
sheet = st.selectbox("Выбери лист для просмотра", xl.sheet_names)
df = xl.parse(sheet)
st.dataframe(df.head(50), use_container_width=True)
