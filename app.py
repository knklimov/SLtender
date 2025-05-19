
import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill
import io

# Целевой зелёный цвет в формате ARGB, предоставленный пользователем
TARGET_GREEN = "FF5AFC4C"

def process_excel(file):
    wb = openpyxl.load_workbook(file)
    modifications = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                fill = cell.fill
                if fill and fill.fill_type == "solid" and fill.fgColor and fill.fgColor.rgb:
                    color = fill.fgColor.rgb
                    if color.upper() != TARGET_GREEN:
                        try:
                            original_value = cell.value
                            if isinstance(original_value, (int, float)):
                                new_value = original_value - 3
                                cell.value = new_value
                                modifications.append((sheet_name, cell.coordinate, original_value, new_value, color))
                        except Exception:
                            pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, modifications

st.title("Изменение значений в Excel по цвету ячеек")

uploaded_file = st.file_uploader("Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    processed_file, log = process_excel(uploaded_file)

    st.download_button(
        label="📥 Скачать обработанный файл",
        data=processed_file,
        file_name="modified_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if log:
        st.write("Изменённые ячейки:")
        st.dataframe(
            {"Лист": [x[0] for x in log],
             "Ячейка": [x[1] for x in log],
             "Старое значение": [x[2] for x in log],
             "Новое значение": [x[3] for x in log],
             "Цвет": [x[4] for x in log]}
        )
    else:
        st.success("Файл обработан: изменений не найдено.")
