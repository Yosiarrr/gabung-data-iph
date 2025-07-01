import streamlit as st
from openpyxl import load_workbook
import xlwt
import io

st.title("📥 Ekspor Data IPH Tanpa Pivot")

uploaded_file = st.file_uploader("Upload File Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        wb = load_workbook(uploaded_file, data_only=True)
        sheet_names = wb.sheetnames

        # Pilih sheet
        selected_sheet = st.selectbox("Pilih Sheet", sheet_names)
        ws = wb[selected_sheet]

        # Ambil data semua baris dan kolom
        data = list(ws.iter_rows(values_only=True))
        header = data[0]
        rows = data[1:]

        # Preview data di Streamlit
        st.subheader("📊 Preview Data (tanpa pivot)")
        st.dataframe(data)

        # Ekspor ke .xls
        output = io.BytesIO()
        book = xlwt.Workbook()
        sheet = book.add_sheet("Data_IPH")

        for col_index, col_name in enumerate(header):
            sheet.write(0, col_index, col_name)

        for row_index, row in enumerate(rows, start=1):
            for col_index, value in enumerate(row):
                sheet.write(row_index, col_index, value)

        book.save(output)
        output.seek(0)

        st.success("✅ Data berhasil diekspor tanpa pivot!")
        st.download_button(
            label="📥 Unduh Hasil (.xls)",
            data=output,
            file_name="IPH_Tanpa_Pivot.xls",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"❌ Gagal membaca file: {e}")
