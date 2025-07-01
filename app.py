import streamlit as st
from openpyxl import load_workbook
import xlwt
import io
from datetime import datetime

st.title("📈 Ekspor Data IPH - Komoditas Andil Terbesar")

uploaded_file = st.file_uploader("Upload File Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        wb = load_workbook(uploaded_file, data_only=True)
        sheet_names = wb.sheetnames

        # Pilih sheet yang ingin digunakan
        selected_sheet = st.selectbox("Pilih Sheet", sheet_names)
        ws = wb[selected_sheet]

        data = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and any(row):  # hanya baris dengan data
                selected = [row[i] if i < len(row) else None for i in [0, 1, 2, 3, 4]]
                data.append(selected)

        if data:
            # Tampilkan preview
            st.subheader("📋 Preview Data")
            st.dataframe(data, use_container_width=True)

            # Buat file .xls
            output_excel = io.BytesIO()
            book = xlwt.Workbook()
            sheet = book.add_sheet("IPH_Andil_Provinsi")

            headers = ["Provinsi", "Perubahan IPH", "Komoditas Andil Terbesar", "Nama Komoditas", "CV"]
            for col, val in enumerate(headers):
                sheet.write(0, col, val)

            for row_idx, row in enumerate(data, start=1):
                for col_idx, val in enumerate(row):
                    sheet.write(row_idx, col_idx, val)

            book.save(output_excel)
            output_excel.seek(0)

            # Unduh file
            st.success("✅ Data berhasil diekspor!")
            st.download_button(
                label="📥 Unduh Hasil (.xls)",
                data=output_excel,
                file_name=f"IPH_Andil_{datetime.today().strftime('%Y%m%d')}.xls",
                mime="application/vnd.ms-excel"
            )
        else:
            st.warning("❗ Tidak ada data ditemukan di sheet tersebut.")

    except Exception as e:
        st.error(f"❌ Gagal membaca file: {e}")
