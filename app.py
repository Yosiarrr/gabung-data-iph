import streamlit as st
from openpyxl import load_workbook
import xlwt
import io
import zipfile

st.title("📊 Gabung Data Excel IPH - Versi Full Tanpa Pivot")

# Pilih tahun & bulan
tahun = st.selectbox("Pilih Tahun", [2023, 2024, 2025], index=2, key="tahun")
bulan = st.selectbox(
    "Pilih Bulan",
    ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
     "Juli", "Agustus", "September", "Oktober", "November", "Desember"],
    index=0,
    key="bulan"
)
map_bulan = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12",
}
bulan_num = map_bulan[bulan]

uploaded_files = st.file_uploader("Upload file Excel (.xlsx)", type="xlsx", accept_multiple_files=True)

if st.button("🔄 Proses & Unduh ZIP"):

    semua_kab, semua_prov = [], []
    header_kab, header_prov = [], []

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in uploaded_files:
            wb = load_workbook(f, data_only=True)
            sheetnames = wb.sheetnames

            sheet_kab = wb["360 KabKota"] if "360 KabKota" in sheetnames else None
            sheet_prov = wb["Provinsi"] if "Provinsi" in sheetnames else None

            # Gabung KabKota
            if sheet_kab:
                rows = list(sheet_kab.iter_rows(values_only=True))
                if not header_kab:
                    header_kab = [cell for cell in rows[0]]
                for row in rows[1:]:
                    if row[0] and str(row[0]).startswith("18"):  # kode_kab
                        semua_kab.append(list(row))

            # Gabung Provinsi
            if sheet_prov:
                rows = list(sheet_prov.iter_rows(values_only=True))
                if not header_prov:
                    header_prov = [cell for cell in rows[0]]
                for row in rows[1:]:
                    if row[0]:
                        semua_prov.append(list(row))

        # Tulis hasil kabupaten
        if semua_kab:
            bk = xlwt.Workbook()
            sk = bk.add_sheet("Gabungan_Kabupaten")

            for i, col in enumerate(header_kab):
                sk.write(0, i, col)
            for i, row in enumerate(semua_kab, 1):
                for j, val in enumerate(row):
                    sk.write(i, j, val)

            buf = io.BytesIO()
            bk.save(buf)
            buf.seek(0)
            zf.writestr(f"kabupaten_{bulan_num}_{tahun}.xls", buf.read())

        # Tulis hasil provinsi
        if semua_prov:
            bp = xlwt.Workbook()
            sp = bp.add_sheet("Gabungan_Provinsi")

            for i, col in enumerate(header_prov):
                sp.write(0, i, col)
            for i, row in enumerate(semua_prov, 1):
                for j, val in enumerate(row):
                    sp.write(i, j, val)

            buf = io.BytesIO()
            bp.save(buf)
            buf.seek(0)
            zf.writestr(f"provinsi_{bulan_num}_{tahun}.xls", buf.read())

    zip_buffer.seek(0)
    st.download_button(
        "📥 Unduh Gabungan IPH (Tanpa Pivot)",
        data=zip_buffer,
        file_name=f"gabungan_IPH_{bulan_num}_{tahun}.zip",
        mime="application/zip"
    )
