import streamlit as st
import openpyxl
from io import BytesIO
import os

# Konfigurasi Halaman
st.set_page_config(page_title="Input Memo", layout="centered")

st.title("Input Data Memo Transaksi")

# Form Input
with st.form("memo_form"):
    no_memo = st.text_input("No. Memo")
    no_po = st.text_input("No. PO")
    jml_artikel = st.number_input("Total jumlah artikel", min_value=0, step=1, format="%d")
    harga_jual = st.number_input("Total Harga Jual Produk", min_value=0, step=1, format="%d")
    biaya_delivery = st.number_input("Total Biaya Delivery", min_value=0, step=1, format="%d")
    total_transfer = st.number_input("Total transfer", min_value=0, step=1, format="%d")
    lokasi_transaksi = st.text_input("Lokasi transaksi")
    rencana_transaksi = st.date_input("Rencana transaksi (estimasi)")
    
    submitted = st.form_submit_button("Generate Memo")

if submitted:
    template_file = "Draft_Memo_Template.xlsx"
    
    if not os.path.exists(template_file):
        st.error(f"File {template_file} tidak ditemukan di repositori GitHub Anda!")
    else:
        try:
            # Load workbook
            # keep_vba=True membantu mempertahankan elemen grafis/macro jika ada
            wb = openpyxl.load_workbook(template_file, keep_vba=True)
            ws = wb.active
            
            # Mengisi data ke sel target
            ws['D6'] = str(no_memo)
            ws['D8'] = str(no_po)
            ws['F18'] = int(jml_artikel)
            ws['F19'] = int(harga_jual)
            ws['F20'] = int(biaya_delivery)
            ws['F21'] = int(total_transfer)
            ws['F22'] = str(lokasi_transaksi)
            ws['F23'] = rencana_transaksi.strftime("%d %b %Y")
            
            # Memastikan format angka tanpa titik/pemisah ribuan
            for cell in ['F19', 'F20', 'F21']:
                ws[cell].number_format = '0'
            
            # Simpan ke memori
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            
            st.success("Memo berhasil dibuat!")
            
            # Tombol Download
            st.download_button(
                label="Download Memo Excel",
                data=buffer,
                file_name=f"Memo_{no_memo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file: {e}")
