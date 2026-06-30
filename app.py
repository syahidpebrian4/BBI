import streamlit as st
import openpyxl
from io import BytesIO
import os

st.set_page_config(page_title="Input Memo", layout="centered")
st.title("Input Data Memo Transaksi")

with st.form("memo_form"):
    no_memo = st.text_input("No. Memo")
    no_po = st.text_input("No. PO")
    jml_artikel = st.number_input("Total jumlah artikel", min_value=0, step=1)
    harga_jual = st.number_input("Total Harga Jual Produk", min_value=0, step=1)
    biaya_delivery = st.number_input("Total Biaya Delivery", min_value=0, step=1)
    total_transfer = st.number_input("Total transfer", min_value=0, step=1)
    lokasi_transaksi = st.text_input("Lokasi transaksi")
    rencana_transaksi = st.date_input("Rencana transaksi (estimasi)")
    submitted = st.form_submit_button("Generate Memo")

if submitted:
    template_file = "Draft_Memo_Template.xlsx"
    
    if not os.path.exists(template_file):
        st.error("File template tidak ditemukan!")
    else:
        # Menjaga struktur template agar tidak rusak
        wb = openpyxl.load_workbook(template_file, keep_vba=True)
        ws = wb.active
        
        # Mengisi data sel
        ws['D6'] = no_memo
        ws['D8'] = no_po
        ws['F18'] = jml_artikel
        ws['F19'] = harga_jual
        ws['F20'] = biaya_delivery
        ws['F21'] = total_transfer
        ws['F22'] = lokasi_transaksi
        ws['F23'] = rencana_transaksi.strftime("%d %b %Y")
        
        # Simpan ke buffer
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        st.success("Memo berhasil dibuat!")
        st.download_button(
            label="Download Memo",
            data=buffer,
            file_name=f"Memo_{no_memo}.xlsx",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12"
        )
