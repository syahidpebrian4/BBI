import streamlit as st
import openpyxl
from openpyxl.styles import Alignment  # Tambahkan baris ini
from io import BytesIO
import os

st.set_page_config(page_title="Input Draft Memo", layout="centered")

st.title("Input Data Memo Transaksi")

# --- FORM INPUT ---
with st.form("memo_form"):
    no_memo = st.text_input("No. Memo")
    no_po = st.text_input("No. PO")
    jml_artikel = st.number_input("Total jumlah artikel yang ditransaksikan", min_value=0, step=1)
    harga_jual = st.number_input("Total Harga Jual Produk (tanpa titik)", min_value=0, step=1)
    biaya_delivery = st.number_input("Total Biaya Delivery (tanpa titik)", min_value=0, step=1)
    total_transfer = st.number_input("Total transfer (tanpa titik)", min_value=0, step=1)
    lokasi_transaksi = st.text_input("Lokasi transaksi")
    rencana_transaksi = st.date_input("Rencana transaksi (estimasi)")
    
    submitted = st.form_submit_button("Generate Draft Memo")

if submitted:
    template_path = "Draft_Memo_Template.xlsx" 
    if not os.path.exists(template_path):
        st.error(f"File {template_path} tidak ditemukan di folder!")
    else:
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        
        # Mapping Data
        ws['D6'] = no_memo
        ws['D8'] = no_po
        ws['F18'] = jml_artikel
        ws['G19'] = harga_jual
        ws['G20'] = biaya_delivery
        ws['G21'] = total_transfer
        ws['F22'] = lokasi_transaksi
        ws['F23'] = rencana_transaksi.strftime("%d %b %Y")
        
        # --- FORMAT ANGKA & RATA KIRI ---
        format_number = '#,##0'
        rata_kiri = Alignment(horizontal='left')
        
        # Daftar sel yang ingin diformat
        cells_to_format = ['G19', 'G20', 'G21']
        
        for cell in cells_to_format:
            ws[cell].number_format = format_number
            ws[cell].alignment = rata_kiri
            
        # Simpan ke memori
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        st.success("Draft siap diunduh!")
        st.download_button(
            label="Download Excel",
            data=buffer,
            file_name=f"Memo_{no_memo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
