import streamlit as st
import openpyxl
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
    # Cek apakah file template ada
    template_path = "Draft_Memo_Template.xlsx" # Pastikan nama file sesuai
    if not os.path.exists(template_path):
        st.error(f"File {template_path} tidak ditemukan di folder!")
    else:
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        
        # Mapping Data ke Sel
        ws['D6'] = no_memo
        ws['D8'] = no_po
        ws['F18'] = jml_artikel
        ws['F19'] = harga_jual
        ws['F20'] = biaya_delivery
        ws['F21'] = total_transfer
        ws['F22'] = lokasi_transaksi
        ws['F23'] = rencana_transaksi.strftime("%d %b %Y")
        
        # Format Angka (Tanpa pemisah ribuan)
        # --- FORMAT ANGKA ---
            # Format: 'Rp #,##0' akan menampilkan "Rp 210.000.000" secara otomatis di Excel
            format_rupiah = 'Rp #,##0'
            
            ws['F19'].number_format = format_rupiah
            ws['F20'].number_format = format_rupiah
            ws['F21'].number_format = format_rupiah
            
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
