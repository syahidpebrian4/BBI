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
    jml_artikel = st.number_input("Total jumlah artikel yang ditransaksikan", min_value=0, step=1, format="%d")
    
    # Input angka tanpa format di UI Streamlit agar user mudah mengetik
    harga_jual = st.number_input("Total Harga Jual Produk", min_value=0, step=1, format="%d")
    biaya_delivery = st.number_input("Total Biaya Delivery", min_value=0, step=1, format="%d")
    total_transfer = st.number_input("Total transfer", min_value=0, step=1, format="%d")
    
    lokasi_transaksi = st.text_input("Lokasi transaksi")
    rencana_transaksi = st.date_input("Rencana transaksi (estimasi)")
    
    submitted = st.form_submit_button("Generate Draft Memo")

if submitted:
    template_path = "Draft_Memo_Template.xlsx"
    
    if not os.path.exists(template_path):
        st.error(f"File {template_path} tidak ditemukan!")
    else:
        try:
            # Menggunakan keep_vba=True untuk menjaga elemen grafis (gambar/tanda tangan)
            wb = openpyxl.load_workbook(template_path, keep_vba=True)
            ws = wb.active
            
            # --- MENGISI DATA ---
            ws['D6'] = str(no_memo)
            ws['D8'] = str(no_po)
            ws['F18'] = int(jml_artikel)
            ws['F19'] = int(harga_jual)
            ws['F20'] = int(biaya_delivery)
            ws['F21'] = int(total_transfer)
            ws['F22'] = str(lokasi_transaksi)
            ws['F23'] = rencana_transaksi.strftime("%d %b %Y")
            
            # --- FORMAT RUPIAH ---
            # 'Rp #,##0' akan otomatis menambah 'Rp' dan titik ribuan di Excel
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
                label="Download Memo Excel",
                data=buffer,
                file_name=f"Memo_{no_memo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
