import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIG ---
st.set_page_config(page_title="Input Draft Memo", layout="centered")
st.title("Input Data Memo Transaksi")

# --- FUNGSI GOOGLE SHEETS ---
def save_to_sheets(data):
    try:
        creds_dict = st.secrets["gcp_service_account"]
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        sheet = client.open("Draft Memo BBI").sheet1
        sheet.append_row(data)
        return True
    except Exception as e:
        st.error(f"Gagal simpan ke Database: {e}")
        return False

# --- FORM INPUT ---
with st.form("memo_form"):
    # No. Memo dihapus karena otomatis di Spreadsheet
    no_po = st.text_input("No. PO")
    jml_artikel = st.number_input("Total jumlah artikel yang ditransaksikan", min_value=0, step=1)
    harga_jual = st.number_input("Total Harga Jual Produk", min_value=0, step=1)
    biaya_delivery = st.number_input("Total Biaya Delivery", min_value=0, step=1)
    total_transfer = st.number_input("Total transfer", min_value=0, step=1)
    
    # Lokasi dirubah menjadi Dropdown
    lokasi_transaksi = st.selectbox("Lokasi transaksi", [
        "Lotte Grosir Pasar Rebo",
        "Lotte Grosir Kelapa Gading",
        "Lotte Grosir Ciputat"
    ])
    
    rencana_transaksi = st.date_input("Rencana transaksi (estimasi)")
    
    submitted = st.form_submit_button("Generate Draft Memo")

if submitted:
    # 1. Simpan ke Google Sheets (Urutan: No, No Memo, No PO, Jml, Harga, Biaya, Transfer, Lokasi, Rencana)
    # Kolom B (No Memo) diisi dengan string kosong atau simbol agar rumus di Sheets yang bekerja
    data_to_save = [
        "",              # Kolom A (No)
        "",              # Kolom B (No Memo - dikosongkan untuk rumus)
        no_po,           # Kolom C
        jml_artikel,     # Kolom D
        harga_jual,      # Kolom E
        biaya_delivery,  # Kolom F
        total_transfer,  # Kolom G
        lokasi_transaksi,# Kolom H
        str(rencana_transaksi) # Kolom I
    ]
    
    success_db = save_to_sheets(data_to_save)
    
    if success_db:
        # 2. Proses Excel Template
        template_path = "Draft_Memo_Template.xlsx" 
        if not os.path.exists(template_path):
            st.error(f"File {template_path} tidak ditemukan!")
        else:
            try:
                wb = openpyxl.load_workbook(template_path, keep_vba=True)
                ws = wb.active
                
                # Mapping ke Excel
                ws['D8'] = no_po # D6 (No Memo) dikosongkan/dihapus
                ws['F18'] = jml_artikel
                ws['G19'] = harga_jual
                ws['G20'] = biaya_delivery
                ws['G21'] = total_transfer
                ws['F22'] = lokasi_transaksi
                ws['F23'] = rencana_transaksi.strftime("%d %b %Y")
                
                # Format Angka & Rata Kiri
                format_number = '#,##0'
                rata_kiri = Alignment(horizontal='left')
                
                for cell in ['G19', 'G20', 'G21']:
                    ws[cell].number_format = format_number
                    ws[cell].alignment = rata_kiri
                    
                buffer = BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                st.success("Data tersimpan ke database & Excel siap!")
                st.download_button(
                    label="Download Excel",
                    data=buffer,
                    file_name=f"Memo_{no_po}.xlsx", # Menggunakan No PO sebagai nama file
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan pada proses Excel: {e}")
