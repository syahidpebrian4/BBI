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
        
        # Simpan data
        sheet.append_row(data)
        
        # Ambil kembali baris terakhir untuk mendapatkan No Memo (yang di-generate otomatis oleh rumus)
        # Kita ambil baris terakhir yang baru saja ditambahkan
        all_values = sheet.get_all_values()
        return all_values[-1][1] # Kolom B adalah indeks 1
    except Exception as e:
        st.error(f"Gagal simpan ke Database: {e}")
        return None

# --- FORM INPUT ---
with st.form("memo_form"):
    no_po = st.text_input("No. PO")
    jml_artikel = st.number_input("Total jumlah artikel yang ditransaksikan", min_value=0, step=1)
    harga_jual = st.number_input("Total Harga Jual Produk", min_value=0, step=1)
    biaya_delivery = st.number_input("Total Biaya Delivery", min_value=0, step=1)
    total_transfer = st.number_input("Total transfer", min_value=0, step=1)
    lokasi_transaksi = st.selectbox("Lokasi transaksi", ["Lotte Grosir Pasar Rebo", "Lotte Grosir Kelapa Gading", "Lotte Grosir Ciputat"])
    rencana_transaksi = st.date_input("Rencana transaksi (estimasi)")
    submitted = st.form_submit_button("Generate Draft Memo")

if submitted:
    data_to_save = ["", "", no_po, jml_artikel, harga_jual, biaya_delivery, total_transfer, lokasi_transaksi, str(rencana_transaksi)]
    
    # Simpan dan ambil No Memo hasil rumus
    no_memo_otomatis = save_to_sheets(data_to_save)
    
    if no_memo_otomatis:
        template_path = "Draft_Memo_Template.xlsx" 
        if not os.path.exists(template_path):
            st.error(f"File {template_path} tidak ditemukan!")
        else:
            try:
                wb = openpyxl.load_workbook(template_path, keep_vba=True)
                ws = wb.active
                
                # Mapping ke Excel
                ws['D6'] = no_memo_otomatis # Mengisi No Memo dari hasil Sheet
                ws['D8'] = no_po
                ws['F18'] = jml_artikel
                ws['G19'] = harga_jual
                ws['G20'] = biaya_delivery
                ws['G21'] = total_transfer
                ws['F22'] = lokasi_transaksi
                ws['F23'] = rencana_transaksi.strftime("%d %b %Y")
                
                # Format Angka & Rata Kiri
                for cell in ['G19', 'G20', 'G21']:
                    ws[cell].number_format = '#,##0'
                    ws[cell].alignment = Alignment(horizontal='left')
                    
                # PENTING: Gunakan BytesIO yang benar
                buffer = BytesIO()
                wb.save(buffer)
                buffer.seek(0) # <--- INI KUNCI AGAR TIDAK CORRUPT
                
                st.success(f"Data tersimpan! No Memo: {no_memo_otomatis}")
                st.download_button(
                    label="Download Excel",
                    data=buffer,
                    file_name=f"Memo_{no_memo_otomatis}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan pada proses Excel: {e}")
