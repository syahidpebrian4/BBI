import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO
import os
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIG ---
# 1. Konfigurasi
st.set_page_config(page_title="Input Draft Memo", layout="centered")

# 2. Logo Resmi (Satu baris saja, tidak akan double, otomatis di atas)
st.logo("logo.png")

# 3. Judul
st.title("Input Data Memo Transaksi")

# --- INISIALISASI SESSION STATE ---
if 'memo_data' not in st.session_state:
    st.session_state.memo_data = None
if 'excel_buffer' not in st.session_state:
    st.session_state.excel_buffer = None
if 'file_name' not in st.session_state:
    st.session_state.file_name = None

# --- FUNGSI LOGIKA NOMOR MEMO ---
def generate_memo_data(sheet, lokasi_transaksi):
    mapping_lokasi = {
        "Lotte Grosir Pasar Rebo": "01",
        "Lotte Grosir Kelapa Gading": "03",
        "Lotte Grosir Ciputat": "06"
    }
    kode_lokasi = mapping_lokasi.get(lokasi_transaksi, "00")
    all_rows = sheet.get_all_values()
    last_no = 0
    if len(all_rows) > 1:
        for row in reversed(all_rows[1:]):
            if row[0] and row[0].isdigit():
                last_no = int(row[0])
                break
    new_no = last_no + 1
    new_no_str = f"{new_no:04d}"
    bulan_romawi = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
    bulan = bulan_romawi[datetime.now().month]
    tahun = datetime.now().year
    no_memo = f"{new_no_str}/ST{kode_lokasi}/CDHO/{bulan}/{tahun}"
    return new_no_str, no_memo

# --- FUNGSI GOOGLE SHEETS ---
def save_to_sheets(data):
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open("Draft Memo BBI").sheet1
    new_no, no_memo = generate_memo_data(sheet, data[7])
    data[0] = new_no
    data[1] = no_memo
    sheet.append_row(data)
    return new_no, no_memo

# --- FORM INPUT ---
with st.form("memo_form"):
    no_po = st.text_input("No. PO")
    jml_artikel = st.number_input("Total jumlah artikel", min_value=0, step=1)
    harga_jual = st.number_input("Total Harga Jual Produk", min_value=0, step=1)
    biaya_delivery = st.number_input("Total Biaya Delivery", min_value=0, step=1)
    total_transfer = st.number_input("Total transfer", min_value=0, step=1)
    lokasi_transaksi = st.selectbox("Lokasi transaksi", ["Lotte Grosir Pasar Rebo", "Lotte Grosir Kelapa Gading", "Lotte Grosir Ciputat"])
    rencana_transaksi = st.date_input("Rencana transaksi (estimasi)")
    submitted = st.form_submit_button("Generate Draft Memo")

if submitted:
    data_to_save = ["", "", no_po, jml_artikel, harga_jual, biaya_delivery, total_transfer, lokasi_transaksi, str(rencana_transaksi)]
    try:
        new_no, no_memo = save_to_sheets(data_to_save)
        
        template_path = "Draft_Memo_Template.xlsx"
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        
        ws['D6'] = str(no_memo)
        ws['D8'] = str(no_po)
        ws['F18'] = int(jml_artikel)
        ws['G19'] = int(harga_jual)
        ws['G20'] = int(biaya_delivery)
        ws['G21'] = int(total_transfer)
        ws['F22'] = str(lokasi_transaksi)
        ws['F23'] = str(rencana_transaksi) 
        
        for cell in ['G19', 'G20', 'G21']:
            ws[cell].number_format = '#,##0'
            ws[cell].alignment = Alignment(horizontal='left')
            
        output = BytesIO()
        wb.save(output)
        
        # Simpan ke Session State
        st.session_state.memo_data = no_memo
        st.session_state.excel_buffer = output.getvalue()
        st.session_state.file_name = f"Draft Memo_{new_no}.xlsx"
        
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

# --- TAMPILAN HASIL (Selalu muncul jika ada data di session state) ---
if st.session_state.memo_data:
    st.success(f"Berhasil! Nomor Memo: {st.session_state.memo_data}")
    st.download_button(
        label="Download Excel",
        data=st.session_state.excel_buffer,
        file_name=st.session_state.file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
