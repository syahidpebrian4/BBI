import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO
import os
import gspread
import time
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIG ---
st.set_page_config(page_title="Input Draft Memo", layout="centered")
st.title("Input Data Memo Transaksi")

# --- FUNGSI LOGIKA NOMOR MEMO (DIUPDATE) ---
def generate_memo_data(sheet, lokasi_transaksi):
    # Mapping Lokasi
    mapping_lokasi = {
        "Lotte Grosir Pasar Rebo": "01",
        "Lotte Grosir Kelapa Gading": "03",
        "Lotte Grosir Ciputat": "06"
    }
    kode_lokasi = mapping_lokasi.get(lokasi_transaksi, "00")
    
    # Ambil semua data untuk menghitung nomor urut
    all_rows = sheet.get_all_values()
    if len(all_rows) > 1:
        # Mengambil nilai kolom A (indeks 0) dari baris terakhir
        last_val = all_rows[-1][0]
        # Jika nilai adalah "0001", kita ambil angkanya saja
        last_no = int(last_val) if last_val.isdigit() else 0
    else:
        last_no = 0
    
    new_no = last_no + 1
    # Format nomor urut menjadi 0001, 0002, dst
    new_no_str = f"{new_no:04d}"
    
    # Format Bulan Romawi
    bulan_romawi = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
    bulan = bulan_romawi[datetime.now().month]
    tahun = datetime.now().year
    
    # Format No Memo: 0001/ST01/CDHO/VII/2026
    # Sesuai permintaan: [No urut]/ST[Lokasi]/CDHO/[Bulan]/[Tahun]
    no_memo = f"{new_no_str}/ST{kode_lokasi}/CDHO/{bulan}/{tahun}"
    
    return new_no_str, no_memo

# --- FUNGSI GOOGLE SHEETS ---
def save_to_sheets(data):
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open("Draft Memo BBI").sheet1
    
    # Generate nomor
    new_no, no_memo = generate_memo_data(sheet, data[7]) # index 7 adalah lokasi
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
    
    new_no, no_memo = save_to_sheets(data_to_save)
    
    # Proses Excel
    template_path = "Draft_Memo_Template.xlsx"
    wb = openpyxl.load_workbook(template_path, keep_vba=True)
    ws = wb.active
    
    ws['D6'] = no_memo
    ws['D8'] = no_po
    ws['F18'] = jml_artikel
    ws['G19'] = harga_jual
    ws['G20'] = biaya_delivery
    ws['G21'] = total_transfer
    ws['F22'] = lokasi_transaksi
    ws['F23'] = rencana_transaksi.strftime("%d %b %Y")
    
    for cell in ['G19', 'G20', 'G21']:
        ws[cell].number_format = '#,##0'
        ws[cell].alignment = Alignment(horizontal='left')
        
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    st.success(f"Berhasil! Memo: {no_memo}")
    st.download_button("Download Excel", buffer, f"Memo_{no_memo.replace('/', '-')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
