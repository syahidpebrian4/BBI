import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIG ---
st.set_page_config(page_title="Input Draft Memo", layout="centered")
st.title("Input Data Memo Transaksi")

# --- INISIALISASI SESSION STATE ---
if 'memo_data' not in st.session_state: st.session_state.memo_data = None
if 'excel_buffer' not in st.session_state: st.session_state.excel_buffer = None

# --- FUNGSI PEMBANTU FORMAT RIBUAN ---
def format_input_number(label):
    val = st.text_input(label, value="0")
    # Hapus titik agar bisa jadi integer untuk perhitungan
    clean_val = val.replace(".", "")
    try:
        return int(clean_val)
    except:
        return 0

# --- FUNGSI LOGIKA NOMOR MEMO ---
def generate_memo_data(sheet, lokasi_transaksi):
    mapping_lokasi = {"Lotte Grosir Pasar Rebo": "01", "Lotte Grosir Kelapa Gading": "03", "Lotte Grosir Ciputat": "06"}
    kode_lokasi = mapping_lokasi.get(lokasi_transaksi, "00")
    all_rows = sheet.get_all_values()
    last_no = 0
    if len(all_rows) > 1:
        for row in reversed(all_rows[1:]):
            if row[0] and row[0].isdigit():
                last_no = int(row[0])
                break
    new_no_str = f"{last_no + 1:04d}"
    bulan_romawi = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
    bulan = bulan_romawi[datetime.now().month]
    no_memo = f"{new_no_str}/ST{kode_lokasi}/CDHO/{bulan}/{datetime.now().year}"
    return new_no_str, no_memo

# --- FORM INPUT ---
with st.form("memo_form"):
    tanggal_input = st.date_input("Tanggal")
    no_po = st.text_input("No. PO")
    jml_artikel = st.number_input("Total jumlah artikel", min_value=0, step=1)
    
    # Input harga dengan trik pemisah ribuan
    harga_jual_str = st.text_input("Total Harga Jual Produk", value="0")
    biaya_delivery_str = st.text_input("Total Biaya Delivery", value="0")
    
    # Konversi ke int
    harga_jual = int(harga_jual_str.replace(".", "")) if harga_jual_str.replace(".", "").isdigit() else 0
    biaya_delivery = int(biaya_delivery_str.replace(".", "")) if biaya_delivery_str.replace(".", "").isdigit() else 0
    
    total_transfer = harga_jual + biaya_delivery
    st.write(f"**Total Transfer Terhitung:** {total_transfer:,}".replace(",", "."))
    
    lokasi_transaksi = st.selectbox("Lokasi transaksi", ["Lotte Grosir Pasar Rebo", "Lotte Grosir Kelapa Gading", "Lotte Grosir Ciputat"])
    rencana_transaksi = st.date_input("Rencana transaksi")
    submitted = st.form_submit_button("Generate Draft Memo")

if submitted:
    # --- PROSES SIMPAN ---
    try:
        # Koneksi Google Sheets
        creds_dict = st.secrets["gcp_service_account"]
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet = client.open("Draft Memo BBI").sheet1
        
        new_no, no_memo = generate_memo_data(sheet, lokasi_transaksi)
        
        # Data untuk kolom: No, Tanggal, No Memo, No PO, Total Artikel, Harga, Delivery, Total, Lokasi, Rencana
        row_data = [new_no, str(tanggal_input), no_memo, no_po, jml_artikel, harga_jual, biaya_delivery, total_transfer, lokasi_transaksi, str(rencana_transaksi)]
        sheet.append_row(row_data)
        
        # --- PROSES EXCEL ---
        wb = openpyxl.load_workbook("Draft_Memo_Template.xlsx")
        ws = wb.active
        # Sesuaikan cell ini dengan template baru kamu setelah ada penambahan kolom
        ws['D6'] = no_memo
        ws['D8'] = no_po
        ws['F18'] = jml_artikel
        ws['G19'] = harga_jual
        ws['G20'] = biaya_delivery
        ws['G21'] = total_transfer
        ws['F22'] = lokasi_transaksi
        ws['F23'] = str(rencana_transaksi)

        ws['G19'] = int(harga_jual)
        ws['G20'] = int(biaya_delivery)
        ws['G21'] = int(total_transfer)
        
        # Format angka di Excel
        for cell in ['G19', 'G20', 'G21']:
            ws[cell].number_format = '#.##0'
            
        output = BytesIO()
        wb.save(output)
        st.session_state.memo_data = no_memo
        st.session_state.excel_buffer = output.getvalue()
        st.session_state.file_name = f"Draft Memo_{new_no}.xlsx"
        st.rerun()
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

if st.session_state.memo_data:
    st.success(f"Berhasil! Nomor Memo: {st.session_state.memo_data}")
    st.download_button("Download Excel", st.session_state.excel_buffer, st.session_state.file_name)
