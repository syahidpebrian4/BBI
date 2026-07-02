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
if 'file_name' not in st.session_state: st.session_state.file_name = None

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
    
    # Input harga di Streamlit menggunakan titik agar user mudah mengetik ratusan juta
    harga_jual_str = st.text_input("Total Harga Jual Produk (contoh: 200.000.000)")
    biaya_delivery_str = st.text_input("Total Biaya Delivery (contoh: 100.000)")
    
    # Konversi string dengan titik menjadi integer untuk perhitungan
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
        
        # Data ke Google Sheets
        row_data = [new_no, str(tanggal_input), no_memo, no_po, jml_artikel, harga_jual, biaya_delivery, total_transfer, lokasi_transaksi, str(rencana_transaksi)]
        sheet.append_row(row_data)
        
        # --- PROSES EXCEL ---
        wb = openpyxl.load_workbook("Draft_Memo_Template.xlsx")
        ws = wb.active

        bulan_indo = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        str_tanggal = f"Jakarta, {tanggal_input.day} {bulan_indo[tanggal_input.month]} {tanggal_input.year}"
        ws['I6'] = str_tanggal
        
        ws['D6'] = no_memo
        ws['D8'] = no_po
        ws['F18'] = jml_artikel
        ws['G19'] = harga_jual
        ws['G20'] = biaya_delivery
        ws['G21'] = total_transfer
        ws['F22'] = lokasi_transaksi
        ws['F23'] = str(rencana_transaksi)

        # Format ribuan koma di Excel (#,##0)
        for cell in ['G19', 'G20', 'G21']:
            ws[cell].number_format = '#,##0'
            ws[cell].alignment = Alignment(horizontal='left')
            
        output = BytesIO()
        wb.save(output)
        
        # Simpan ke session state
        st.session_state.memo_data = no_memo
        st.session_state.excel_buffer = output.getvalue()
        st.session_state.file_name = f"Draft Memo_{new_no}.xlsx"
        st.rerun()
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

# --- TAMPILAN HASIL ---
if st.session_state.memo_data:
    st.success(f"Berhasil! Nomor Memo: {st.session_state.memo_data}")
    st.download_button(
        label="Download Excel", 
        data=st.session_state.excel_buffer, 
        file_name=st.session_state.file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
