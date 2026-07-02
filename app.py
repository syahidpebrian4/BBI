import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Download Memo", layout="centered")
st.title("Download Draft Memo Terakhir")

# --- FUNGSI AMBIL DATA TERAKHIR ---
def get_latest_data():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    
    sheet = client.open("Draft Memo BBI").sheet1
    all_data = sheet.get_all_values()
    
    if len(all_data) > 1:
        # Mengambil baris terakhir
        return all_data[-1] 
    return None

if st.button("Generate & Download Memo Terakhir"):
    data = get_latest_data()
    
    if data:
        # Mapping Data sesuai urutan kolom (A=0, B=1, C=2, dst)
        # B=1 (No Memo), C=2 (No PO), D=3 (Jml), E=4 (Harga), F=5 (Biaya), G=6 (Transfer), H=7 (Lokasi), I=8 (Rencana)
        no_memo = data[1]
        no_po = data[2]
        jml_artikel = int(data[3])
        harga_jual = int(data[4])
        biaya_delivery = int(data[5])
        total_transfer = int(data[6])
        lokasi = data[7]
        rencana = data[8]
        
        template_path = "Draft_Memo_Template.xlsx"
        if os.path.exists(template_path):
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            
            # Mapping ke koordinat sel
            ws['D6'] = no_memo
            ws['D8'] = no_po
            ws['F18'] = jml_artikel
            ws['G19'] = harga_jual
            ws['G20'] = biaya_delivery
            ws['G21'] = total_transfer
            ws['F22'] = lokasi
            ws['F23'] = rencana
            
            # Formatting
            for cell in ['G19', 'G20', 'G21']:
                ws[cell].number_format = '#,##0'
                ws[cell].alignment = Alignment(horizontal='left')
                
            buffer = BytesIO()
            wb.save(buffer)
            
            st.success("Berhasil mengambil data terakhir!")
            st.download_button("Download Excel", data=buffer.getvalue(), 
                               file_name=f"Memo_{no_memo}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("Template tidak ditemukan!")
    else:
        st.warning("Data kosong di spreadsheet.")
