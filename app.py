import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIG ---
st.set_page_config(page_title="Draft Memo & Surat Generator", layout="wide")

# --- URL RAW LOGO DARI GITHUB (PASTI MUNCUL & TIDAK KAN ERROR PATH) ---
LOGO_URL = "https://raw.githubusercontent.com/syahidpebrian4/BBI/main/lsi_logo.png"

# --- INISIALISASI SESSION STATE ---
if 'selected_category' not in st.session_state:
    st.session_state.selected_category = None
if 'memo_data' not in st.session_state: 
    st.session_state.memo_data = None
if 'excel_buffer' not in st.session_state: 
    st.session_state.excel_buffer = None
if 'file_name' not in st.session_state: 
    st.session_state.file_name = None

# --- CUSTOM CSS (Fixed Header Logo & Layout) ---
st.markdown(f"""
    <style>
    /* Styling Header Bawaan Streamlit */
    header[data-testid="stHeader"] {{
        background-color: transparent !important;
        z-index: 100 !important;
    }}

    /* Fixed Header Bar Logo (Tetap di atas saat scroll) */
    .fixed-header-container {{
        position: fixed;
        top: 0px;
        left: 0px;
        width: 100%;
        height: 70px;
        background-color: #ffffff;
        border-bottom: 2px solid #fee2e2;
        z-index: 999;
        display: flex;
        align-items: center;
        padding-left: 3rem;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.06);
    }}

    .fixed-header-container img {{
        height: 42px;
        width: auto;
        object-fit: contain;
    }}

    /* Padding Atas Agar Konten Tidak Tertutup Navbar Logo */
    .block-container {{
        padding-top: 6rem !important;
    }}

    /* Styling Title Landing Page */
    .main-title {{
        text-align: center;
        font-size: 32px;
        font-weight: 800;
        color: #b91c1c;
        margin-bottom: 4px;
        letter-spacing: -0.5px;
    }}
    
    .sub-title {{
        text-align: center;
        font-size: 15px;
        color: #64748b;
        margin-bottom: 35px;
    }}

    /* Card Styling */
    .custom-card {{
        background: #ffffff;
        border: 1px solid #fee2e2;
        border-radius: 16px;
        padding: 28px 24px;
        box-shadow: 0 10px 25px -5px rgba(185, 28, 28, 0.08);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
    }}
    
    .custom-card::before {{
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 6px;
        height: 100%;
        background: linear-gradient(180deg, #dc2626 0%, #991b1b 100%);
    }}

    .card-tag {{
        display: inline-block;
        background-color: #fef2f2;
        color: #991b1b;
        font-size: 11px;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 20px;
        margin-bottom: 12px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        border: 1px solid #fecaca;
    }}

    .card-h2 {{
        font-size: 22px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 8px;
    }}

    /* Tombol Utama Streamlit */
    div.stButton > button {{
        background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 10px 20px !important;
        box-shadow: 0 4px 12px rgba(220, 38, 38, 0.25) !important;
        transition: all 0.2s ease !important;
    }}

    div.stButton > button:hover {{
        background: linear-gradient(135deg, #b91c1c 0%, #991b1b 100%) !important;
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(220, 38, 38, 0.35) !important;
    }}
    </style>

    <!-- HTML Navbar Logo Melayang -->
    <div class="fixed-header-container">
        <img src="{LOGO_URL}" alt="LSI Logo" />
    </div>
""", unsafe_allow_html=True)
