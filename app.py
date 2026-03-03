import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime

# --- KONFIGURASI GOOGLE SHEETS ---
# Di Streamlit Cloud, masukkan isi file JSON key Anda ke 'Settings > Secrets'
def init_connection():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    # Mengambil kredensial dari secrets Streamlit
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], 
        scopes=scope
    )
    client = gspread.authorize(creds)
    return client

# --- FUNGSI LOGIKA (TERJEMAHAN DARI CODE.GS) ---

def process_attendance(nisn):
    try:
        client = init_connection()
        # GANTI ID_INI dengan ID Spreadsheet Anda
        ss = client.open_by_key("ID_SPREADSHEET_ANDA") 
        sheet = ss.worksheet("absensi")
        
        now = datetime.now()
        tgl_skrg = now.strftime("%Y-%m-%d")
        jam_skrg = now.strftime("%H:%M:%S")

        # Contoh logika pembersihan string (Perbaikan error Anda sebelumnya)
        # Di Python gunakan # untuk komentar, bukan //
        clean_nisn = str(nisn).replace("'", "").strip() 
        
        # Simpan ke spreadsheet
        sheet.append_row([clean_nisn, tgl_skrg, jam_skrg, "Hadir"])
        return True
    except Exception as e:
        st.error(f"Error: {e}")
        return False

# --- TAMPILAN INTERFACE (STREAMLIT) ---

st.title("Sistem Absensi MIN 1 Ciamis")
st.subheader("Scan atau Input NISN Siswa")

with st.form("absensi_form"):
    nisn_input = st.text_input("Nomor Induk Siswa Nasional (NISN)")
    submit = st.form_submit_button("Kirim Absensi")

if submit:
    if nisn_input:
        success = process_attendance(nisn_input)
        if success:
            st.success(f"Absensi berhasil dicatat untuk NISN: {nisn_input}")
    else:
        st.warning("Mohon masukkan NISN terlebih dahulu.")
