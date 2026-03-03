import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
import uuid

# --- KONFIGURASI SPREADSHEET ---
# Pastikan Anda sudah memiliki file JSON Service Account dari Google Cloud
# Ganti 'path/to/your/credentials.json' dengan nama file rahasia Anda
# SPREADSHEET_ID = 'waji diubah id nya' 

def get_gsheet():
    # Setup scope untuk Google Sheets API
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # Gunakan secrets streamlit jika di deploy ke GitHub/Streamlit Cloud
    # Jika lokal, gunakan: credentials = Credentials.from_service_account_file('credentials.json', scopes=scope)
    try:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)
        # Buka spreadsheet berdasarkan ID yang ada di file asli 
        return client.open_by_key('ID_SPREADSHEET_ANDA') 
    except Exception as e:
        st.error(f"Koneksi ke Google Sheets Gagal: {e}")
        return None

# --- FUNGSI HELPER (KONVERSI DARI JS KE PYTHON) ---

def calculate_time_diff(start_time, end_time):
    """Menghitung selisih menit (Konversi dari function calculateTimeDiff di code.gs) [cite: 463]"""
    fmt = '%H:%M'
    t1 = datetime.strptime(start_time, fmt)
    t2 = datetime.strptime(end_time, fmt)
    diff = (t2 - t1).total_seconds() / 60
    return int(diff)

def clean_date_format(raw_tgl):
    """Memperbaiki format tanggal (Konversi dari logika baris 376 di code.gs) """
    if not raw_tgl:
        return ""
    
    # Ganti // menjadi # untuk komentar di Python
    # Hapus kutip jika ada (Penyelesaian error SyntaxError sebelumnya) 
    clean_tgl = str(raw_tgl).replace("'", "").strip() 
    
    if '-' in clean_tgl:
        parts = clean_tgl.split('-')
        # Jika format dd-mm-yyyy (tahun di belakang) [cite: 378]
        if len(parts[2]) == 4:
            return f"{parts[2]}-{parts[1]}-{parts[0]}" # Jadi yyyy-mm-dd
    return clean_tgl

# --- LOGIKA UTAMA (STREAMLIT UI) ---

def main():
    st.set_page_config(page_title="Aplikasi Absensi Sekolah", layout="wide") # [cite: 352]
    
    st.title("Aplikasi Absensi Sekolah MIN 1 Ciamis")
    
    # Inisialisasi Google Sheet
    ss = get_gsheet()
    if not ss: return

    # Sidebar Login (Konversi dari function login di code.gs) [cite: 353]
    with st.sidebar:
        st.header("Login Sistem")
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        btn_login = st.button("Masuk")

    if btn_login:
        # Contoh pengecekan login sederhana dari sheet 'users' [cite: 358]
        users_sheet = ss.worksheet('users')
        users_data = users_sheet.get_all_records()
        
        user_found = next((u for u in users_data if str(u['Username']) == username and str(u['Password']) == password), None)
        
        if user_found:
            st.success(f"Selamat Datang, {username}!")
            st.session_state['role'] = user_found['Role']
            st.session_state['token'] = str(uuid.uuid4()) # Generate token unik [cite: 361]
        else:
            st.error("Username atau password salah") # [cite: 360]

    # MENU UTAMA (Hanya jika sudah login)
    if 'role' in st.session_state:
        menu = st.selectbox("Menu", ["Scan Absensi", "Data Siswa", "Laporan Absensi"])
        
        if menu == "Data Siswa":
            # Menampilkan List Siswa (Konversi dari getSiswaList) [cite: 372]
            siswa_sheet = ss.worksheet('siswa')
            data_siswa = pd.DataFrame(siswa_sheet.get_all_records())
            st.dataframe(data_siswa)

        elif menu == "Scan Absensi":
            st.subheader("Scan QR Code Absensi")
            nisn_input = st.text_input("Masukkan NISN (Simulasi Scan)")
            
            if st.button("Absen Sekarang"):
                # Jalankan logika scanAbsensi dari code.gs [cite: 427]
                # (Logika absen datang/pulang dimasukkan di sini)
                st.info(f"Memproses absensi untuk NISN: {nisn_input}")

if __name__ == "__main__":
    main()
