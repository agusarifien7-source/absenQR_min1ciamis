# app.py
import os
import uuid
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from flask import Flask, render_template, request, jsonify, session, send_file
from functools import wraps
import json
import pandas as pd
from io import BytesIO

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Untuk session management

# ====================================
# KONFIGURASI GOOGLE SHEETS
# ====================================
SPREADSHEET_ID = '14_ngdeGuv24xiJ5vmctLjsvb27LjH8hMZnUsFyD6J5c'

# Setup Google Sheets credentials
# Pastikan file credentials.json ada di folder yang sama
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

def get_spreadsheet():
    """Mendapatkan spreadsheet berdasarkan ID"""
    return client.open_by_key(SPREADSHEET_ID)

# ====================================
# DECORATOR UNTUK CEK SESSION (AUTH)
# ====================================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            return jsonify({'success': False, 'message': 'Silakan login terlebih dahulu'}), 401
        return f(*args, **kwargs)
    return decorated_function

def role_required(required_role):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user' not in session:
                return jsonify({'success': False, 'message': 'Silakan login terlebih dahulu'}), 401
            
            user_role = session['user'].get('role')
            if user_role != required_role and user_role != 'admin':
                return jsonify({'success': False, 'message': 'Akses Ditolak: Anda tidak memiliki izin.'}), 403
            return f(*args, **kwargs)
        return decorated_function
    return decorator

# ====================================
# ROUTES UTAMA
# ====================================
@app.route('/')
def index():
    """Main entry point - render halaman index"""
    return render_template('index.html')

@app.route('/api/login', methods=['POST'])
def login():
    """Endpoint login"""
    try:
        data = request.json
        username = data.get('username', '')
        password = data.get('password', '')
        nisn = data.get('nisn', '')
        
        ss = get_spreadsheet()
        
        # Login Siswa (Pakai NISN)
        if nisn:
            siswa_sheet = ss.worksheet('siswa')
            siswa_data = siswa_sheet.get_all_values()
            
            user_found = None
            for i, row in enumerate(siswa_data[1:], start=2):  # Lewati header
                if str(row[1]) == str(nisn):
                    user_found = {
                        'role': 'siswa',
                        'identifier': row[1],
                        'nama': row[0],
                        'kelas': row[8] if len(row) > 8 else ''
                    }
                    break
            
            if not user_found:
                return jsonify({'success': False, 'message': 'NISN tidak ditemukan'})
        
        # Login Admin & Guru
        else:
            users_sheet = ss.worksheet('users')
            users_data = users_sheet.get_all_values()
            
            user_found = None
            for i, row in enumerate(users_data[1:], start=2):
                if row[0] == username and row[1] == password:
                    user_found = {
                        'role': row[2],
                        'identifier': row[0],
                        'nama': row[0],
                        'kelas': row[3] if len(row) > 3 else ''
                    }
                    break
            
            if not user_found:
                return jsonify({'success': False, 'message': 'Username atau password salah'})
        
        # Simpan user di session (Flask session)
        session['user'] = {
            'role': user_found['role'],
            'identifier': user_found['identifier'],
            'nama': user_found['nama'],
            'kelas': user_found['kelas']
        }
        
        return jsonify({
            'success': True,
            'role': user_found['role'],
            'username': user_found['identifier'],
            'nama': user_found['nama'],
            'kelas': user_found['kelas'],
            'nisn': user_found['identifier'] if user_found['role'] == 'siswa' else None
        })
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Login Error: {str(e)}'})

@app.route('/api/logout', methods=['POST'])
def logout():
    """Logout user"""
    session.pop('user', None)
    return jsonify({'success': True, 'message': 'Logout berhasil'})

# ====================================
# SISWA - CRUD OPERATIONS
# ====================================
@app.route('/api/siswa', methods=['GET'])
@login_required
def get_siswa_list():
    """Mendapatkan daftar semua siswa"""
    try:
        ss = get_spreadsheet()
        sheet = ss.worksheet('siswa')
        data = sheet.get_all_values()
        
        siswa_list = []
        for row in data[1:]:  # Lewati header
            if row[0]:  # Jika nama tidak kosong
                # Format tanggal lahir
                tgl_lahir = row[3] if len(row) > 3 else ''
                if tgl_lahir:
                    try:
                        # Coba parse tanggal
                        if '-' in tgl_lahir:
                            parts = tgl_lahir.replace("'", "").split('-')
                            if len(parts) == 3:
                                if len(parts[2]) == 4:  # dd-mm-yyyy
                                    tgl_lahir = f"{parts[2]}-{parts[1]}-{parts[0]}"
                    except:
                        pass
                
                siswa_list.append({
                    'nama': row[0],
                    'nisn': row[1],
                    'jenisKelamin': row[2] if len(row) > 2 else '',
                    'tanggalLahir': tgl_lahir,
                    'agama': row[4] if len(row) > 4 else '',
                    'namaAyah': row[5] if len(row) > 5 else '',
                    'namaIbu': row[6] if len(row) > 6 else '',
                    'noHp': row[7] if len(row) > 7 else '',
                    'kelas': row[8] if len(row) > 8 else '',
                    'alamat': row[9] if len(row) > 9 else ''
                })
        
        return jsonify({'success': True, 'data': siswa_list})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/siswa/<nisn>', methods=['GET'])
@login_required
def get_siswa_by_nisn(nisn):
    """Mendapatkan data siswa berdasarkan NISN"""
    try:
        ss = get_spreadsheet()
        sheet = ss.worksheet('siswa')
        data = sheet.get_all_values()
        
        for row in data[1:]:
            if row[1] == nisn:
                # Format tanggal lahir
                tgl_lahir = row[3] if len(row) > 3 else ''
                if tgl_lahir:
                    try:
                        if '-' in tgl_lahir:
                            parts = tgl_lahir.replace("'", "").split('-')
                            if len(parts) == 3:
                                if len(parts[2]) == 4:  # dd-mm-yyyy
                                    tgl_lahir = f"{parts[2]}-{parts[1]}-{parts[0]}"
                    except:
                        pass
                
                return jsonify({
                    'success': True,
                    'data': {
                        'nama': row[0],
                        'nisn': row[1],
                        'jenisKelamin': row[2] if len(row) > 2 else '',
                        'tanggalLahir': tgl_lahir,
                        'agama': row[4] if len(row) > 4 else '',
                        'namaAyah': row[5] if len(row) > 5 else '',
                        'namaIbu': row[6] if len(row) > 6 else '',
                        'noHp': row[7] if len(row) > 7 else '',
                        'kelas': row[8] if len(row) > 8 else '',
                        'alamat': row[9] if len(row) > 9 else ''
                    }
                })
        
        return jsonify({'success': False, 'message': 'Siswa tidak ditemukan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/siswa', methods=['POST'])
@role_required('admin')
def add_siswa():
    """Menambahkan data siswa baru"""
    try:
        data = request.json
        siswa_data = data.get('siswaData', {})
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('siswa')
        
        # Cek NISN duplikat
        all_data = sheet.get_all_values()
        for row in all_data[1:]:
            if row[1] == siswa_data.get('nisn'):
                return jsonify({'success': False, 'message': 'NISN sudah terdaftar'})
        
        # Format tanggal
        tgl_simpan = siswa_data.get('tanggalLahir', '')
        if tgl_simpan and '-' in tgl_simpan:
            parts = tgl_simpan.split('-')
            if len(parts) == 3:
                tgl_simpan = f"'{parts[2]}-{parts[1]}-{parts[0]}"
        
        # Tambah data
        sheet.append_row([
            siswa_data.get('nama', ''),
            siswa_data.get('nisn', ''),
            siswa_data.get('jenisKelamin', ''),
            tgl_simpan,
            siswa_data.get('agama', ''),
            siswa_data.get('namaAyah', ''),
            siswa_data.get('namaIbu', ''),
            siswa_data.get('noHp', ''),
            siswa_data.get('kelas', ''),
            siswa_data.get('alamat', '')
        ])
        
        return jsonify({'success': True, 'message': 'Siswa berhasil ditambahkan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'GAGAL: {str(e)}'})

@app.route('/api/siswa/<old_nisn>', methods=['PUT'])
@role_required('admin')
def update_siswa(old_nisn):
    """Update data siswa"""
    try:
        data = request.json
        siswa_data = data.get('siswaData', {})
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('siswa')
        all_data = sheet.get_all_values()
        
        # Format tanggal
        tgl_simpan = siswa_data.get('tanggalLahir', '')
        if tgl_simpan and '-' in tgl_simpan:
            parts = tgl_simpan.split('-')
            if len(parts) == 3 and len(parts[0]) == 4:  # yyyy-mm-dd
                tgl_simpan = f"'{parts[2]}-{parts[1]}-{parts[0]}"
        
        for i, row in enumerate(all_data[1:], start=2):
            if row[1] == old_nisn:
                # Update data
                sheet.update(f'A{i}:J{i}', [[
                    siswa_data.get('nama', ''),
                    siswa_data.get('nisn', ''),
                    siswa_data.get('jenisKelamin', ''),
                    tgl_simpan,
                    siswa_data.get('agama', ''),
                    siswa_data.get('namaAyah', ''),
                    siswa_data.get('namaIbu', ''),
                    siswa_data.get('noHp', ''),
                    siswa_data.get('kelas', ''),
                    siswa_data.get('alamat', '')
                ]])
                return jsonify({'success': True, 'message': 'Siswa berhasil diupdate'})
        
        return jsonify({'success': False, 'message': 'Siswa tidak ditemukan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/siswa/<nisn>', methods=['DELETE'])
@role_required('admin')
def delete_siswa(nisn):
    """Hapus data siswa"""
    try:
        ss = get_spreadsheet()
        sheet = ss.worksheet('siswa')
        all_data = sheet.get_all_values()
        
        for i, row in enumerate(all_data[1:], start=2):
            if str(row[1]) == str(nisn):
                sheet.delete_rows(i)
                return jsonify({'success': True, 'message': 'Data siswa berhasil dihapus'})
        
        return jsonify({'success': False, 'message': 'Data siswa tidak ditemukan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# ====================================
# GURU - CRUD OPERATIONS
# ====================================
@app.route('/api/guru', methods=['GET'])
@role_required('admin')
def get_guru_list():
    """Mendapatkan daftar semua guru"""
    try:
        ss = get_spreadsheet()
        sheet = ss.worksheet('users')
        data = sheet.get_all_values()
        
        guru_list = []
        for row in data[1:]:  # Lewati header
            if len(row) > 2 and row[2] == 'guru':
                guru_list.append({
                    'username': row[0],
                    'password': row[1],
                    'role': row[2],
                    'kelas': row[3] if len(row) > 3 else ''
                })
        
        return jsonify({'success': True, 'data': guru_list})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/guru', methods=['POST'])
@role_required('admin')
def add_guru():
    """Menambahkan guru baru"""
    try:
        data = request.json
        username = data.get('username')
        password = data.get('password')
        kelas = data.get('kelas', '')
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('users')
        
        # Cek username duplikat
        all_data = sheet.get_all_values()
        for row in all_data[1:]:
            if row[0] == username:
                return jsonify({'success': False, 'message': 'Username sudah terdaftar'})
        
        # Tambah guru
        sheet.append_row([username, password, 'guru', kelas])
        return jsonify({'success': True, 'message': 'Guru berhasil ditambahkan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Akses Ditolak: {str(e)}'})

@app.route('/api/guru/<old_username>', methods=['PUT'])
@role_required('admin')
def update_guru(old_username):
    """Update data guru"""
    try:
        data = request.json
        new_username = data.get('newUsername')
        password = data.get('password')
        kelas = data.get('kelas', '')
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('users')
        all_data = sheet.get_all_values()
        
        for i, row in enumerate(all_data[1:], start=2):
            if row[0] == old_username and row[2] == 'guru':
                sheet.update(f'A{i}:D{i}', [[new_username, password, 'guru', kelas]])
                return jsonify({'success': True, 'message': 'Guru berhasil diupdate'})
        
        return jsonify({'success': False, 'message': 'Guru tidak ditemukan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Akses Ditolak: {str(e)}'})

@app.route('/api/guru/<username>', methods=['DELETE'])
@role_required('admin')
def delete_guru(username):
    """Hapus data guru"""
    try:
        ss = get_spreadsheet()
        sheet = ss.worksheet('users')
        all_data = sheet.get_all_values()
        
        for i, row in enumerate(all_data[1:], start=2):
            if row[0] == username and row[2] == 'guru':
                sheet.delete_rows(i)
                return jsonify({'success': True, 'message': 'Guru berhasil dihapus'})
        
        return jsonify({'success': False, 'message': 'Guru tidak ditemukan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Akses Ditolak: {str(e)}'})

# ====================================
# FUNGSI HELPER
# ====================================
def calculate_time_diff(start_time, end_time):
    """Menghitung selisih menit antara dua waktu"""
    try:
        start_parts = start_time.split(':')
        end_parts = end_time.split(':')
        
        start_minutes = int(start_parts[0]) * 60 + int(start_parts[1])
        end_minutes = int(end_parts[0]) * 60 + int(end_parts[1])
        
        return end_minutes - start_minutes
    except:
        return 0

def get_app_config():
    """Mendapatkan konfigurasi aplikasi"""
    try:
        ss = get_spreadsheet()
        config = {
            'jam_masuk_mulai': '06:00',
            'jam_masuk_akhir': '07:15',
            'jam_pulang_mulai': '15:00',
            'jam_pulang_akhir': '17:00'
        }
        
        try:
            sheet = ss.worksheet('konfigurasi')
            data = sheet.get_all_values()
            for row in data[1:]:
                if len(row) >= 2 and row[0] in config:
                    config[row[0]] = row[1]
        except:
            pass
            
        return {'success': True, 'data': config}
    except Exception as e:
        return {'success': False, 'message': str(e)}

# ====================================
# ABSENSI OPERATIONS
# ====================================
@app.route('/api/absensi/scan', methods=['POST'])
@login_required
def scan_absensi():
    """Scan QR Code untuk absensi"""
    try:
        data = request.json
        nisn = data.get('nisn')
        scanner_role = session['user']['role']
        scanner_kelas = session['user'].get('kelas', '')
        
        today = datetime.now().strftime('%Y-%m-%d')
        now_time = datetime.now().strftime('%H:%M')
        
        # Ambil konfigurasi
        config_result = get_app_config()
        config = config_result['data'] if config_result['success'] else {
            'jam_masuk_akhir': '07:15',
            'jam_pulang_mulai': '15:00',
            'jam_pulang_akhir': '17:00'
        }
        
        ss = get_spreadsheet()
        
        # Cek hari libur
        try:
            libur_sheet = ss.worksheet('hari_libur')
            libur_data = libur_sheet.get_all_values()
            for row in libur_data[1:]:
                if len(row) >= 1 and row[0]:
                    try:
                        tgl_libur = datetime.strptime(row[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                        if tgl_libur == today:
                            return jsonify({
                                'success': False, 
                                'message': f"Absensi DITUTUP. Hari ini libur: {row[1] if len(row) > 1 else ''}"
                            })
                    except:
                        pass
        except:
            pass
        
        absensi_sheet = ss.worksheet('absensi')
        siswa_sheet = ss.worksheet('siswa')
        
        # Validasi NISN
        scanned_nisn = str(nisn).strip()
        if not scanned_nisn or scanned_nisn == "undefined":
            return jsonify({'success': False, 'message': 'QR Code tidak valid atau kosong'})
        
        # Cari data siswa
        siswa_data = siswa_sheet.get_all_values()
        siswa = None
        for row in siswa_data[1:]:
            if len(row) > 1 and str(row[1]).strip() == scanned_nisn:
                siswa = {
                    'nama': row[0],
                    'nisn': row[1],
                    'kelas': row[8] if len(row) > 8 else ''
                }
                break
        
        if not siswa:
            return jsonify({'success': False, 'message': 'NISN tidak terdaftar di database'})
        
        # Validasi kelas guru
        if scanner_role == 'guru':
            kelas_siswa = str(siswa['kelas']).strip().upper()
            kelas_guru = str(scanner_kelas).strip().upper()
            if kelas_guru and kelas_siswa != kelas_guru:
                return jsonify({
                    'success': False,
                    'message': f"Ditolak! Siswa ini kelas {siswa['kelas']}. Anda hanya bisa scan kelas {scanner_kelas}"
                })
        
        # Proses absensi
        absensi_data = absensi_sheet.get_all_values()
        
        # Cari data absensi hari ini
        for i, row in enumerate(absensi_data[1:], start=2):
            if len(row) >= 2:
                try:
                    row_date = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
                except:
                    try:
                        row_date = datetime.strptime(row[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                    except:
                        continue
                
                row_nisn = str(row[1]).strip()
                
                # SKENARIO ABSEN PULANG
                if row_date == today and row_nisn == scanned_nisn:
                    # Cek sudah checkout
                    if len(row) > 5 and row[5]:
                        return jsonify({'success': False, 'message': 'Siswa sudah melakukan absen pulang hari ini'})
                    
                    # Cek batas akhir pulang
                    if now_time > config['jam_pulang_akhir']:
                        return jsonify({
                            'success': False,
                            'message': f"Gagal! Batas waktu pulang ({config['jam_pulang_akhir']}) sudah lewat"
                        })
                    
                    # Cek jeda waktu
                    jam_datang = row[4] if len(row) > 4 else ''
                    if jam_datang:
                        minutes_diff = calculate_time_diff(jam_datang[:5], now_time)
                        if minutes_diff < 10:
                            return jsonify({'success': False, 'message': 'Terlalu Cepat! Tunggu sebentar lagi'})
                    
                    # Update jam pulang
                    ket_saat_ini = row[6] if len(row) > 6 else ''
                    ket_baru = ket_saat_ini
                    pesan_pulang = 'Absen Pulang Berhasil'
                    
                    if now_time < config['jam_pulang_mulai']:
                        ket_baru = ket_saat_ini + " & Pulang Cepat" if ket_saat_ini else "Pulang Cepat"
                        pesan_pulang = 'Absen Pulang (Pulang Cepat)'
                    
                    jam_pulang = datetime.now().strftime('%H:%M:%S')
                    
                    # Update sheet
                    absensi_sheet.update(f'F{i}', jam_pulang)
                    absensi_sheet.update(f'G{i}', ket_baru)
                    
                    return jsonify({
                        'success': True,
                        'message': pesan_pulang,
                        'type': 'pulang',
                        'jamPulang': jam_pulang,
                        'nama': siswa['nama'],
                        'kelas': siswa['kelas'],
                        'status': 'Hadir'
                    })
        
        # SKENARIO ABSEN DATANG
        if now_time > config['jam_pulang_akhir']:
            return jsonify({'success': False, 'message': 'Absensi Ditutup! Sudah melewati jam operasional'})
        
        # Logika keterlambatan
        keterangan_waktu = 'Tepat Waktu'
        status_kehadiran = 'Hadir'
        
        if now_time > config['jam_masuk_akhir']:
            late_minutes = calculate_time_diff(config['jam_masuk_akhir'], now_time)
            keterangan_waktu = f"Terlambat ({late_minutes} m)"
        
        jam_datang = datetime.now().strftime('%H:%M:%S')
        
        # Insert data baru
        absensi_sheet.append_row([
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            f"'{scanned_nisn}",
            siswa['nama'],
            siswa['kelas'],
            jam_datang,
            '',
            keterangan_waktu,
            status_kehadiran
        ])
        
        response_message = 'Absen Masuk Berhasil'
        if 'Terlambat' in keterangan_waktu:
            response_message = f"Absen Masuk ({keterangan_waktu})"
        
        return jsonify({
            'success': True,
            'message': response_message,
            'type': 'datang',
            'jamDatang': jam_datang,
            'nama': siswa['nama'],
            'kelas': siswa['kelas'],
            'status': status_kehadiran
        })
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error Server: {str(e)}'})

@app.route('/api/absensi/today', methods=['GET'])
@login_required
def get_absensi_today():
    """Mendapatkan data absensi hari ini untuk siswa"""
    try:
        nisn = request.args.get('nisn')
        if not nisn:
            return jsonify({'success': False, 'message': 'NISN diperlukan'})
        
        ss = get_spreadsheet()
        today_str = datetime.now().strftime('%Y-%m-%d')
        
        # Cek libur
        is_libur = False
        keterangan_libur = ""
        
        try:
            libur_sheet = ss.worksheet('hari_libur')
            libur_data = libur_sheet.get_all_values()
            for row in libur_data[1:]:
                if len(row) >= 1 and row[0]:
                    try:
                        tgl = datetime.strptime(row[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                        if tgl == today_str:
                            is_libur = True
                            keterangan_libur = row[1] if len(row) > 1 else ""
                            break
                    except:
                        pass
        except:
            pass
        
        # Cari data absensi
        absensi_sheet = ss.worksheet('absensi')
        absensi_data = absensi_sheet.get_all_values()
        search_nisn = str(nisn).strip()
        
        absensi_data_result = None
        
        for row in absensi_data[1:]:
            if len(row) >= 2:
                try:
                    row_date = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
                except:
                    try:
                        row_date = datetime.strptime(row[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                    except:
                        continue
                
                row_nisn = str(row[1]).strip()
                
                if row_date == today_str and row_nisn == search_nisn:
                    jam_datang = row[4] if len(row) > 4 else ''
                    jam_pulang = row[5] if len(row) > 5 and row[5] else ''
                    
                    absensi_data_result = {
                        'tanggal': today_str,
                        'jamDatang': jam_datang,
                        'jamPulang': jam_pulang,
                        'status': row[6] if len(row) > 6 else ''
                    }
                    break
        
        return jsonify({
            'success': True,
            'data': absensi_data_result,
            'isLibur': is_libur,
            'keteranganLibur': keterangan_libur
        })
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/absensi', methods=['GET'])
@login_required
def get_absensi_list():
    """Mendapatkan daftar absensi dengan filter"""
    try:
        filter_nama = request.args.get('nama', '')
        filter_kelas = request.args.get('kelas', '')
        filter_tgl_mulai = request.args.get('tanggalMulai', '')
        filter_tgl_akhir = request.args.get('tanggalAkhir', '')
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('absensi')
        data = sheet.get_all_values()
        
        absensi_list = []
        
        for row in data[1:]:
            if len(row) >= 1 and row[0]:
                try:
                    raw_date = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
                    tanggal_str = raw_date.strftime('%Y-%m-%d')
                    date_for_filter = tanggal_str
                except:
                    try:
                        raw_date = datetime.strptime(row[0], '%Y-%m-%d')
                        tanggal_str = raw_date.strftime('%Y-%m-%d')
                        date_for_filter = tanggal_str
                    except:
                        continue
                
                jam_datang = row[4] if len(row) > 4 else ''
                jam_pulang = row[5] if len(row) > 5 and row[5] else '-'
                
                item = {
                    'tanggal': tanggal_str,
                    'nisn': row[1] if len(row) > 1 else '',
                    'nama': row[2] if len(row) > 2 else '',
                    'kelas': row[3] if len(row) > 3 else '',
                    'jamDatang': jam_datang,
                    'jamPulang': jam_pulang,
                    'keterangan': row[6] if len(row) > 6 else '',
                    'status': row[7] if len(row) > 7 else ''
                }
                
                # Filter
                match = True
                if filter_tgl_mulai and tanggal_str < filter_tgl_mulai:
                    match = False
                if filter_tgl_akhir and tanggal_str > filter_tgl_akhir:
                    match = False
                if filter_nama and filter_nama.lower() not in item['nama'].lower():
                    match = False
                if filter_kelas and item['kelas'] != filter_kelas:
                    match = False
                
                if match:
                    absensi_list.append(item)
        
        # Urutkan dari terbaru
        absensi_list.sort(key=lambda x: x['tanggal'], reverse=True)
        
        return jsonify({'success': True, 'data': absensi_list})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/kelas', methods=['GET'])
@login_required
def get_kelas_list():
    """Mendapatkan daftar kelas unik"""
    try:
        ss = get_spreadsheet()
        sheet = ss.worksheet('siswa')
        data = sheet.get_all_values()
        
        kelas_set = set()
        for row in data[1:]:
            if len(row) > 8 and row[8]:
                kelas_set.add(row[8])
        
        return jsonify({'success': True, 'data': sorted(list(kelas_set))})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# ====================================
# MONITORING
# ====================================
@app.route('/api/monitoring', methods=['GET'])
@login_required
def get_monitoring():
    """Mendapatkan data monitoring realtime"""
    try:
        filter_kelas = request.args.get('kelas')
        user_role = session['user']['role']
        
        # Jika guru, filter otomatis sesuai kelasnya
        if user_role == 'guru':
            filter_kelas = session['user'].get('kelas')
        
        ss = get_spreadsheet()
        today_str = datetime.now().strftime('%Y-%m-%d')
        
        siswa_sheet = ss.worksheet('siswa')
        absensi_sheet = ss.worksheet('absensi')
        
        data_siswa = siswa_sheet.get_all_values()
        data_absensi = absensi_sheet.get_all_values()
        
        # Mapping data absensi hari ini
        absensi_map = {}
        for row in data_absensi[1:]:
            if len(row) >= 1 and row[0]:
                try:
                    tgl = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
                except:
                    try:
                        tgl = datetime.strptime(row[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                    except:
                        continue
                
                nisn = str(row[1]).strip() if len(row) > 1 else ''
                
                if tgl == today_str and nisn:
                    absensi_map[nisn] = {
                        'jamDatang': row[4] if len(row) > 4 else '',
                        'jamPulang': row[5] if len(row) > 5 else '',
                        'keterangan': row[6] if len(row) > 6 else '',
                        'status': row[7] if len(row) > 7 else ''
                    }
        
        result = []
        for row in data_siswa[1:]:
            if len(row) >= 1 and row[0]:
                nama = row[0]
                nisn = str(row[1]).strip() if len(row) > 1 else ''
                kelas = row[8] if len(row) > 8 else ''
                
                # Filter kelas
                if filter_kelas and kelas != filter_kelas:
                    continue
                
                status_info = absensi_map.get(nisn, {})
                
                # Default value
                jam_datang = status_info.get('jamDatang', '-')
                jam_pulang = status_info.get('jamPulang', '-')
                display_status = status_info.get('status', 'Belum Absen')
                keterangan_waktu = status_info.get('keterangan', '-')
                
                # Format jam
                if jam_datang and jam_datang != '-':
                    if len(jam_datang) > 5:
                        jam_datang = jam_datang[:5]
                
                if jam_pulang and jam_pulang != '-' and len(jam_pulang) > 5:
                    jam_pulang = jam_pulang[:5]
                
                if not keterangan_waktu and display_status == 'Hadir':
                    keterangan_waktu = 'Tepat Waktu'
                
                result.append({
                    'nama': nama,
                    'nisn': nisn,
                    'kelas': kelas,
                    'jamDatang': jam_datang,
                    'jamPulang': jam_pulang,
                    'status': display_status,
                    'keterangan': keterangan_waktu
                })
        
        # Sort by kelas then nama
        result.sort(key=lambda x: (x['kelas'], x['nama']))
        
        return jsonify({'success': True, 'data': result})
        
    except Exception as e:
        return jsonify({'success':
                # Sort by kelas then nama
                result.sort(key=lambda x: (x['kelas'], x['nama']))
                
                return jsonify({'success': True, 'data': result})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/absensi/status', methods=['PUT'])
@login_required
@role_required('guru')
def update_absensi_status():
    """Update status absensi siswa"""
    try:
        data = request.json
        nisn = data.get('nisn')
        nama = data.get('nama')
        kelas = data.get('kelas')
        new_status = data.get('newStatus')
        
        ss = get_spreadsheet()
        absensi_sheet = ss.worksheet('absensi')
        today_str = datetime.now().strftime('%Y-%m-%d')
        
        all_data = absensi_sheet.get_all_values()
        
        found = False
        row_index = None
        
        for i, row in enumerate(all_data[1:], start=2):
            if len(row) >= 1 and row[0]:
                try:
                    tgl = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
                except:
                    try:
                        tgl = datetime.strptime(row[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                    except:
                        continue
                
                row_nisn = str(row[1]).strip() if len(row) > 1 else ''
                
                if tgl == today_str and row_nisn == str(nisn).strip():
                    found = True
                    row_index = i
                    break
        
        if found:
            # Update kolom H (kolom 8) - Status
            absensi_sheet.update(f'H{row_index}', new_status)
        else:
            # Insert data baru
            jam_datang = '-'
            if new_status == 'Hadir':
                jam_datang = datetime.now().strftime('%H:%M:%S')
            
            absensi_sheet.append_row([
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                f"'{nisn}",
                nama,
                kelas,
                jam_datang,
                '',
                '-',
                new_status
            ])
        
        return jsonify({'success': True, 'message': 'Status berhasil diubah'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Gagal: {str(e)}'})

# ====================================
# HARI LIBUR
# ====================================
@app.route('/api/hari-libur', methods=['GET'])
@login_required
def get_hari_libur():
    """Mendapatkan daftar hari libur"""
    try:
        ss = get_spreadsheet()
        sheet = ss.worksheet('hari_libur')
        data = sheet.get_all_values()
        
        list_libur = []
        for row in data[1:]:  # Lewati header
            if len(row) >= 1 and row[0]:
                try:
                    tgl = datetime.strptime(row[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                    list_libur.append({
                        'tanggal': tgl,
                        'keterangan': row[1] if len(row) > 1 else ''
                    })
                except:
                    pass
        
        # Urutkan descending
        list_libur.sort(key=lambda x: x['tanggal'], reverse=True)
        
        return jsonify({'success': True, 'data': list_libur})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/hari-libur', methods=['POST'])
@role_required('admin')
def add_hari_libur():
    """Menambahkan hari libur baru"""
    try:
        data = request.json
        tanggal = data.get('tanggal')
        keterangan = data.get('keterangan')
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('hari_libur')
        
        sheet.append_row([tanggal, keterangan])
        
        return jsonify({'success': True, 'message': 'Hari libur berhasil ditambahkan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/hari-libur/<tanggal>', methods=['PUT'])
@role_required('admin')
def update_hari_libur(tanggal):
    """Update hari libur"""
    try:
        data = request.json
        new_tanggal = data.get('newTanggal')
        new_keterangan = data.get('newKeterangan')
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('hari_libur')
        all_data = sheet.get_all_values()
        
        found = False
        
        for i, row in enumerate(all_data[1:], start=2):
            if len(row) >= 1 and row[0]:
                try:
                    row_date = datetime.strptime(row[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                    if row_date == tanggal:
                        sheet.update(f'A{i}:B{i}', [[new_tanggal, new_keterangan]])
                        found = True
                        break
                except:
                    pass
        
        if found:
            return jsonify({'success': True, 'message': 'Hari libur berhasil diperbarui'})
        else:
            return jsonify({'success': False, 'message': 'Data tanggal tidak ditemukan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/hari-libur/<tanggal>', methods=['DELETE'])
@role_required('admin')
def delete_hari_libur(tanggal):
    """Hapus hari libur"""
    try:
        ss = get_spreadsheet()
        sheet = ss.worksheet('hari_libur')
        all_data = sheet.get_all_values()
        
        for i, row in enumerate(all_data[1:], start=2):
            if len(row) >= 1 and row[0]:
                try:
                    row_date = datetime.strptime(row[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                    if row_date == tanggal:
                        sheet.delete_rows(i)
                        return jsonify({'success': True, 'message': 'Hari libur dihapus'})
                except:
                    pass
        
        return jsonify({'success': False, 'message': 'Data tidak ditemukan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# ====================================
# KONFIGURASI
# ====================================
@app.route('/api/config', methods=['GET'])
@login_required
def get_config():
    """Mendapatkan konfigurasi aplikasi"""
    try:
        result = get_app_config()
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/config', methods=['POST'])
@role_required('admin')
def save_config():
    """Menyimpan konfigurasi aplikasi"""
    try:
        data = request.json
        new_config = data.get('config', {})
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('konfigurasi')
        all_data = sheet.get_all_values()
        
        # Fungsi untuk update berdasarkan key
        def update_row(key, val):
            for i, row in enumerate(all_data[1:], start=2):
                if len(row) >= 1 and row[0] == key:
                    sheet.update(f'B{i}', f"'{val}")
                    return
        
        update_row('jam_masuk_mulai', new_config.get('jam_masuk_mulai', '06:00'))
        update_row('jam_masuk_akhir', new_config.get('jam_masuk_akhir', '07:15'))
        update_row('jam_pulang_mulai', new_config.get('jam_pulang_mulai', '15:00'))
        update_row('jam_pulang_akhir', new_config.get('jam_pulang_akhir', '17:00'))
        
        return jsonify({'success': True, 'message': 'Konfigurasi waktu berhasil disimpan'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# ====================================
# SETUP DATABASE
# ====================================
@app.route('/api/setup', methods=['POST'])
def setup_database():
    """Setup initial database"""
    try:
        ss = get_spreadsheet()
        
        # Setup sheet users
        try:
            users_sheet = ss.worksheet('users')
        except:
            users_sheet = ss.add_worksheet('users', 100, 20)
            users_sheet.append_row(['Username', 'Password', 'Role', 'Kelas'])
            users_sheet.append_row(['admin', 'admin123', 'admin', ''])
            users_sheet.append_row(['guru1', 'guru123', 'guru', 'VI B'])
        
        # Setup sheet siswa
        try:
            siswa_sheet = ss.worksheet('siswa')
        except:
            siswa_sheet = ss.add_worksheet('siswa', 100, 20)
            siswa_sheet.append_row([
                'Nama Lengkap', 'NISN', 'Jenis Kelamin', 'Tanggal Lahir', 'Agama',
                'Nama Ayah', 'Nama Ibu', 'No Handphone', 'Kelas', 'Alamat'
            ])
            siswa_sheet.append_row([
                'Ahmad Rizki', '1234567890', 'Laki-laki', '2008-05-15', 'Islam',
                'Budi Santoso', 'Siti Aminah', '081234567890', 'VI B',
                'Jl. Merdeka No. 10, Bengkulu'
            ])
        
        # Setup sheet absensi
        try:
            absensi_sheet = ss.worksheet('absensi')
        except:
            absensi_sheet = ss.add_worksheet('absensi', 100, 20)
            absensi_sheet.append_row([
                'Tanggal', 'NISN', 'Nama', 'Kelas', 'Jam Datang', 
                'Jam Pulang', 'Keterangan Waktu', 'Status'
            ])
        
        # Setup sheet hari_libur
        try:
            libur_sheet = ss.worksheet('hari_libur')
        except:
            libur_sheet = ss.add_worksheet('hari_libur', 100, 10)
            libur_sheet.append_row(['Tanggal', 'Keterangan'])
        
        # Setup sheet konfigurasi
        try:
            config_sheet = ss.worksheet('konfigurasi')
        except:
            config_sheet = ss.add_worksheet('konfigurasi', 100, 10)
            config_sheet.append_row(['Key', 'Value', 'Keterangan'])
            config_sheet.append_row(['jam_masuk_mulai', '06:00', 'Waktu absen datang dibuka'])
            config_sheet.append_row(['jam_masuk_akhir', '07:15', 'Batas waktu terlambat'])
            config_sheet.append_row(['jam_pulang_mulai', '15:00', 'Waktu absen pulang dibuka'])
            config_sheet.append_row(['jam_pulang_akhir', '17:00', 'Batas akhir absen pulang'])
        
        return jsonify({'success': True, 'message': 'Setup database berhasil'})
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# ====================================
# IMPORT/EXPORT
# ====================================
@app.route('/api/import/siswa', methods=['POST'])
@role_required('admin')
def import_siswa():
    """Import data siswa dari Excel"""
    try:
        data = request.json
        data_array = data.get('data', [])
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('siswa')
        existing_data = sheet.get_all_values()
        
        # Ambil daftar NISN yang sudah ada
        existing_nisn = set()
        for row in existing_data[1:]:
            if len(row) > 1:
                existing_nisn.add(str(row[1]).strip())
        
        rows_to_add = []
        added_count = 0
        skipped_count = 0
        
        for item in data_array:
            nisn = str(item.get('nisn', '')).strip()
            
            if not item.get('nama') or not nisn:
                skipped_count += 1
                continue
            
            if nisn in existing_nisn:
                skipped_count += 1
                continue
            
            # Format tanggal
            tgl_lahir = item.get('tanggalLahir', '')
            
            rows_to_add.append([
                item.get('nama', ''),
                f"'{nisn}",
                item.get('jenisKelamin', ''),
                tgl_lahir,
                item.get('agama', ''),
                item.get('namaAyah', ''),
                item.get('namaIbu', ''),
                f"'{item.get('noHp', '')}",
                item.get('kelas', ''),
                item.get('alamat', '')
            ])
            
            existing_nisn.add(nisn)
            added_count += 1
        
        if rows_to_add:
            for row in rows_to_add:
                sheet.append_row(row)
        
        return jsonify({
            'success': True,
            'added': added_count,
            'skipped': skipped_count,
            'message': f'Import selesai. Berhasil: {added_count}, Duplikat/Gagal: {skipped_count}'
        })
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/import/guru', methods=['POST'])
@role_required('admin')
def import_guru():
    """Import data guru dari Excel"""
    try:
        data = request.json
        data_array = data.get('data', [])
        
        ss = get_spreadsheet()
        sheet = ss.worksheet('users')
        existing_data = sheet.get_all_values()
        
        # Ambil daftar username yang sudah ada
        existing_usernames = set()
        for row in existing_data[1:]:
            if len(row) > 0:
                existing_usernames.add(str(row[0]).strip())
        
        rows_to_add = []
        added_count = 0
        skipped_count = 0
        
        for item in data_array:
            username = str(item.get('username', '')).strip()
            
            if not username or not item.get('password'):
                skipped_count += 1
                continue
            
            if username in existing_usernames:
                skipped_count += 1
                continue
            
            rows_to_add.append([
                f"'{username}",
                f"'{item.get('password', '')}",
                'guru',
                item.get('kelas', '')
            ])
            
            existing_usernames.add(username)
            added_count += 1
        
        if rows_to_add:
            for row in rows_to_add:
                sheet.append_row(row)
        
        return jsonify({
            'success': True,
            'added': added_count,
            'skipped': skipped_count,
            'message': f'Import selesai. Berhasil: {added_count}, Duplikat/Gagal: {skipped_count}'
        })
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/export/<type>', methods=['GET'])
@login_required
def export_data(type):
    """Export data ke Excel"""
    try:
        filter_kelas = request.args.get('kelas')
        filter_tgl_mulai = request.args.get('tanggalMulai')
        filter_tgl_akhir = request.args.get('tanggalAkhir')
        
        ss = get_spreadsheet()
        
        if type == 'laporan_absensi':
            # Export laporan absensi
            sheet = ss.worksheet('absensi')
            data = sheet.get_all_values()
            
            result = []
            no = 1
            
            for row in data[1:]:
                if len(row) >= 1 and row[0]:
                    try:
                        raw_date = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
                        tanggal_str = raw_date.strftime('%d-%m-%Y')
                        date_for_filter = raw_date.strftime('%Y-%m-%d')
                    except:
                        try:
                            raw_date = datetime.strptime(row[0], '%Y-%m-%d')
                            tanggal_str = raw_date.strftime('%d-%m-%Y')
                            date_for_filter = raw_date.strftime('%Y-%m-%d')
                        except:
                            continue
                    
                    row_kelas = row[3] if len(row) > 3 else ''
                    
                    # Filter
                    match = True
                    if filter_tgl_mulai and date_for_filter < filter_tgl_mulai:
                        match = False
                    if filter_tgl_akhir and date_for_filter > filter_tgl_akhir:
                        match = False
                    if filter_kelas and row_kelas != filter_kelas:
                        match = False
                    
                    if match:
                        jam_datang = row[4] if len(row) > 4 else ''
                        jam_pulang = row[5] if len(row) > 5 and row[5] else '-'
                        
                        result.append([
                            no,
                            tanggal_str,
                            f"'{row[1]}" if len(row) > 1 else '',
                            row[2] if len(row) > 2 else '',
                            row_kelas,
                            jam_datang,
                            jam_pulang,
                            row[6] if len(row) > 6 else '',
                            row[7] if len(row) > 7 else ''
                        ])
                        no += 1
            
            # Buat DataFrame
            df = pd.DataFrame(result, columns=[
                'No', 'Tanggal', 'NISN', 'Nama Siswa', 'Kelas', 
                'Jam Datang', 'Jam Pulang', 'Keterangan Waktu', 'Status Kehadiran'
            ])
            
        elif type == 'monitoring':
            # Export monitoring harian
            filter_kelas = filter_kelas or (session['user'].get('kelas') if session['user']['role'] == 'guru' else None)
            
            # Panggil fungsi monitoring
            from flask import Response
            import requests
            
            # Gunakan request internal
            monitoring_result = get_monitoring().json
            
            if not monitoring_result.get('success'):
                return jsonify({'success': False, 'message': 'Gagal mengambil data monitoring'})
            
            data = monitoring_result.get('data', [])
            
            result = []
            for i, item in enumerate(data):
                result.append([
                    i + 1,
                    item.get('nama', ''),
                    f"'{item.get('nisn', '')}",
                    item.get('kelas', ''),
                    item.get('jamDatang', '-'),
                    item.get('jamPulang', '-'),
                    item.get('keterangan', '-'),
                    item.get('status', 'Belum Absen')
                ])
            
            # Buat DataFrame
            df = pd.DataFrame(result, columns=[
                'No', 'Nama Siswa', 'NISN', 'Kelas', 
                'Jam Datang', 'Jam Pulang', 'Keterangan Waktu', 'Status Terkini'
            ])
            
        else:
            return jsonify({'success': False, 'message': 'Tipe export tidak valid'})
        
        # Buat file Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # Styling (sederhana)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Format header
            for col in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=1, column=col)
                cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
                cell.fill = openpyxl.styles.PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
                cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
            
            # Auto adjust column width
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width
        
        output.seek(0)
        
        # Generate filename
        timestamp = datetime.now().strftime('%d-%m-%Y %H%M')
        filename = f"{'Laporan Absensi' if type == 'laporan_absensi' else 'Monitoring Harian'} - {timestamp}.xlsx"
        
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Gagal generate Excel: {str(e)}'})

# ====================================
# CEK SESSION
# ====================================
@app.route('/api/me', methods=['GET'])
@login_required
def get_current_user():
    """Mendapatkan data user yang sedang login"""
    return jsonify({
        'success': True,
        'user': session.get('user')
    })

# ====================================
# MAIN
# ====================================
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
