
// KONFIGURASI SPREADSHEET
const SPREADSHEET_ID = 'waji diubah id nya';

// Fungsi untuk mendapatkan spreadsheet
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ====================================
// MAIN ENTRY POINT
// ====================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Aplikasi Absensi Sekolah')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://if.polibatam.ac.id/pamerin/uploads/pbl/3312011063/3312011063_gambar1_20220710.png');
}

// ====================================
// AUTENTIKASI & SESSION
// ====================================
function login(username, password, nisn) {
  try {
    const ss = getSpreadsheet();
    const usersSheet = ss.getSheetByName('users');
    const siswaSheet = ss.getSheetByName('siswa');

    let userFound = null;

    // -------------------------------------------
    // 1. LOGIKA CEK KREDENSIAL (SAMA SEPERTI LAMA)
    // -------------------------------------------
    
    // A. Login SISWA (Pakai NISN)
    if (nisn) {
      const siswaData = siswaSheet.getDataRange().getValues();
      for (let i = 1; i < siswaData.length; i++) {
        if (String(siswaData[i][1]) === String(nisn)) { 
          userFound = {
            role: 'siswa',
            identifier: siswaData[i][1], // NISN
            nama: siswaData[i][0],
            kelas: siswaData[i][8]
          };
          break;
        }
      }
      if (!userFound) return { success: false, message: 'NISN tidak ditemukan' };
    } 
    // B. Login ADMIN & GURU (Username/Pass)
    else {
      const userData = usersSheet.getDataRange().getValues();
      for (let i = 1; i < userData.length; i++) {
        if (userData[i][0] == username && userData[i][1] == password) {
          userFound = {
            role: userData[i][2],
            identifier: userData[i][0], // Username
            nama: userData[i][0],       // Gunakan username sbg nama
            kelas: userData[i][3] || "" 
          };
          break;
        }
      }
      if (!userFound) return { success: false, message: 'Username atau password salah' };
    }

    // -------------------------------------------
    // 2. GENERATE & SIMPAN TOKEN (KEAMANAN BARU)
    // -------------------------------------------
    
    // Buat Token Unik (UUID)
    const token = Utilities.getUuid();
    
    // Set Waktu Expired (24 Jam dari sekarang)
    const expiry = new Date();
    expiry.setTime(expiry.getTime() + (24 * 60 * 60 * 1000)); 

    // Siapkan Sheet Sessions
    let sessionSheet = ss.getSheetByName('sessions');
    if (!sessionSheet) {
      sessionSheet = ss.insertSheet('sessions');
      sessionSheet.appendRow(['Token', 'Identifier', 'Role', 'Expiry']);
      // Format kolom tanggal agar mudah dibaca (Opsional)
      sessionSheet.getRange("D:D").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    }

    // Simpan data sesi baru ke Sheet
    sessionSheet.appendRow([
      token, 
      userFound.identifier, 
      userFound.role, 
      expiry
    ]);

    // -------------------------------------------
    // 3. KEMBALIKAN TOKEN KE FRONTEND
    // -------------------------------------------
    return {
      success: true,
      token: token,      // Token dikirim ke frontend
      role: userFound.role,
      username: userFound.identifier,
      nama: userFound.nama,
      kelas: userFound.kelas,
      nisn: (userFound.role === 'siswa') ? userFound.identifier : null
    };

  } catch (error) {
    return { success: false, message: "Login Error: " + error.toString() };
  }
}


function verifyUser(token, requiredRole) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // Pastikan ID Spreadsheet global terbaca
  const sheet = ss.getSheetByName('sessions');
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    // data[i][0] = Token, data[i][2] = Role, data[i][3] = Expiry
    if (data[i][0] === token) {
      // Cek apakah token expired (misal 24 jam)
      if (data[i][3] instanceof Date && data[i][3] > now) {
        // Cek Role
        if (requiredRole && data[i][2] !== requiredRole && data[i][2] !== 'admin') {
           // Admin biasanya boleh akses semua, jika role user != required, tolak
           throw new Error("Akses Ditolak: Anda tidak memiliki izin.");
        }
        return true; // Valid
      } else {
        throw new Error("Sesi berakhir. Silakan login ulang.");
      }
    }
  }
  throw new Error("Token tidak valid atau tidak ditemukan.");
}
// ====================================
// SISWA - CRUD OPERATIONS
// ====================================
function getSiswaList() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('siswa');
    const data = sheet.getDataRange().getValues();
    const siswaList = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { 
        
        // --- PERUBAHAN LOGIKA BACA TANGGAL ---
        let rawTgl = data[i][3];
        let tglLahir = '';

        if (rawTgl instanceof Date) {
          // Jika Sheet masih membacanya sebagai Date Object
          tglLahir = Utilities.formatDate(rawTgl, 'Asia/Jakarta', 'yyyy-MM-dd');
        } else if (typeof rawTgl === 'string') {
          # Jika format string 'dd-mm-yyyy' (cth: 06-01-2026)
          // Kita harus balik jadi yyyy-mm-dd agar bisa dibaca input HTML
          let cleanTgl = rawTgl.replace(/'/g, "").trim(); // Hapus kutip jika ada
          if (cleanTgl.includes('-')) {
            let parts = cleanTgl.split('-');
            // Cek apakah formatnya dd-mm-yyyy (tahun di belakang)
            if (parts[2] && parts[2].length === 4) {
               tglLahir = parts[2] + '-' + parts[1] + '-' + parts[0];
            } else {
               tglLahir = cleanTgl; // Asumsi sudah benar
            }
          }
        }
        // ---------------------------------------

        siswaList.push({
          nama: data[i][0],
          nisn: data[i][1],
          jenisKelamin: data[i][2],
          tanggalLahir: tglLahir, // Hasil konversi untuk HTML Input
          agama: data[i][4],
          namaAyah: data[i][5],
          namaIbu: data[i][6],
          noHp: data[i][7],
          kelas: data[i][8],
          alamat: data[i][9]
        });
      }
    }
    
    return { success: true, data: siswaList };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getSiswaByNisn(nisn) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('siswa');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == nisn) {
        
        // FORMAT TANGGAL LAHIR DISINI JUGA
        let tglLahir = '';
        if (data[i][3]) {
          tglLahir = Utilities.formatDate(new Date(data[i][3]), 'Asia/Jakarta', 'yyyy-MM-dd');
        }

        return {
          success: true,
          data: {
            nama: data[i][0],
            nisn: data[i][1],
            jenisKelamin: data[i][2],
            tanggalLahir: tglLahir,
            agama: data[i][4],
            namaAyah: data[i][5],
            namaIbu: data[i][6],
            noHp: data[i][7],
            kelas: data[i][8],
            alamat: data[i][9]
          }
        };
      }
    }
    
    return { success: false, message: 'Siswa tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function addSiswa(token, siswaData) { 
  try {
    // --- 1. PASANG SATPAM DISINI ---
    // Cek apakah pengirim punya token admin yang valid?
    verifyUser(token, 'admin'); 
    // -------------------------------

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('siswa');

    // Cek NISN duplikat
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == siswaData.nisn) {
        return { success: false, message: 'NISN sudah terdaftar' };
      }
    }
    
    // Format Tanggal
    let tglSimpan = siswaData.tanggalLahir;
    if (tglSimpan && tglSimpan.includes('-')) {
      let parts = tglSimpan.split('-');
      tglSimpan = "'" + parts[2] + '-' + parts[1] + '-' + parts[0];
    }

    sheet.appendRow([
      siswaData.nama,
      siswaData.nisn,
      siswaData.jenisKelamin,
      tglSimpan,
      siswaData.agama,
      siswaData.namaAyah,
      siswaData.namaIbu,
      siswaData.noHp,
      siswaData.kelas,
      siswaData.alamat
    ]);
    return { success: true, message: 'Siswa berhasil ditambahkan (Aman)' };

  } catch (error) {
    // Jika token salah, akan masuk ke sini
    return { success: false, message: "GAGAL: " + error.message };
  }
}

function updateSiswa(oldNisn, siswaData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('siswa');
    const data = sheet.getDataRange().getValues();
    
    // --- PERUBAHAN DISINI: Format Tanggal Lahir (YYYY-MM-DD -> dd-MM-yyyy) ---
    let tglSimpan = siswaData.tanggalLahir;
    if (tglSimpan && tglSimpan.includes('-')) {
       // Cek apakah format input masih yyyy-mm-dd (default HTML)
       let parts = tglSimpan.split('-');
       if(parts[0].length === 4) { // Jika tahun di depan (2026-...)
           tglSimpan = "'" + parts[2] + '-' + parts[1] + '-' + parts[0];
       }
    }
    // ------------------------------------------------------------------------

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == oldNisn) {
        sheet.getRange(i + 1, 1, 1, 10).setValues([[
          siswaData.nama,
          siswaData.nisn,
          siswaData.jenisKelamin,
          tglSimpan, // Gunakan variabel yang sudah diformat
          siswaData.agama,
          siswaData.namaAyah,
          siswaData.namaIbu,
          siswaData.noHp,
          siswaData.kelas,
          siswaData.alamat
        ]]);
        return { success: true, message: 'Siswa berhasil diupdate' };
      }
    }
    
    return { success: false, message: 'Siswa tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteSiswa(token, nisn) {
  try {
    // 1. CEK KEAMANAN (Server Side Validation)
    // Hanya user dengan role 'admin' yang memiliki token valid yang boleh menghapus
    verifyUser(token, 'admin'); 

    // 2. LOGIKA HAPUS (Jika token valid, kode lanjut ke sini)
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('siswa');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      // Cek NISN pada kolom index 1
      if (String(data[i][1]) === String(nisn)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data siswa berhasil dihapus.' };
      }
    }
    
    return { success: false, message: 'Data siswa tidak ditemukan.' };

  } catch (error) {
    // Jika verifyUser gagal, error akan tertangkap di sini
    return { success: false, message: error.message };
  }
}

// ====================================
// GURU - CRUD OPERATIONS
// ====================================
function getGuruList(token) {
  try {
    // 1. CEK KEAMANAN (Server Side Validation)
    // Wajib role 'admin'. Token siswa/guru akan ditolak disini.
    verifyUser(token, 'admin'); 

    // 2. LOGIKA AMBIL DATA (Hanya jalan jika token valid & admin)
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('users');
    const data = sheet.getDataRange().getValues();
    const guruList = [];

    // Loop data mulai baris 1 (lewati header)
    for (let i = 1; i < data.length; i++) {
      // Cek kolom Role (Index 2) apakah 'guru'
      if (data[i][2] == 'guru') {
        guruList.push({
          username: data[i][0], // Kolom A
          password: data[i][1], // Kolom B (Sensitif)
          role: data[i][2],     // Kolom C
          kelas: data[i][3] || "" // Kolom D (Wali Kelas)
        });
      }
    }

    return { success: true, data: guruList };

  } catch (error) {
    // Jika verifyUser melempar error (Akses Ditolak), akan ditangkap disini
    return { success: false, message: error.message };
  }
}

function addGuru(token, username, password, kelas) { // Tambah param token
  try {
    verifyUser(token, 'admin'); // <-- SATPAM: Cek Token Admin

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == username) {
        return { success: false, message: 'Username sudah terdaftar' };
      }
    }
    sheet.appendRow([username, password, 'guru', kelas]);
    return { success: true, message: 'Guru berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: "Akses Ditolak: " + error.message };
  }
}

function updateGuru(token, oldUsername, newUsername, password, kelas) { // Tambah param token
  try {
    verifyUser(token, 'admin'); // <-- SATPAM

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('users');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == oldUsername && data[i][2] == 'guru') {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[newUsername, password, 'guru', kelas]]);
        return { success: true, message: 'Guru berhasil diupdate' };
      }
    }
    return { success: false, message: 'Guru tidak ditemukan' };
  } catch (error) {
    return { success: false, message: "Akses Ditolak: " + error.message };
  }
}

function deleteGuru(token, username) { // Tambah param token
  try {
    verifyUser(token, 'admin'); // <-- SATPAM

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('users');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == username && data[i][2] == 'guru') {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Guru berhasil dihapus' };
      }
    }
    return { success: false, message: 'Guru tidak ditemukan' };
  } catch (error) {
    return { success: false, message: "Akses Ditolak: " + error.message };
  }
}
// ====================================
// ABSENSI OPERATIONS
// ====================================
function scanAbsensi(nisn, scannerRole, scannerKelas) {
  try {
    const ss = getSpreadsheet();
    const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
    const nowTime = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'HH:mm');
    
    // 1. AMBIL KONFIGURASI
    const configResult = getAppConfig();
    const config = configResult.success ? configResult.data : {
      jam_masuk_akhir: '07:15',
      jam_pulang_mulai: '15:00',
      jam_pulang_akhir: '17:00'
    };

    // 2. CEK HARI LIBUR
    const liburSheet = ss.getSheetByName('hari_libur');
    if (liburSheet) {
      const liburData = liburSheet.getDataRange().getValues();
      for (let i = 1; i < liburData.length; i++) {
        if (liburData[i][0]) {
          let tglLibur = Utilities.formatDate(new Date(liburData[i][0]), 'Asia/Jakarta', 'yyyy-MM-dd');
          if (tglLibur === today) {
            return { success: false, message: 'Absensi DITUTUP. Hari ini libur: ' + liburData[i][1] };
          }
        }
      }
    }

    const absensiSheet = ss.getSheetByName('absensi');
    const siswaSheet = ss.getSheetByName('siswa');
    
    // 3. VALIDASI INPUT
    const scannedNisn = String(nisn).trim();
    if (scannedNisn === "" || scannedNisn === "undefined") {
      return { success: false, message: 'QR Code tidak valid atau kosong.' };
    }

    // 4. CARI DATA SISWA
    const siswaData = siswaSheet.getDataRange().getValues();
    let siswa = null;
    
    for (let i = 1; i < siswaData.length; i++) {
      if (String(siswaData[i][1]).trim() === scannedNisn) {
        siswa = {
          nama: siswaData[i][0],
          nisn: siswaData[i][1],
          kelas: siswaData[i][8]
        };
        break;
      }
    }
    
    if (!siswa) {
      return { success: false, message: 'NISN tidak terdaftar di database.' };
    }

    // VALIDASI KELAS GURU
    if (scannerRole === 'guru') {
      const kelasSiswa = String(siswa.kelas).trim().toUpperCase();
      const kelasGuru = String(scannerKelas).trim().toUpperCase();
      if (kelasGuru && kelasSiswa !== kelasGuru) {
        return { 
          success: false, 
          message: `Ditolak! Siswa ini kelas ${siswa.kelas}. Anda hanya bisa scan kelas ${scannerKelas}.` 
        };
      }
    }

    // 5. PROSES ABSENSI
    const absensiData = absensiSheet.getDataRange().getValues();
    
    for (let i = 1; i < absensiData.length; i++) {
      const rowDateCell = absensiData[i][0];
      if (!rowDateCell) continue;

      const rowDateStr = Utilities.formatDate(new Date(rowDateCell), 'Asia/Jakarta', 'yyyy-MM-dd');
      const rowNisn = String(absensiData[i][1]).trim();

      // === SKENARIO ABSEN PULANG (Data Hari Ini Ditemukan) ===
      if (rowDateStr === today && rowNisn === scannedNisn) {
        
        // Cek apakah sudah checkout sebelumnya (Kolom F / Index 5)
        if (absensiData[i][5]) { 
          return { success: false, message: 'Siswa sudah melakukan absen pulang hari ini.' };
        } else {
          
          // Cek Batas Akhir Pulang
          if (nowTime > config.jam_pulang_akhir) {
             return { 
               success: false, 
               message: `Gagal! Batas waktu pulang (${config.jam_pulang_akhir}) sudah lewat.` 
             };
          }

          // Cek Jeda Waktu (Mencegah double scan cepat)
          let jamDatangRaw = absensiData[i][4];
          let jamDatangStr = (jamDatangRaw instanceof Date) ? 
              Utilities.formatDate(jamDatangRaw, 'Asia/Jakarta', 'HH:mm') : 
              String(jamDatangRaw).substring(0, 5);
          
          const minutesDiff = calculateTimeDiff(jamDatangStr, nowTime);
          if (minutesDiff < 10) { // Jeda minimal 10 menit
             return { success: false, message: `Terlalu Cepat! Tunggu sebentar lagi.` };
          }

          // --- LOGIKA UPDATE PULANG (8 KOLOM) ---
          // Ambil keterangan yang sudah ada (misal: "Terlambat (5 m)") dari Kolom G (Index 6)
          let ketSaatIni = absensiData[i][6]; 
          let ketBaru = ketSaatIni;
          let pesanPulang = 'Absen Pulang Berhasil';

          // Cek Pulang Cepat
          if (nowTime < config.jam_pulang_mulai) {
             // Append status pulang cepat
             ketBaru = ketSaatIni + " & Pulang Cepat"; 
             pesanPulang = 'Absen Pulang (Pulang Cepat)';
          }

          const jamPulang = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'HH:mm:ss');
          
          // UPDATE KE SHEET
          // Baris ke-(i+1)
          // Kolom 6 (F) = Jam Pulang
          absensiSheet.getRange(i + 1, 6).setValue(jamPulang);
          // Kolom 7 (G) = Keterangan Waktu (Update ket baru)
          absensiSheet.getRange(i + 1, 7).setValue(ketBaru);
          // Kolom 8 (H) = Status (Tidak perlu diubah, tetap 'Hadir')
          
          return {
            success: true,
            message: pesanPulang,
            type: 'pulang',
            jamPulang: jamPulang,
            nama: siswa.nama,
            kelas: siswa.kelas,
            status: 'Hadir' // Status display
          };
        }
      }
    }

    // === SKENARIO ABSEN DATANG (Data Hari Ini Belum Ada) ===
    
    // Blokir jika sudah lewat jam operasional
    if (nowTime > config.jam_pulang_akhir) {
         return { success: false, message: `Absensi Ditutup! Sudah melewati jam operasional.` };
    }

    // --- LOGIKA DATANG (8 KOLOM) ---
    let keteranganWaktu = 'Tepat Waktu'; // Default
    let statusKehadiran = 'Hadir';       // Default

    if (nowTime > config.jam_masuk_akhir) {
      const lateMinutes = calculateTimeDiff(config.jam_masuk_akhir, nowTime);
      keteranganWaktu = `Terlambat (${lateMinutes} m)`;
      // Status tetap 'Hadir' meskipun terlambat
    }

    const jamDatang = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'HH:mm:ss');
    
    // INSERT DATA BARU (APPEND)
    // Urutan: [Tanggal, NISN, Nama, Kelas, JamDatang, JamPulang, Keterangan, Status]
    absensiSheet.appendRow([
      new Date(),        
      "'" + scannedNisn, 
      siswa.nama,        
      siswa.kelas,       
      jamDatang,         
      '',                // Jam Pulang kosong
      keteranganWaktu,   // Kolom G: Keterangan Waktu
      statusKehadiran    // Kolom H: Status Kehadiran
    ]);

    let responseMessage = 'Absen Masuk Berhasil';
    if (keteranganWaktu.includes('Terlambat')) {
       responseMessage = `Absen Masuk (${keteranganWaktu})`;
    }

    return {
      success: true,
      message: responseMessage,
      type: 'datang',
      jamDatang: jamDatang,
      nama: siswa.nama,
      kelas: siswa.kelas,
      status: statusKehadiran
    };

  } catch (error) {
    return { success: false, message: "Error Server: " + error.toString() };
  }
}

// Pastikan helper ini ada di paling bawah file code.gs
function calculateTimeDiff(startTime, endTime) {
  const [h1, m1] = startTime.split(':').map(Number);
  const [h2, m2] = endTime.split(':').map(Number);
  
  const totalMinutes1 = h1 * 60 + m1;
  const totalMinutes2 = h2 * 60 + m2;
  
  return totalMinutes2 - totalMinutes1;
}

function getAbsensiToday(nisn) {
  try {
    const ss = getSpreadsheet();
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');

    // --- CEK APAKAH HARI INI LIBUR? ---
    const liburSheet = ss.getSheetByName('hari_libur');
    let isLibur = false;
    let keteranganLibur = "";

    if (liburSheet) {
      const liburData = liburSheet.getDataRange().getValues();
      for (let i = 1; i < liburData.length; i++) {
        let tgl = Utilities.formatDate(new Date(liburData[i][0]), 'Asia/Jakarta', 'yyyy-MM-dd');
        if (tgl === todayStr) {
          isLibur = true;
          keteranganLibur = liburData[i][1];
          break;
        }
      }
    }
    // ----------------------------------

    const sheet = ss.getSheetByName('absensi');
    const data = sheet.getDataRange().getValues();
    const searchNisn = String(nisn).trim();

    let absensiData = null;
    
    // Loop cari data absen siswa hari ini
    for (let i = 1; i < data.length; i++) {
      const rowDateCell = data[i][0];
      if (!rowDateCell) continue;
      
      const rowDateStr = Utilities.formatDate(new Date(rowDateCell), 'Asia/Jakarta', 'yyyy-MM-dd');
      const rowNisn = String(data[i][1]).trim();

      if (rowDateStr === todayStr && rowNisn === searchNisn) {
        // Format Jam Datang
        let jamDatang = data[i][4];
        if (jamDatang instanceof Date) {
          jamDatang = Utilities.formatDate(jamDatang, 'Asia/Jakarta', 'HH:mm:ss');
        }
        
        // Format Jam Pulang
        let jamPulang = data[i][5];
        if (jamPulang instanceof Date) {
          jamPulang = Utilities.formatDate(jamPulang, 'Asia/Jakarta', 'HH:mm:ss');
        } else if (!jamPulang) {
          jamPulang = ""; // Pastikan string kosong jika belum pulang
        }

        absensiData = {
          tanggal: rowDateStr,
          jamDatang: jamDatang,
          jamPulang: jamPulang,
          status: data[i][6]
        };
        break; 
      }
    }

    // Kembalikan data absen BESERTA status libur
    return { 
      success: true, 
      data: absensiData,
      isLibur: isLibur,
      keteranganLibur: keteranganLibur
    };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ====================================
// GANTI FUNCTION getAbsensiList LAMA DENGAN INI
// ====================================
function getAbsensiList(filter = {}) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('absensi');
    const data = sheet.getDataRange().getValues();
    const absensiList = [];
    
    const fStart = filter.tanggalMulai || "";
    const fEnd = filter.tanggalAkhir || "";
    const fKelas = filter.kelas || ""; // Filter Kelas

    // Loop data (Mulai dari baris ke-1, lewati header)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        
        let rawDate = new Date(data[i][0]);
        let tanggalStr = Utilities.formatDate(rawDate, 'Asia/Jakarta', 'yyyy-MM-dd');

        // Format Jam
        let jamDatangStr = data[i][4];
        if (data[i][4] instanceof Date) {
             jamDatangStr = Utilities.formatDate(data[i][4], 'Asia/Jakarta', 'HH:mm:ss');
        }

        let jamPulangStr = data[i][5];
        if (data[i][5] && data[i][5] instanceof Date) {
             jamPulangStr = Utilities.formatDate(data[i][5], 'Asia/Jakarta', 'HH:mm:ss');
        } else if (!jamPulangStr) {
             jamPulangStr = "-";
        }
        
        // Mapping Data (SESUAI STRUKTUR 8 KOLOM)
        const item = {
          tanggal: tanggalStr, 
          nisn: data[i][1],
          nama: data[i][2],
          kelas: data[i][3],
          jamDatang: jamDatangStr,
          jamPulang: jamPulangStr,
          keterangan: data[i][6], // Kolom G: Keterangan Waktu
          status: data[i][7]      // Kolom H: Status Kehadiran
        };

        // Logika Filter
        let match = true;
        if (fStart && tanggalStr < fStart) match = false;
        if (fEnd && tanggalStr > fEnd) match = false;
        if (filter.nama && !String(item.nama).toLowerCase().includes(filter.nama.toLowerCase())) match = false;
        if (fKelas && item.kelas != fKelas) match = false;

        if (match) {
          absensiList.push(item);
        }
      }
    }
    
    // Urutkan data dari yang terbaru
    absensiList.sort((a, b) => b.tanggal.localeCompare(a.tanggal));

    return { success: true, data: absensiList };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getKelasList() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('siswa');
    const data = sheet.getDataRange().getValues();
    const kelasSet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][8]) {
        kelasSet.add(data[i][8]);
      }
    }
    
    return { success: true, data: Array.from(kelasSet).sort() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ====================================
// SETUP INITIAL DATA
// ====================================
// ====================================
// SETUP INITIAL DATA (DATABASE)
// ====================================
function setupInitialData() {
  try {
    const ss = getSpreadsheet();

    // 1. Setup sheet 'users' (Admin & Guru)
    let usersSheet = ss.getSheetByName('users');
    if (!usersSheet) {
      usersSheet = ss.insertSheet('users');
      // Header: Username, Password, Role, Kelas (Wajib ada kolom Kelas untuk Guru)
      usersSheet.appendRow(['Username', 'Password', 'Role', 'Kelas']);
      
      // Data Default Admin
      usersSheet.appendRow(['admin', 'admin123', 'admin', '']);
      // Data Default Guru (Contoh Wali Kelas)
      usersSheet.appendRow(['guru1', 'guru123', 'guru', 'VI B']);
    }

    // 2. Setup sheet 'siswa'
    let siswaSheet = ss.getSheetByName('siswa');
    if (!siswaSheet) {
      siswaSheet = ss.insertSheet('siswa');
      siswaSheet.appendRow([
        'Nama Lengkap', 
        'NISN', 
        'Jenis Kelamin', 
        'Tanggal Lahir', 
        'Agama',
        'Nama Ayah', 
        'Nama Ibu', 
        'No Handphone', 
        'Kelas', 
        'Alamat'
      ]);
      // Sample Data Siswa
      siswaSheet.appendRow([
        'Ahmad Rizki', 
        '1234567890', 
        'Laki-laki', 
        '2008-05-15', 
        'Islam',
        'Budi Santoso', 
        'Siti Aminah', 
        '081234567890', 
        'VI B',
        'Jl. Merdeka No. 10, Bengkulu'
      ]);
    }

    // 3. Setup sheet 'absensi' (DIPERBARUI: 8 KOLOM)
    let absensiSheet = ss.getSheetByName('absensi');
    if (!absensiSheet) {
      absensiSheet = ss.insertSheet('absensi');
      // Perubahan struktur header:
      // Kolom G (Index 7) = Keterangan Waktu (Terlambat/Pulang Cepat)
      // Kolom H (Index 8) = Status (Hadir/Sakit/Izin/Alpa)
      absensiSheet.appendRow([
        'Tanggal', 
        'NISN', 
        'Nama', 
        'Kelas', 
        'Jam Datang', 
        'Jam Pulang', 
        'Keterangan Waktu', 
        'Status'
      ]);
    }

    // 4. Setup sheet 'hari_libur'
    let liburSheet = ss.getSheetByName('hari_libur');
    if (!liburSheet) {
      liburSheet = ss.insertSheet('hari_libur');
      liburSheet.appendRow(['Tanggal', 'Keterangan']);
    }

    // 5. Setup sheet 'konfigurasi'
    let configSheet = ss.getSheetByName('konfigurasi');
    if (!configSheet) {
      configSheet = ss.insertSheet('konfigurasi');
      configSheet.appendRow(['Key', 'Value', 'Keterangan']);
      
      // Default Config
      configSheet.appendRow(['jam_masuk_mulai', '06:00', 'Waktu absen datang dibuka']);
      configSheet.appendRow(['jam_masuk_akhir', '07:15', 'Batas waktu terlambat']);
      configSheet.appendRow(['jam_pulang_mulai', '15:00', 'Waktu absen pulang dibuka']);
      configSheet.appendRow(['jam_pulang_akhir', '17:00', 'Batas akhir absen pulang']);
    }
    
    return { success: true, message: 'Setup database berhasil. Struktur 8 kolom diterapkan pada sheet Absensi.' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
// ====================================
// KELOLA HARI LIBUR (BARU)
// ====================================
function getHariLibur() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('hari_libur');
    const data = sheet.getDataRange().getValues();
    const list = [];
    
    // Loop dari baris 1 (lewati header)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        let tgl = Utilities.formatDate(new Date(data[i][0]), 'Asia/Jakarta', 'yyyy-MM-dd');
        list.push({
          tanggal: tgl,
          keterangan: data[i][1]
        });
      }
    }
    // Urutkan tanggal descending
    list.sort((a, b) => b.tanggal.localeCompare(a.tanggal));
    return { success: true, data: list };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function addHariLibur(tanggal, keterangan) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('hari_libur');
    
    // Validasi format tanggal string yyyy-mm-dd
    sheet.appendRow([tanggal, keterangan]);
    return { success: true, message: 'Hari libur berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteHariLibur(tanggalStr) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('hari_libur');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let rowDate = Utilities.formatDate(new Date(data[i][0]), 'Asia/Jakarta', 'yyyy-MM-dd');
      if (rowDate === tanggalStr) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Hari libur dihapus' };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ====================================
// FITUR MONITORING & UPDATE STATUS (BARU)
// ====================================

// Function untuk mengambil data Monitoring (Versi Update: Pisah Terlambat & Pulang Cepat)
function getMonitoringRealtime(filterKelas = null) {
  try {
    const ss = getSpreadsheet();
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
    const siswaSheet = ss.getSheetByName('siswa');
    const dataSiswa = siswaSheet.getDataRange().getValues();
    const absensiSheet = ss.getSheetByName('absensi');
    const dataAbsensi = absensiSheet.getDataRange().getValues();
    
    // Mapping data absensi hari ini
    let absensiMap = {};
    for (let i = 1; i < dataAbsensi.length; i++) {
      let rowDate = dataAbsensi[i][0];
      if (!rowDate) continue; // Skip baris kosong

      let tgl = Utilities.formatDate(new Date(rowDate), 'Asia/Jakarta', 'yyyy-MM-dd');
      let nisn = String(dataAbsensi[i][1]).trim();
      
      if (tgl === todayStr) {
        absensiMap[nisn] = {
          jamDatang: dataAbsensi[i][4],
          jamPulang: dataAbsensi[i][5],
          keterangan: dataAbsensi[i][6], // Kolom G (Index 6) -> Terlambat/Tepat Waktu
          status: dataAbsensi[i][7]      // Kolom H (Index 7) -> Hadir/Sakit/Izin
        };
      }
    }

    let result = [];
    for (let i = 1; i < dataSiswa.length; i++) {
      let nama = dataSiswa[i][0];
      let nisn = String(dataSiswa[i][1]).trim();
      let kelas = dataSiswa[i][8];

      // Filter Kelas
      if (filterKelas && kelas !== filterKelas) continue;
      
      let statusInfo = absensiMap[nisn];
      
      // Default Value (Jika siswa belum absen)
      let jamDatang = '-';
      let jamPulang = '-';
      let displayStatus = 'Belum Absen'; 
      let keteranganWaktu = '-';         

      if (statusInfo) {
        // 1. Ambil Jam
        if (statusInfo.jamDatang instanceof Date) {
            jamDatang = Utilities.formatDate(statusInfo.jamDatang, 'Asia/Jakarta', 'HH:mm');
        } else if (statusInfo.jamDatang) jamDatang = String(statusInfo.jamDatang);

        if (statusInfo.jamPulang instanceof Date) {
            jamPulang = Utilities.formatDate(statusInfo.jamPulang, 'Asia/Jakarta', 'HH:mm');
        } else if (statusInfo.jamPulang) jamPulang = String(statusInfo.jamPulang);

        // 2. Ambil Status & Keterangan LANGSUNG DARI SHEET
        // Kita tidak mengubah logika di sini, kita percaya data di Sheet sudah benar
        let rawKet = statusInfo.keterangan; 
        let rawStat = statusInfo.status;

        displayStatus = rawStat ? String(rawStat) : "";

        // Logika tampilan Keterangan Waktu
        if (rawKet && String(rawKet).trim() !== "") {
            // Jika di sheet ada tulisan (misal: "Terlambat (900 m)"), tampilkan itu
            keteranganWaktu = String(rawKet);
        } else {
            // Jika kosong di sheet, tapi status Hadir, anggap Tepat Waktu
            if (displayStatus === 'Hadir') {
                keteranganWaktu = 'Tepat Waktu';
            } else {
                keteranganWaktu = '-';
            }
        }
      }

      result.push({
        nama: nama,
        nisn: nisn,
        kelas: kelas,
        jamDatang: jamDatang,
        jamPulang: jamPulang,
        status: displayStatus,       // Dropdown
        keterangan: keteranganWaktu  // Kolom Teks (Terlambat/Tepat Waktu)
      });
    }

    // Sort: Kelas dulu, baru Nama
    result.sort((a, b) => {
      if (a.kelas === b.kelas) return a.nama.localeCompare(b.nama);
      return a.kelas.localeCompare(b.kelas);
    });
    
    return { success: true, data: result };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}



function updateAbsensiStatus(token, nisn, nama, kelas, newStatus) {
  try {
    // SATPAM: Hanya Guru atau Admin yang boleh ubah status manual
    verifyUser(token, 'guru');
    
    const ss = getSpreadsheet();
    const absensiSheet = ss.getSheetByName('absensi');
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
    const data = absensiSheet.getDataRange().getValues();
    
    let found = false;
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      let tgl = Utilities.formatDate(new Date(data[i][0]), 'Asia/Jakarta', 'yyyy-MM-dd');
      let rowNisn = String(data[i][1]).trim();
      
      if (tgl === todayStr && rowNisn === String(nisn).trim()) {
        found = true;
        rowIndex = i + 1;
        break;
      }
    }

    if (found) {
      // PERBAIKAN 1: Ubah dari kolom 7 ke kolom 8
      // Kolom 7 = Keterangan Waktu
      // Kolom 8 = Status (Hadir/Sakit/Izin/Alpa)
      absensiSheet.getRange(rowIndex, 8).setValue(newStatus); 
    } else {
      let jamDatang = '-';
      if (newStatus === 'Hadir') {
        jamDatang = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'HH:mm:ss');
      }
      
      // PERBAIKAN 2: Sesuaikan urutan array appendRow agar masuk ke kolom yang benar
      // Struktur: [Tanggal, NISN, Nama, Kelas, JamDt, JamPlg, Keterangan(Col 7), Status(Col 8)]
      // Kita isi Keterangan (Col 7) dengan "-" atau kosong, lalu Status (Col 8) dengan newStatus
      absensiSheet.appendRow([
        new Date(), 
        "'" + nisn, 
        nama, 
        kelas, 
        jamDatang, 
        '',   // Jam Pulang
        '-',  // Kolom 7 (Keterangan Waktu) -> Diisi strip agar tidak error
        newStatus // Kolom 8 (Status Kehadiran) -> Target yang benar
      ]);
    }

    return { success: true, message: 'Status berhasil diubah' };
  } catch (error) {
    return { success: false, message: "Gagal: " + error.message };
  }
}

function updateHariLibur(oldDateStr, newDateStr, newKeterangan) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('hari_libur');
    const data = sheet.getDataRange().getValues();
    
    let found = false;
    
    // Loop dari baris 1 (lewati header)
    for (let i = 1; i < data.length; i++) {
      // Format tanggal dari sheet agar sama dengan format string input (yyyy-MM-dd)
      let rowDate = Utilities.formatDate(new Date(data[i][0]), 'Asia/Jakarta', 'yyyy-MM-dd');
      
      if (rowDate === oldDateStr) {
        // Update baris: Kolom 1 (Tanggal), Kolom 2 (Keterangan)
        // Gunakan new Date() untuk kolom tanggal agar format di sheet tetap Date Object
        sheet.getRange(i + 1, 1, 1, 2).setValues([[new Date(newDateStr), newKeterangan]]);
        found = true;
        break;
      }
    }
    
    if (found) {
      return { success: true, message: 'Hari libur berhasil diperbarui' };
    } else {
      return { success: false, message: 'Data tanggal lama tidak ditemukan' };
    }
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ====================================
// FITUR EXPORT EXCEL
// ====================================


function getExportData(type, filters) {
  const ss = getSpreadsheet();
  
  // --- 1. LOGIKA UNTUK LAPORAN PERIODE (YANG SUDAH ADA) ---
  if (type === 'laporan_absensi') {
    const sheet = ss.getSheetByName('absensi');
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    const fStart = filters.tanggalMulai || "";
    const fEnd = filters.tanggalAkhir || "";
    const fKelas = filters.kelas || ""; 
    
    let no = 1;
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      let rawDate = new Date(data[i][0]);
      let tanggalStr = Utilities.formatDate(rawDate, 'Asia/Jakarta', 'dd-MM-yyyy');
      let dateForFilter = Utilities.formatDate(rawDate, 'Asia/Jakarta', 'yyyy-MM-dd');
      let rowKelas = data[i][3];

      let match = true;
      if (fStart && dateForFilter < fStart) match = false;
      if (fEnd && dateForFilter > fEnd) match = false;
      if (fKelas && rowKelas != fKelas) match = false;

      if (match) {
        let jamDatang = data[i][4];
        if (data[i][4] instanceof Date) {
             jamDatang = Utilities.formatDate(data[i][4], 'Asia/Jakarta', 'HH:mm:ss');
        }
        let jamPulang = data[i][5];
        if (data[i][5] && data[i][5] instanceof Date) {
             jamPulang = Utilities.formatDate(data[i][5], 'Asia/Jakarta', 'HH:mm:ss');
        }

        result.push([
          no++,                        
          tanggalStr,             
          "'" + data[i][1], 
          data[i][2],                  
          rowKelas,                    
          jamDatang,             
          jamPulang || '-',            
          data[i][6],                  
          data[i][7]                   
        ]);
      }
    }
    return result;
  }
  
  // --- 2. LOGIKA BARU: UNTUK MONITORING (REALTIME HARI INI) ---
  else if (type === 'monitoring') {
    // Kita gunakan ulang fungsi getMonitoringRealtime agar datanya konsisten
    // filters.kelas bisa dikirim jika yang request adalah Guru Wali Kelas
    const realtimeData = getMonitoringRealtime(filters.kelas);
    
    if (!realtimeData.success) return [];
    
    const data = realtimeData.data; // Array object hasil monitoring
    const result = [];
    
    // Mapping dari JSON Object ke Array Row Excel
    // Header nanti: No, Nama Siswa, NISN, Kelas, Jam Datang, Jam Pulang, Keterangan, Status
    data.forEach((item, index) => {
       result.push([
         index + 1,
         item.nama,
         "'" + item.nisn, // Pakai kutip agar tidak jadi scientific number
         item.kelas,
         item.jamDatang,
         item.jamPulang,
         item.keterangan, // Terlambat / Tepat Waktu
         item.status      // Hadir / Sakit / Izin / Alpa / Belum Absen
       ]);
    });
    
    return result;
  }
  
  return [];
}

function generateExcel(type, filters) {
  try {
    // 1. SETUP FILE & SHEET
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd-MM-yyyy HHmm');
    var fileName = "";
    var headers = [];

    // TENTUKAN HEADER & NAMA FILE BERDASARKAN TIPE
    if (type === 'laporan_absensi') {
        fileName = "Laporan Absensi - " + timestamp;
        // Header 9 Kolom
        headers = ["No", "Tanggal", "NISN", "Nama Siswa", "Kelas", "Jam Datang", "Jam Pulang", "Keterangan Waktu", "Status Kehadiran"];
    } 
    else if (type === 'monitoring') {
        fileName = "Monitoring Harian - " + timestamp;
        // Header 8 Kolom (Tanpa Tanggal karena ini laporan harian)
        headers = ["No", "Nama Siswa", "NISN", "Kelas", "Jam Datang", "Jam Pulang", "Keterangan Waktu", "Status Terkini"];
    }

    // Buat Spreadsheet Sementara
    var ss = SpreadsheetApp.create(fileName);
    var sheet = ss.getActiveSheet();
    
    // Ambil Data (Pastikan fungsi getExportData sudah diupdate juga)
    var data = getExportData(type, filters);

    // 2. TULIS HEADER
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);

    // --- STYLING HEADER MODERN ---
    headerRange
      .setFontWeight('bold')
      .setFontColor('#FFFFFF')           
      .setBackground('#4F46E5')          // Indigo 600 (Sesuai tema aplikasi)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontFamily('Roboto')
      .setFontSize(11);
    sheet.setRowHeight(1, 45); 
    
    // 3. TULIS DATA BODY
    if (data && data.length > 0) {
      var numRows = data.length;
      var numCols = headers.length;
      var startRow = 2;
      
      // Pembersihan Data (Deep Copy)
      var cleanData = data.map(function(row) { return row; });
      var dataRange = sheet.getRange(startRow, 1, numRows, numCols);
      dataRange.setValues(cleanData);
      
      // --- STYLING DATA BODY ---
      dataRange
        .setFontFamily('Roboto')
        .setFontSize(10)
        .setVerticalAlignment('middle');
      sheet.setRowHeights(startRow, numRows, 30); 
      
      // ALIGNMENT (Perataan Teks)
      // Default: Center semua
      dataRange.setHorizontalAlignment('center');
      
      // Khusus Kolom Nama (Biasanya kolom ke-2 di Monitoring, ke-4 di Laporan)
      // Kita cari index kolom yang judulnya mengandung "Nama"
      var namaColIndex = headers.findIndex(h => h.includes("Nama"));
      if (namaColIndex > -1) {
          sheet.getRange(startRow, namaColIndex + 1, numRows, 1).setHorizontalAlignment('left');
      }

      // BORDERS (Tipis & Rapi warna abu muda)
      dataRange.setBorder(true, true, true, true, true, true, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);

      // ZEBRA STRIPING (Warna selang-seling)
      for (var i = 0; i < numRows; i++) {
        if (i % 2 === 1) { 
           sheet.getRange(startRow + i, 1, 1, numCols).setBackground('#F8FAFC');
        }
      }

      // --- CONDITIONAL FORMATTING (WARNA STATUS OTOMATIS) ---
      // Logika dinamis: Kolom Status selalu kolom terakhir, Keterangan kolom sebelum terakhir
      var statusColIndex = headers.length; 
      var ketColIndex = headers.length - 1;

      var statusRange = sheet.getRange(startRow, statusColIndex, numRows, 1);
      var ketRange = sheet.getRange(startRow, ketColIndex, numRows, 1);
      
      var rules = sheet.getConditionalFormatRules();

      // 1. Status: Hadir (Hijau)
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Hadir")
        .setBackground("#DCFCE7") .setFontColor("#166534") .setBold(true)
        .setRanges([statusRange])
        .build());

      // 2. Status: Sakit (Kuning)
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Sakit")
        .setBackground("#FEF9C3") .setFontColor("#854D0E") .setBold(true)
        .setRanges([statusRange])
        .build());

      // 3. Status: Izin (Biru)
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Izin")
        .setBackground("#DBEAFE") .setFontColor("#1E40AF") .setBold(true)
        .setRanges([statusRange])
        .build());

      // 4. Status: Alpa (Merah)
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Alpa")
        .setBackground("#FEE2E2") .setFontColor("#991B1B") .setBold(true)
        .setRanges([statusRange])
        .build());
        
      // 5. Status: Belum Absen (Abu-abu - Khusus Monitoring)
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Belum Absen")
        .setBackground("#F3F4F6") .setFontColor("#6B7280") .setItalic(true)
        .setRanges([statusRange])
        .build());

      // 6. Keterangan: Terlambat (Teks Merah)
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Terlambat")
        .setFontColor("#DC2626") .setBold(true)
        .setRanges([ketRange])
        .build());
        
      // 7. Keterangan: Pulang Cepat (Teks Oranye)
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Pulang Cepat")
        .setFontColor("#C2410C") .setBold(true)
        .setRanges([ketRange])
        .build());

      sheet.setConditionalFormatRules(rules);
    }
    
    // 4. FINISHING TOUCHES
    sheet.autoResizeColumns(1, headers.length);
    // Padding manual agar kolom lebih lega
    for(var c=1; c<=headers.length; c++) {
       var w = sheet.getColumnWidth(c);
       sheet.setColumnWidth(c, w + 20); 
    }
    
    sheet.setFrozenRows(1);
    sheet.setHiddenGridlines(true);

    // 5. GENERATE DOWNLOAD LINK
    var fileId = ss.getId();
    var file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // URL download langsung format XLSX
    var downloadUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
    
    return { success: true, url: downloadUrl };

  } catch (e) {
    return { success: false, message: 'Gagal generate Excel: ' + e.toString() };
  }
}

function getAppConfig() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('konfigurasi');
    // Default config jika sheet belum ada/kosong
    let config = {
      jam_masuk_mulai: '06:00',
      jam_masuk_akhir: '07:15',
      jam_pulang_mulai: '15:00',
      jam_pulang_akhir: '17:00'
    };
    
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const key = data[i][0];
        const val = data[i][1];
        if (config.hasOwnProperty(key)) {
          // Pastikan format HH:mm (terkadang Google Sheet menyimpan sebagai Date)
          if (val instanceof Date) {
            config[key] = Utilities.formatDate(val, 'Asia/Jakarta', 'HH:mm');
          } else {
            config[key] = String(val);
          }
        }
      }
    }
    return { success: true, data: config };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// 3. Simpan Konfigurasi dari Frontend
function saveAppConfig(newConfig) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('konfigurasi');
    if (!sheet) return { success: false, message: 'Sheet konfigurasi tidak ditemukan' };
    
    // Kita update baris berdasarkan Key (Asumsi urutan tidak berubah, tapi lebih aman cari key)
    const data = sheet.getDataRange().getValues();
    
    // Helper untuk update
    const updateRow = (key, val) => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          // Format value agar tersimpan sebagai Text di Sheet (diberi tanda kutip satu di awal) 
          // atau string biasa. Untuk jam lebih aman string.
          sheet.getRange(i + 1, 2).setValue("'" + val); 
          return;
        }
      }
    };

    updateRow('jam_masuk_mulai', newConfig.jam_masuk_mulai);
    updateRow('jam_masuk_akhir', newConfig.jam_masuk_akhir);
    updateRow('jam_pulang_mulai', newConfig.jam_pulang_mulai);
    updateRow('jam_pulang_akhir', newConfig.jam_pulang_akhir);

    return { success: true, message: 'Konfigurasi waktu berhasil disimpan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// Helper: Menghitung selisih menit antara dua waktu (HH:mm)
function calculateTimeDiff(startTime, endTime) {
  const [h1, m1] = startTime.split(':').map(Number);
  const [h2, m2] = endTime.split(':').map(Number);
  
  const totalMinutes1 = h1 * 60 + m1;
  const totalMinutes2 = h2 * 60 + m2;
  
  return totalMinutes2 - totalMinutes1;
}

// ====================================
// IMPORT SISWA DARI EXCEL (BULK)
// ====================================
function importSiswaBulk(dataArray) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('siswa');
    const existingData = sheet.getDataRange().getValues();
    
    // 1. Ambil daftar NISN yang sudah ada untuk cek duplikasi
    const existingNISN = new Set();
    for (let i = 1; i < existingData.length; i++) {
      existingNISN.add(String(existingData[i][1]).trim());
    }

    const rowsToAdd = [];
    let addedCount = 0;
    let skippedCount = 0;

    // 2. Loop data import
    for (let i = 0; i < dataArray.length; i++) {
      const item = dataArray[i];
      const nisn = String(item.nisn).trim();

      // Validasi dasar
      if (!item.nama || !nisn) {
        skippedCount++;
        continue;
      }

      // Cek Duplikasi NISN
      if (existingNISN.has(nisn)) {
        skippedCount++;
        continue;
      }

      // Format Tanggal Lahir (Jika Excel mengirim format angka tanggal)
      // SheetJS kadang mengirim string, kadang angka. Kita simpan string aman.
      let tglLahir = item.tanggalLahir;
      
      // Siapkan baris
      rowsToAdd.push([
        item.nama,
        "'" + nisn, // Pakai kutip satu agar format text terjaga
        item.jenisKelamin,
        tglLahir,
        item.agama,
        item.namaAyah,
        item.namaIbu,
        "'" + item.noHp,
        item.kelas,
        item.alamat
      ]);
      
      // Tambahkan ke Set agar tidak duplikat di dalam file import itu sendiri
      existingNISN.add(nisn);
      addedCount++;
    }

    // 3. Tulis ke Sheet sekaligus (Batch Operation) agar cepat
    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }

    return { 
      success: true, 
      added: addedCount, 
      skipped: skippedCount, 
      message: `Import selesai. Berhasil: ${addedCount}, Duplikat/Gagal: ${skippedCount}` 
    };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ====================================
// IMPORT GURU DARI EXCEL (BULK)
// ====================================
function importGuruBulk(dataArray) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('users');
    const existingData = sheet.getDataRange().getValues();
    
    // 1. Ambil daftar Username yang sudah ada untuk cek duplikasi
    const existingUsernames = new Set();
    for (let i = 1; i < existingData.length; i++) {
      existingUsernames.add(String(existingData[i][0]).trim());
    }

    const rowsToAdd = [];
    let addedCount = 0;
    let skippedCount = 0;

    // 2. Loop data import
    for (let i = 0; i < dataArray.length; i++) {
      const item = dataArray[i];
      const username = String(item.username).trim();

      // Validasi dasar
      if (!username || !item.password) {
        skippedCount++;
        continue;
      }

      // Cek Duplikasi Username
      if (existingUsernames.has(username)) {
        skippedCount++;
        continue;
      }

      // Siapkan baris: [Username, Password, Role, Kelas]
      rowsToAdd.push([
        "'" + username, // Pakai kutip satu agar format text terjaga
        "'" + item.password,
        'guru',         // Role otomatis di-set 'guru'
        item.kelas || '' // Kelas opsional
      ]);

      // Tambahkan ke Set agar tidak duplikat di dalam file import itu sendiri
      existingUsernames.add(username);
      addedCount++;
    }

    // 3. Tulis ke Sheet sekaligus (Batch Operation)
    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }

    return { 
      success: true, 
      added: addedCount, 
      skipped: skippedCount, 
      message: `Import selesai. Berhasil: ${addedCount}, Duplikat/Gagal: ${skippedCount}` 
    };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

