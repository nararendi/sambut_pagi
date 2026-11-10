/**
 * ==============================================================================
 * BACKEND GOOGLE APPS SCRIPT - SISTEM PRESENSI SAMBUT PAGI SMK ASSALAAM
 * ==============================================================================
 * File: Code.gs
 * Kolom Pegawai Sederhana: No, Nama Guru Karyawan, Jenis Kelamin, Jabatan Tugas
 * Kolom Presensi: ID Presensi, Tanggal, ID Pegawai, Nama Pegawai, Jabatan, 
 *                 Jam Kehadiran, Status Ketepatan, Status (Sambut Pagi / Tidak Sambut Pagi), 
 *                 Keterangan, Timestamp Sistem
 * ==============================================================================
 */

// Konfigurasi Nama Sheet & Jam Masuk
const CONFIG = {
  SHEET_PEGAWAI: 'Data_Pegawai',
  SHEET_PRESENSI: 'Riwayat_Presensi',
  SHEET_USERS: 'Data_Users',
  TIMEZONE: 'Asia/Jakarta',
  BATAS_SAMBUT_PAGI: '06:35', // Batas waktu tepat waktu sambut pagi
  BATAS_MASUK_KERJA: '06:40', // Batas toleransi masuk kerja
  MAX_RECENT_DAYS: 35         // Batas riwayat default agar hemat kuota eksekusi Google Apps Script
};

/**
 * 1. FUNGSI SETUP AWAL DATABASE
 * Jalankan fungsi ini SEKALI di editor Apps Script untuk membuat lembar kerja (Sheet)
 * dan header kolom secara otomatis dengan format yang rapi.
 */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Sheet Data_Pegawai
  let sheetPegawai = ss.getSheetByName(CONFIG.SHEET_PEGAWAI);
  if (!sheetPegawai) {
    sheetPegawai = ss.insertSheet(CONFIG.SHEET_PEGAWAI);
  }

  // Format Kolom: ID, No, Nama Guru Karyawan, Jenis Kelamin, Jabatan Tugas, Waktu Dibuat
  const headersPegawai = [
    'ID', 'No', 'Nama Guru Karyawan', 'Jenis Kelamin', 
    'Jabatan Tugas', 'Waktu Dibuat'
  ];

  if (sheetPegawai.getLastRow() === 0) {
    sheetPegawai.appendRow(headersPegawai);
    const headerRange = sheetPegawai.getRange(1, 1, 1, headersPegawai.length);
    headerRange.setBackground('#0d1527')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    sheetPegawai.setFrozenRows(1);
    sheetPegawai.setColumnWidth(1, 140); // ID
    sheetPegawai.setColumnWidth(2, 60);  // No
    sheetPegawai.setColumnWidth(3, 260); // Nama Guru Karyawan
    sheetPegawai.setColumnWidth(4, 130); // Jenis Kelamin
    sheetPegawai.setColumnWidth(5, 240); // Jabatan Tugas
    sheetPegawai.setColumnWidth(6, 170); // Waktu Dibuat
  }

  // 2. Setup Sheet Riwayat_Presensi
  let sheetPresensi = ss.getSheetByName(CONFIG.SHEET_PRESENSI);
  if (!sheetPresensi) {
    sheetPresensi = ss.insertSheet(CONFIG.SHEET_PRESENSI);
  }

  const headersPresensi = [
    'ID Presensi', 'Tanggal', 'ID Pegawai', 'Nama Pegawai', 
    'Jabatan', 'Jam Kehadiran', 'Status Ketepatan', 'Status', 
    'Keterangan', 'Timestamp Sistem'
  ];

  if (sheetPresensi.getLastRow() === 0) {
    sheetPresensi.appendRow(headersPresensi);
    const headerRange = sheetPresensi.getRange(1, 1, 1, headersPresensi.length);
    headerRange.setBackground('#10b981')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    sheetPresensi.setFrozenRows(1);
    sheetPresensi.setColumnWidth(1, 140); // ID Presensi
    sheetPresensi.setColumnWidth(2, 110); // Tanggal
    sheetPresensi.setColumnWidth(3, 140); // ID Pegawai
    sheetPresensi.setColumnWidth(4, 240); // Nama Pegawai
    sheetPresensi.setColumnWidth(5, 200); // Jabatan
    sheetPresensi.setColumnWidth(6, 130); // Jam Kehadiran
    sheetPresensi.setColumnWidth(7, 180); // Status Ketepatan
    sheetPresensi.setColumnWidth(8, 160); // Status: Sambut Pagi / Tidak Sambut Pagi
    sheetPresensi.setColumnWidth(9, 200); // Keterangan
    sheetPresensi.setColumnWidth(10, 160); // Timestamp
  }

  // 3. Setup Sheet Data_Users (Autentikasi Pengguna Terpusat)
  let sheetUsers = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(CONFIG.SHEET_USERS);
  }

  const headersUsers = [
    'ID', 'Username', 'Password', 'Nama Pengguna', 
    'Role', 'Role Name', 'Waktu Dibuat'
  ];

  if (sheetUsers.getLastRow() === 0) {
    sheetUsers.appendRow(headersUsers);
    const headerRange = sheetUsers.getRange(1, 1, 1, headersUsers.length);
    headerRange.setBackground('#1e293b')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    sheetUsers.setFrozenRows(1);
    sheetUsers.setColumnWidth(1, 120); // ID
    sheetUsers.setColumnWidth(2, 130); // Username
    sheetUsers.setColumnWidth(3, 140); // Password
    sheetUsers.setColumnWidth(4, 200); // Nama
    sheetUsers.setColumnWidth(5, 110); // Role
    sheetUsers.setColumnWidth(6, 160); // Role Name
    sheetUsers.setColumnWidth(7, 160); // Timestamp

    const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    sheetUsers.appendRow(["'usr-1", "admin", "admin123", "Administrator", "admin", "Super User (Admin)", nowStr]);
    sheetUsers.appendRow(["'usr-2", "petugas", "petugas123", "Petugas Piket", "petugas", "Petugas Sambut Pagi", nowStr]);
  }

  return {
    status: 'success',
    message: 'Database Google Sheets Sambut Pagi (Pegawai, Presensi, Users) berhasil diinisialisasi!'
  };
}

/**
 * 2. ROUTER HTTP GET
 * Digunakan untuk mengambil data (Fetch/Read) atau menyajikan antarmuka Web App
 */
function doGet(e) {
  try {
    const params = e ? e.parameter : {};

    // Jika diakses via browser atau disematkan ke Google Sites via "Sematkan dengan URL"
    // (tanpa parameter action atau format=json), sajikan HTML antarmuka web app dengan ALLOWALL
    if (!params.action && !params.tanggal && params.format !== 'json') {
      try {
        return HtmlService.createHtmlOutputFromFile('index')
          .setTitle('SMK Assalaam - Sambut Pagi Presensi')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
      } catch (errHtml) {
        // Jika file HTML index belum dibuat di Apps Script, lanjutkan merespons data JSON getAll
      }
    }

    const action = params.action || 'getAll';
    let result;

    switch (action) {
      case 'ping':
      case 'test':
        result = {
          status: 'success',
          message: 'Backend Google Apps Script Sambut Pagi Aktif & Terhubung!',
          timestamp: new Date().toISOString()
        };
        break;

      case 'getPegawai':
        result = {
          status: 'success',
          data: getDaftarPegawai()
        };
        break;

      case 'getPresensi':
        const targetTanggal = params.tanggal ? normalizeDateString(params.tanggal) : (params.all ? null : getFormattedDateToday());
        result = {
          status: 'success',
          tanggal: targetTanggal || 'all',
          data: getDaftarPresensi(targetTanggal)
        };
        break;

      case 'getUsers':
        result = {
          status: 'success',
          data: getDaftarUsers()
        };
        break;

      case 'getAll':
      default:
        const today = params.tanggal ? normalizeDateString(params.tanggal) : getFormattedDateToday();
        const fetchAllHistory = (params.all === 'true' || params.all === '1');
        result = {
          status: 'success',
          tanggal: today,
          pegawai: getDaftarPegawai(),
          // Ambil riwayat presensi default MAX_RECENT_DAYS agar performa kencang & hemat kuota Apps Script
          presensi: getDaftarPresensi(fetchAllHistory ? null : null, fetchAllHistory ? null : CONFIG.MAX_RECENT_DAYS),
          users: getDaftarUsers()
        };
        break;
    }

    return createJsonResponse(result);

  } catch (error) {
    return createJsonResponse({
      status: 'error',
      message: error.toString()
    });
  }
}

/**
 * 3. ROUTER HTTP OPTIONS
 * Mendukung preflight CORS untuk request lintas domain (Cross-Origin)
 */
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * 4. ROUTER HTTP POST
 * Digunakan untuk menulis data (Insert, Update, Delete, Batch Import)
 */
function doPost(e) {
  try {
    let payload = {};

    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        payload = e.parameter;
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;
    let result;

    switch (action) {
      // ===== DATA PEGAWAI =====
      case 'addPegawai':
        result = simpanPegawaiBaru(payload.data || payload);
        break;

      case 'updatePegawai':
        result = perbaruiPegawai(payload.data || payload);
        break;

      case 'deletePegawai':
        result = hapusPegawai(payload.id);
        break;

      case 'importPegawai':
        result = batchImportPegawai(payload.list || payload.data || []);
        break;

      case 'deleteAllPegawai':
        result = kosongkanSheetPegawai();
        break;

      // ===== PRESENSI SAMBUT PAGI =====
      case 'simpanPresensi':
      case 'updatePresensi':
        result = rekamPresensiPegawai(payload.data || payload);
        break;

      case 'deletePresensi':
        result = hapusPresensi(payload.id, payload.tanggal, payload.nama, payload.pegawai_id);
        break;

      case 'clearPresensiHariIni':
        const tgl = payload.tanggal || getFormattedDateToday();
        result = kosongkanPresensiTanggal(tgl);
        break;

      // ===== DATA PENGGUNA / USERS (MULTI-DEVICE SYNC) =====
      case 'getUsers':
        result = { status: 'success', data: getDaftarUsers() };
        break;

      case 'saveUser':
      case 'addUser':
      case 'updateUser':
        result = simpanUserBackend(payload.data || payload);
        break;

      case 'deleteUser':
        result = hapusUserBackend(payload.id);
        break;

      default:
        result = {
          status: 'error',
          message: 'Aksi tidak dikenal: ' + action
        };
        break;
    }

    return createJsonResponse(result);

  } catch (error) {
    return createJsonResponse({
      status: 'error',
      message: error.toString()
    });
  }
}

// ==============================================================================
// 5. LOGIKA OPERASI DATA PEGAWAI (No, Nama Guru Karyawan, Jenis Kelamin, Jabatan Tugas)
// ==============================================================================

function getDaftarPegawai() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_PEGAWAI);
  if (!sheet) {
    initialSetup();
    sheet = ss.getSheetByName(CONFIG.SHEET_PEGAWAI);
  }
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const pegawaiList = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[0] || row[2]) { // Ada ID atau Nama
      pegawaiList.push({
        id: String(row[0]).replace(/^'/, '').trim(),
        no: row[1] || (i + 1),
        nama_lengkap: String(row[2] || '').trim(),
        jenis_kelamin: String(row[3] || 'L').trim().toUpperCase(),
        jabatan: String(row[4] || 'Guru').trim(),
        created_at: row[5] ? String(row[5]) : ''
      });
    }
  }

  return pegawaiList;
}

function simpanPegawaiBaru(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PEGAWAI) || ss.insertSheet(CONFIG.SHEET_PEGAWAI);

    const rawId = p.id ? String(p.id).replace(/^'/, '').trim() : String(new Date().getTime());
    const nextNo = Math.max(1, sheet.getLastRow()); // Row 2 = No 1
    const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    const newRow = [
      "'" + rawId,
      p.no || nextNo,
      p.nama_lengkap || '',
      (p.jenis_kelamin || 'L').toUpperCase(),
      p.jabatan || 'Guru',
      nowStr
    ];

    sheet.appendRow(newRow);

    return {
      status: 'success',
      message: 'Pegawai berhasil disimpan ke Google Sheets!',
      data: { id: rawId, nama_lengkap: p.nama_lengkap }
    };
  } finally {
    lock.releaseLock();
  }
}

function perbaruiPegawai(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PEGAWAI);
    if (!sheet) return { status: 'error', message: 'Sheet pegawai tidak ditemukan' };

    const targetId = String(p.id).replace(/^'/, '').trim();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: 'Data pegawai kosong' };

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < ids.length; i++) {
      const rowId = String(ids[i][0]).replace(/^'/, '').trim();
      if (rowId === targetId) {
        const rowIndex = i + 2;
        // Update Nama Guru Karyawan, Jenis Kelamin, Jabatan Tugas (Kolom 3, 4, 5)
        sheet.getRange(rowIndex, 3, 1, 3).setValues([[
          p.nama_lengkap || '',
          (p.jenis_kelamin || 'L').toUpperCase(),
          p.jabatan || 'Guru'
        ]]);

        return {
          status: 'success',
          message: 'Data pegawai ' + p.nama_lengkap + ' berhasil diperbarui!'
        };
      }
    }

    return { status: 'error', message: 'Pegawai dengan ID ' + targetId + ' tidak ditemukan' };
  } finally {
    lock.releaseLock();
  }
}

function hapusPegawai(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PEGAWAI);
    if (!sheet) return { status: 'error', message: 'Sheet pegawai tidak ditemukan' };

    const targetId = String(id).replace(/^'/, '').trim();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: 'Tidak ada data pegawai' };

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < ids.length; i++) {
      const rowId = String(ids[i][0]).replace(/^'/, '').trim();
      if (rowId === targetId) {
        sheet.deleteRow(i + 2);
        return { status: 'success', message: 'Pegawai berhasil dihapus' };
      }
    }

    return { status: 'error', message: 'Pegawai tidak ditemukan' };
  } finally {
    lock.releaseLock();
  }
}

function batchImportPegawai(list) {
  if (!list || list.length === 0) {
    return { status: 'error', message: 'Daftar impor kosong' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PEGAWAI) || ss.insertSheet(CONFIG.SHEET_PEGAWAI);
    const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    const startNo = Math.max(1, sheet.getLastRow());
    const rowsToInsert = [];

    for (let i = 0; i < list.length; i++) {
      const item = list[i];
      const id = item.id ? String(item.id).replace(/^'/, '').trim() : String(new Date().getTime() + i);
      rowsToInsert.push([
        "'" + id,
        item.no || (startNo + i),
        item.nama_lengkap || '',
        (item.jenis_kelamin || 'L').toUpperCase(),
        item.jabatan || 'Guru',
        nowStr
      ]);
    }

    if (rowsToInsert.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToInsert.length, 6).setValues(rowsToInsert);
    }

    return {
      status: 'success',
      message: 'Berhasil mengimpor ' + rowsToInsert.length + ' data pegawai ke Google Sheets!'
    };
  } finally {
    lock.releaseLock();
  }
}

function kosongkanSheetPegawai() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PEGAWAI);
    if (!sheet) return { status: 'error', message: 'Sheet pegawai tidak ditemukan' };

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }

    return { status: 'success', message: 'Seluruh data pegawai berhasil dikosongkan' };
  } finally {
    lock.releaseLock();
  }
}

// ==============================================================================
// 6. LOGIKA OPERASI RIWAYAT PRESENSI
// ==============================================================================

function getDaftarPresensi(targetTanggal, limitDays) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_PRESENSI);
  if (!sheet) {
    initialSetup();
    sheet = ss.getSheetByName(CONFIG.SHEET_PRESENSI);
  }
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const presensiList = [];
  const normalizedTarget = targetTanggal ? normalizeDateString(targetTanggal) : null;

  // Batas tanggal jika limitDays ditentukan
  let minDateLimit = null;
  if (!normalizedTarget && limitDays && typeof limitDays === 'number') {
    const d = new Date();
    d.setDate(d.getDate() - limitDays);
    minDateLimit = Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    // Abaikan baris jika ID dan Nama kosong
    if (!row[0] && !row[3]) continue;

    // Normalisasi tanggal agar kompatibel format YYYY-MM-DD
    const rowTanggal = normalizeDateString(row[1]);

    if (normalizedTarget && rowTanggal !== normalizedTarget) continue;
    if (minDateLimit && rowTanggal && rowTanggal < minDateLimit) continue;

    let jamKehadiran = row[5];
    if (jamKehadiran instanceof Date) {
      jamKehadiran = Utilities.formatDate(jamKehadiran, CONFIG.TIMEZONE, 'HH.mm.ss') + ' WIB';
    } else {
      jamKehadiran = String(jamKehadiran || '');
    }

    // Standarisasi Status: 'Sambut Pagi', 'Tidak Sambut Pagi', atau 'OFF (Tidak Ada Jadwal)'
    const rawStatusPiket = String(row[7] || '').trim();
    let cleanStatusPiket = 'Sambut Pagi';
    if (rawStatusPiket.toLowerCase().includes('off') || rawStatusPiket.toLowerCase().includes('tidak ada jadwal')) {
      cleanStatusPiket = 'OFF (Tidak Ada Jadwal)';
    } else if (rawStatusPiket.toLowerCase().includes('tidak') || rawStatusPiket.toLowerCase().includes('bukan')) {
      cleanStatusPiket = 'Tidak Sambut Pagi';
    } else {
      cleanStatusPiket = 'Sambut Pagi';
    }

    presensiList.push({
      id: String(row[0]).replace(/^'/, '').trim(),
      tanggal: rowTanggal,
      pegawai_id: String(row[2]).replace(/^'/, '').trim(),
      nama: String(row[3] || '').trim(),
      jabatan: String(row[4] || '').trim(),
      jam: jamKehadiran,
      status: String(row[6] || '').trim(),
      piket: cleanStatusPiket,
      keterangan: String(row[8] || '').trim(),
      timestamp: String(row[9] || '').trim()
    });
  }

  // Urutkan presensi: Hadir paling AWAL berada di posisi ATAS (kronologis jam naik)
  presensiList.sort(function(a, b) {
    var secA = parseJamStringToSeconds(a.jam);
    var secB = parseJamStringToSeconds(b.jam);
    if (secA !== secB) return secA - secB;
    return String(a.id || '').localeCompare(String(b.id || ''));
  });

  return presensiList;
}

function rekamPresensiPegawai(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Cegah tabrakan data serentak (Race condition)
  } catch (e) {
    return {
      status: 'error',
      message: 'Server sedang sibuk memproses presensi lain, silakan tunggu beberapa saat.'
    };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_PRESENSI);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_PRESENSI);
    }

    // Pastikan baris judul (header) selalu ada
    if (sheet.getLastRow() === 0) {
      const headersPresensi = [
        'ID Presensi', 'Tanggal', 'ID Pegawai', 'Nama Pegawai', 
        'Jabatan', 'Jam Kehadiran', 'Status Ketepatan', 'Status', 
        'Keterangan', 'Timestamp Sistem'
      ];
      sheet.appendRow(headersPresensi);
      const headerRange = sheet.getRange(1, 1, 1, headersPresensi.length);
      headerRange.setBackground('#10b981')
                 .setFontColor('#ffffff')
                 .setFontWeight('bold')
                 .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }

    const tanggal = normalizeDateString(data.tanggal) || getFormattedDateToday();
    const pegawaiId = String(data.pegawai_id || '').replace(/^'/, '').trim();
    const now = new Date();
    const timestamp = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    // Bersihkan nilai status: 'Sambut Pagi', 'Tidak Sambut Pagi', atau 'OFF (Tidak Ada Jadwal)'
    const rawPiket = String(data.piket || '').trim();
    let statusPiket = 'Sambut Pagi';
    if (rawPiket.toLowerCase().includes('off') || rawPiket.toLowerCase().includes('tidak ada jadwal')) {
      statusPiket = 'OFF (Tidak Ada Jadwal)';
    } else if (rawPiket.toLowerCase().includes('tidak') || rawPiket.toLowerCase().includes('bukan')) {
      statusPiket = 'Tidak Sambut Pagi';
    } else {
      statusPiket = 'Sambut Pagi';
    }

    let jamKehadiran = data.jam;
    if (!jamKehadiran) {
      if (statusPiket === 'OFF (Tidak Ada Jadwal)') {
        jamKehadiran = '-';
      } else {
        jamKehadiran = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH.mm.ss') + ' WIB';
      }
    }

    let status = data.status;
    if (!status) {
      if (statusPiket === 'OFF (Tidak Ada Jadwal)') {
        status = 'OFF (Tidak Ada Jadwal)';
      } else {
        const hhmm = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm');
        if (hhmm <= CONFIG.BATAS_SAMBUT_PAGI) {
          status = 'Tepat Waktu (≤ 06.35)';
        } else if (hhmm <= CONFIG.BATAS_MASUK_KERJA) {
          status = 'Terlambat Sambut Pagi';
        } else {
          status = 'Terlambat (> 06.35)';
        }
      }
    }

    const lastRow = sheet.getLastRow();
    let existingRow = -1;

    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      const cleanTargetDataId = data.id ? String(data.id).replace(/^'/, '').trim() : '';

      for (let i = 0; i < rows.length; i++) {
        const rPresensiId = String(rows[i][0]).replace(/^'/, '').trim();
        const rTanggal = normalizeDateString(rows[i][1]);
        const rPegawaiId = String(rows[i][2]).replace(/^'/, '').trim();

        if ((cleanTargetDataId && rPresensiId === cleanTargetDataId) || (rTanggal === tanggal && rPegawaiId === pegawaiId)) {
          existingRow = i + 2;
          break;
        }
      }
    }

    if (existingRow !== -1) {
      sheet.getRange(existingRow, 6, 1, 5).setValues([[
        jamKehadiran,
        status,
        statusPiket,
        data.keterangan || '',
        timestamp
      ]]);

      return {
        status: 'success',
        action: 'updated',
        message: 'Jam kehadiran ' + data.nama + ' berhasil diperbarui ke ' + jamKehadiran,
        data: {
          id: String(data.id || '').replace(/^'/, '').trim(),
          tanggal: tanggal,
          pegawai_id: pegawaiId,
          nama: data.nama,
          jam: jamKehadiran,
          status: status,
          piket: statusPiket
        }
      };
    } else {
      const presensiId = data.id ? String(data.id).replace(/^'/, '').trim() : String(now.getTime());
      const newRow = [
        "'" + presensiId,
        "'" + tanggal,
        "'" + pegawaiId,
        data.nama || '',
        data.jabatan || '',
        jamKehadiran,
        status,
        statusPiket,
        data.keterangan || '',
        timestamp
      ];

      sheet.appendRow(newRow);

      return {
        status: 'success',
        action: 'inserted',
        message: 'Presensi ' + data.nama + ' berhasil dicatat!',
        data: {
          id: presensiId,
          tanggal: tanggal,
          pegawai_id: pegawaiId,
          nama: data.nama,
          jam: jamKehadiran,
          status: status,
          piket: statusPiket
        }
      };
    }
  } finally {
    lock.releaseLock();
  }
}

function hapusPresensi(id, tanggal, nama, pegawaiId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PRESENSI);
    if (!sheet) return { status: 'error', message: 'Sheet presensi tidak ditemukan' };

    const targetId = id ? String(id).replace(/^'/, '').trim() : '';
    const normTanggal = tanggal ? normalizeDateString(tanggal) : '';
    const targetNama = nama ? String(nama).trim().toLowerCase() : '';
    const targetPegawaiId = pegawaiId ? String(pegawaiId).replace(/^'/, '').trim() : '';

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: 'Data presensi kosong' };

    // Ambil kolom ID (1), Tanggal (2), ID Pegawai (3), Nama Pegawai (4)
    const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    let deletedCount = 0;

    // Loop mundur dari bawah ke atas agar penghapusan baris tidak merusak indeks baris berikutnya
    // Ini juga menjamin jika ada DUPLIKAT presensi, semuanya bersih terhapus
    for (let i = values.length - 1; i >= 0; i--) {
      const rowId = String(values[i][0]).replace(/^'/, '').trim();
      const rowTanggal = normalizeDateString(values[i][1]);
      const rowPegawaiId = String(values[i][2]).replace(/^'/, '').trim();
      const rowNama = String(values[i][3]).trim().toLowerCase();

      // 1. Cocokkan ID presensi secara eksak
      const matchId = targetId && (rowId === targetId);

      // 2. Cocokkan tanggal + pegawai_id atau nama (fallback yang kuat jika ID berbeda format)
      const matchFallback = normTanggal && (rowTanggal === normTanggal) &&
        ((targetPegawaiId && rowPegawaiId === targetPegawaiId) || (targetNama && rowNama === targetNama));

      if (matchId || matchFallback) {
        sheet.deleteRow(i + 2);
        deletedCount++;
      }
    }

    if (deletedCount > 0) {
      return { 
        status: 'success', 
        message: 'Berhasil menghapus ' + deletedCount + ' baris presensi dari Google Sheets' 
      };
    }

    return { status: 'error', message: 'Data presensi tidak ditemukan di Google Sheets' };
  } finally {
    lock.releaseLock();
  }
}

function kosongkanPresensiTanggal(targetTanggal) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PRESENSI);
    if (!sheet) return { status: 'error', message: 'Sheet presensi tidak ditemukan' };

    const normTarget = normalizeDateString(targetTanggal);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'success', message: 'Riwayat presensi sudah kosong' };

    const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let deletedCount = 0;

    for (let i = values.length - 1; i >= 0; i--) {
      const rTanggal = normalizeDateString(values[i][0]);
      if (rTanggal === normTarget) {
        sheet.deleteRow(i + 2);
        deletedCount++;
      }
    }

    return {
      status: 'success',
      message: 'Berhasil mengosongkan ' + deletedCount + ' riwayat presensi tanggal ' + normTarget
    };
  } finally {
    lock.releaseLock();
  }
}

// ==============================================================================
// 7. HELPER FUNCTIONS
// ==============================================================================

/**
 * Normalisasi format tanggal apapun dari Google Sheets (Date Object, string ISO, DD/MM/YYYY)
 * menjadi format standar "YYYY-MM-DD"
 */
function normalizeDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  const str = String(val).trim().replace(/^'/, '');
  // Cek apakah sudah format YYYY-MM-DD
  const matchIso = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (matchIso) {
    const y = matchIso[1];
    const m = matchIso[2].padStart(2, '0');
    const d = matchIso[3].padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  // Cek apakah format DD/MM/YYYY
  const matchDmy = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (matchDmy) {
    const d = matchDmy[1].padStart(2, '0');
    const m = matchDmy[2].padStart(2, '0');
    const y = matchDmy[3];
    return `${y}-${m}-${d}`;
  }
  // Coba parse objek Date jika memungkinkan
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    try {
      return Utilities.formatDate(parsed, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    } catch(e) {}
  }
  return str;
}

function getFormattedDateToday() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper: Konversi string jam (contoh: "06.30.00 WIB", "13.50.53 WIB", "13:50") ke total detik
 */
function parseJamStringToSeconds(jamStr) {
  if (!jamStr || String(jamStr).trim() === '-') return 999999;
  var clean = String(jamStr).trim();
  var match = clean.match(/(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?/);
  if (!match) return 999999;
  var h = parseInt(match[1], 10) || 0;
  var m = parseInt(match[2], 10) || 0;
  var s = parseInt(match[3] || '0', 10) || 0;
  return h * 3600 + m * 60 + s;
}

// ==============================================================================
// 8. LOGIKA OPERASI DATA PENGGUNA / USERS (Terpusat di Google Sheets)
// ==============================================================================

function getDaftarUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) {
    initialSetup();
    sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  }
  if (!sheet) {
    return [
      { id: 'usr-1', username: 'admin', password: 'admin123', nama: 'Administrator', role: 'admin', roleName: 'Super User (Admin)' },
      { id: 'usr-2', username: 'petugas', password: 'petugas123', nama: 'Petugas Piket', role: 'petugas', roleName: 'Petugas Sambut Pagi' }
    ];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const usersList = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[1] && row[2]) { // Ada username & password
      usersList.push({
        id: String(row[0]).replace(/^'/, '').trim() || ('usr-' + (i + 1)),
        username: String(row[1]).trim().toLowerCase(),
        password: String(row[2]).trim(),
        nama: String(row[3] || row[1]).trim(),
        role: String(row[4] || 'petugas').trim().toLowerCase(),
        roleName: String(row[5] || (row[4] === 'admin' ? 'Super User (Admin)' : 'Petugas Sambut Pagi')).trim()
      });
    }
  }

  return usersList;
}

function simpanUserBackend(u) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    if (!sheet) {
      initialSetup();
      sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    }

    const userId = u.id ? String(u.id).replace(/^'/, '').trim() : ('usr-' + new Date().getTime());
    const username = String(u.username || '').trim().toLowerCase();
    const password = String(u.password || '').trim();
    const nama = String(u.nama || username).trim();
    const role = String(u.role || 'petugas').trim().toLowerCase();
    const roleName = String(u.roleName || (role === 'admin' ? 'Super User (Admin)' : 'Petugas Sambut Pagi')).trim();
    const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    if (!username || !password) {
      return { status: 'error', message: 'Username dan password wajib diisi!' };
    }

    const lastRow = sheet.getLastRow();
    let existingRow = -1;

    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < rows.length; i++) {
        const rId = String(rows[i][0]).replace(/^'/, '').trim();
        const rUsername = String(rows[i][1]).trim().toLowerCase();
        if (rId === userId || rUsername === username) {
          existingRow = i + 2;
          break;
        }
      }
    }

    if (existingRow !== -1) {
      // Update
      sheet.getRange(existingRow, 2, 1, 5).setValues([[
        username,
        password,
        nama,
        role,
        roleName
      ]]);
      return {
        status: 'success',
        action: 'updated',
        message: 'Pengguna ' + nama + ' berhasil diperbarui di Google Sheets!'
      };
    } else {
      // Insert baru
      sheet.appendRow([
        "'" + userId,
        username,
        password,
        nama,
        role,
        roleName,
        nowStr
      ]);
      return {
        status: 'success',
        action: 'inserted',
        message: 'Pengguna baru ' + nama + ' berhasil disimpan di Google Sheets!'
      };
    }
  } finally {
    lock.releaseLock();
  }
}

function hapusUserBackend(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    if (!sheet) return { status: 'error', message: 'Sheet pengguna tidak ditemukan' };

    const targetId = String(id || '').replace(/^'/, '').trim();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: 'Tidak ada data pengguna' };

    const ids = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      const rId = String(ids[i][0]).replace(/^'/, '').trim();
      const rUsername = String(ids[i][1]).trim().toLowerCase();

      // Lindungi admin utama agar tidak bisa dihapus
      if (rUsername === 'admin' && (targetId === rId || targetId === 'admin')) {
        return { status: 'error', message: 'Akun Administrator utama tidak boleh dihapus.' };
      }

      if (rId === targetId || rUsername === targetId.toLowerCase()) {
        sheet.deleteRow(i + 2);
        return { status: 'success', message: 'Pengguna berhasil dihapus dari Google Sheets' };
      }
    }

    return { status: 'error', message: 'Pengguna tidak ditemukan' };
  } finally {
    lock.releaseLock();
  }
}
