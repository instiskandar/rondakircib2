/**
 * Aplikasi Presensi & Jadwal Ronda Malam - Full CRUD Engine
 * Terhubung langsung dengan Spreadsheet ID: 1VYPt7lEDhCr-fnriViSzBhDoaFGgzUeOIxYPaDb9mes
 */

var SPREADSHEET_ID = "1VYPt7lEDhCr-fnriViSzBhDoaFGgzUeOIxYPaDb9mes";

function getDbSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Aplikasi Presensi & Jadwal Ronda Malam - Desa Sukarukun')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. READ: AMBIL SEMUA DATA (SINGLE BATCH QUERY)
function getInitialData() {
  var ss = getDbSpreadsheet();
  setupDatabaseSheets(ss);

  return {
    reguList: fetchSheetObjects(ss.getSheetByName("Regu")),
    wargaList: fetchSheetObjects(ss.getSheetByName("Warga")),
    absensiList: fetchSheetObjects(ss.getSheetByName("Absensi")),
    sopList: fetchSheetObjects(ss.getSheetByName("SOP")),
    usersList: fetchSheetObjects(ss.getSheetByName("Users"))
  };
}

function fetchSheetObjects(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row.join('').trim()) continue;

    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var rawKey = headers[j];
      var key = normalizeHeaderKey(rawKey);
      var val = row[j];

      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      obj[key] = val;
    }

    if (obj.tanggal && typeof obj.tanggal === 'string' && obj.tanggal.indexOf(',') !== -1) {
      obj.tanggalArr = obj.tanggal.split(',').map(function(t) { return t.trim(); });
    } else if (obj.tanggal) {
      obj.tanggalArr = [String(obj.tanggal)];
    }

    result.push(obj);
  }
  return result;
}

function normalizeHeaderKey(header) {
  var map = {
    "ID": "id",
    "Nama Regu": "nama",
    "Ketua": "ketua",
    "Blok Ketua": "blokKetua",
    "Tanggal Ronda": "tanggal",
    "Nama": "nama",
    "Blok": "blok",
    "Regu ID": "reguId",
    "Keterangan": "ket",
    "Warga ID": "wargaId",
    "Nama Warga": "namaWarga",
    "Tanggal": "tanggal",
    "Status": "status",
    "Denda": "denda",
    "Judul": "judul",
    "Isi": "isi",
    "Username": "username",
    "Nama Lengkap": "nama",
    "Role": "role",
    "Jabatan / Role": "role",
    "Password": "password"
  };
  return map[header] || header.toLowerCase();
}

function setupDatabaseSheets(ss) {
  if (!ss) return;
  var sheets = ["Regu", "Warga", "Absensi", "SOP", "Users"];
  sheets.forEach(function(s) {
    if (!ss.getSheetByName(s)) {
      var newSheet = ss.insertSheet(s);
      initSheetDefaults(newSheet, s);
    }
  });
}

function initSheetDefaults(sheet, name) {
  if (name === "Regu") {
    sheet.appendRow(["ID", "Nama Regu", "Ketua", "Blok Ketua", "Tanggal Ronda"]);
    sheet.appendRow([1, "Regu 1", "RIZKI AMIN", "Blok F7/6", "31 Oktober 2026, 28 November 2026, 26 Desember 2026"]);
    sheet.appendRow([2, "Regu 2", "AHMAD SUBAGJA", "Blok F8/2", "07 November 2026, 05 Desember 2026"]);
    sheet.appendRow([3, "Regu 3", "BAMBANG HERMANTO", "Blok G1/10", "14 November 2026, 12 Desember 2026"]);
    sheet.appendRow([4, "Regu 4", "DEDI KURNIAWAN", "Blok G3/04", "21 November 2026, 19 Desember 2026"]);
  } else if (name === "Warga") {
    sheet.appendRow(["ID", "Nama", "Blok", "Regu ID", "Tanggal Ronda", "Keterangan"]);
    sheet.appendRow([1, "Rizki Amin", "F7/6", 1, "31 Okt 2026, 28 Nov 2026", "Ketua Regu 1"]);
    sheet.appendRow([2, "Budi Santoso", "F7/1", 1, "31 Okt 2026, 28 Nov 2026", "Anggota"]);
  } else if (name === "Absensi") {
    sheet.appendRow(["ID", "Warga ID", "Nama Warga", "Regu ID", "Tanggal", "Status", "Denda"]);
    sheet.appendRow([101, 1, "Rizki Amin", 1, "2026-10-31", "Hadir", 0]);
    sheet.appendRow([102, 2, "Budi Santoso", 1, "2026-10-31", "Alpa", 20000]);
  } else if (name === "SOP") {
    sheet.appendRow(["ID", "Judul", "Isi"]);
    sheet.appendRow([1, "Ketentuan Denda Ronda", "Warga yang tidak hadir ronda (Alpa) dikenakan kas Rp 20.000."]);
  } else if (name === "Users") {
    sheet.appendRow(["Username", "Nama", "Role", "Password"]);
    sheet.appendRow(["admin", "Administrator RT", "Ketua RT / Admin", "admin123"]);
  }

  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  header.setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
  sheet.setFrozenRows(1);
}

// --- CRUD: REGU ---
function apiSaveRegu(payload) {
  var sheet = getDbSpreadsheet().getSheetByName("Regu");
  var data = sheet.getDataRange().getValues();
  var tgl = Array.isArray(payload.tanggal) ? payload.tanggal.join(", ") : payload.tanggal;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(payload.id)) {
      sheet.getRange(i + 1, 3, 1, 3).setValues([[payload.ketua, payload.blokKetua, tgl]]);
      return { success: true, message: "Regu " + payload.id + " berhasil disimpan!" };
    }
  }
  sheet.appendRow([payload.id, "Regu " + payload.id, payload.ketua, payload.blokKetua, tgl]);
  return { success: true, message: "Regu baru berhasil ditambahkan!" };
}

// --- CRUD: WARGA ---
function apiSaveWarga(payload) {
  var sheet = getDbSpreadsheet().getSheetByName("Warga");
  var data = sheet.getDataRange().getValues();

  if (payload.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.id)) {
        sheet.getRange(i + 1, 2, 1, 5).setValues([[payload.nama, payload.blok, Number(payload.reguId), payload.tanggal, payload.ket]]);
        return { success: true, message: "Data warga berhasil diperbarui!" };
      }
    }
  }
  var newId = new Date().getTime();
  sheet.appendRow([newId, payload.nama, payload.blok, Number(payload.reguId), payload.tanggal, payload.ket]);
  return { success: true, message: "Warga baru berhasil ditambahkan!", id: newId };
}

function apiDeleteWarga(id) {
  var ss = getDbSpreadsheet();
  var sheetWarga = ss.getSheetByName("Warga");
  var dataWarga = sheetWarga.getDataRange().getValues();
  var targetWargaNama = "";

  // 1. Cari & Hapus Warga dari Sheet Warga
  for (var i = 1; i < dataWarga.length; i++) {
    if (String(dataWarga[i][0]) === String(id)) {
      targetWargaNama = dataWarga[i][1]; // Simpan nama warga
      sheetWarga.deleteRow(i + 1);
      break;
    }
  }

  if (!targetWargaNama) {
    return { success: false, message: "Data warga tidak ditemukan!" };
  }

  // 2. Otomatis Hapus Presensi / Daftar Hadir Ronda Terkait di Sheet Absensi
  var sheetAbsensi = ss.getSheetByName("Absensi");
  if (sheetAbsensi) {
    var dataAbsensi = sheetAbsensi.getDataRange().getValues();
    // Hapus dari baris paling bawah ke atas agar indeks baris tidak bergeser
    for (var j = dataAbsensi.length - 1; j >= 1; j--) {
      var absensiWargaId = String(dataAbsensi[j][1]);
      var absensiNamaWarga = String(dataAbsensi[j][2]).toLowerCase().trim();

      if (absensiWargaId === String(id) || absensiNamaWarga === String(targetWargaNama).toLowerCase().trim()) {
        sheetAbsensi.deleteRow(j + 1);
      }
    }
  }

  return { 
    success: true, 
    message: "Data warga '" + targetWargaNama + "' dan seluruh riwayat presensinya berhasil dihapus secara otomatis!" 
  };
}

// --- CRUD: ABSENSI ---
function apiSaveAbsensi(payload) {
  var sheet = getDbSpreadsheet().getSheetByName("Absensi");
  var data = sheet.getDataRange().getValues();

  if (payload.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.id)) {
        sheet.getRange(i + 1, 2, 1, 6).setValues([[payload.wargaId, payload.namaWarga, Number(payload.reguId), payload.tanggal, payload.status, Number(payload.denda)]]);
        return { success: true, message: "Presensi berhasil diperbarui!" };
      }
    }
  }
  var newId = new Date().getTime();
  sheet.appendRow([newId, payload.wargaId, payload.namaWarga, Number(payload.reguId), payload.tanggal, payload.status, Number(payload.denda)]);
  return { success: true, message: "Presensi berhasil dicatat!", id: newId };
}

function apiDeleteAbsensi(id) {
  var sheet = getDbSpreadsheet().getSheetByName("Absensi");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Presensi berhasil dihapus!" };
    }
  }
  return { success: false, message: "Data presensi tidak ditemukan!" };
}

// --- CRUD: SOP & INFORMASI ---
function apiSaveSop(payload) {
  var sheet = getDbSpreadsheet().getSheetByName("SOP");
  var data = sheet.getDataRange().getValues();
  if (payload.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.id)) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[payload.judul, payload.isi]]);
        return { success: true, message: "Informasi SOP berhasil diperbarui!" };
      }
    }
  }
  var newId = new Date().getTime();
  sheet.appendRow([newId, payload.judul, payload.isi]);
  return { success: true, message: "Informasi SOP berhasil ditambahkan!", id: newId };
}

function apiDeleteSop(id) {
  var sheet = getDbSpreadsheet().getSheetByName("SOP");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Informasi SOP berhasil dihapus!" };
    }
  }
  return { success: false, message: "SOP tidak ditemukan!" };
}

// --- CRUD: USERS / PENGURUS ---
function apiSaveUser(payload) {
  var sheet = getDbSpreadsheet().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.username).trim()) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[payload.nama, payload.role, payload.password]]);
      return { success: true, message: "Data pengurus diperbarui!" };
    }
  }
  sheet.appendRow([payload.username, payload.nama, payload.role, payload.password]);
  return { success: true, message: "Pengurus baru berhasil ditambahkan!" };
}

function apiDeleteUser(username) {
  var sheet = getDbSpreadsheet().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(username).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Pengurus berhasil dihapus!" };
    }
  }
  return { success: false, message: "User tidak ditemukan!" };
}

function apiAuthenticate(username, password) {
  var sheet = getDbSpreadsheet().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(username).trim() && String(data[i][3]).trim() === String(password).trim()) {
      return { success: true, user: { username: data[i][0], nama: data[i][1], role: data[i][2] } };
    }
  }
  return { success: false, message: "Username atau Password salah!" };
}