/**
 * =========================================================================
 *  PANDUAN UPDATE GOOGLE APPS SCRIPT
 * =========================================================================
 * 
 * Silakan copy dan paste baris-baris kode di bawah ini ke dalam project
 * Google Apps Script (Code.gs) Anda, lalu lakukan DEPLOY ulang sebagai Web App.
 * Pastikan untuk mempublikasikan "New Version" (versi baru) agar 
 * perubahan URL tidak terjadi namun script-nya update.
 * 
 */

// -------------------------------------------------------------------------
// 1. TAMBAHKAN LOGIKA INI DI DALAM FUNGSI doPost(e)
// Cari fungsi doPost(e) Anda yang sudah ada, lalu di dalam percabangan
// 'action', tambahkan kode untuk 'absen_bulk' dan 'update_absen_status':
// -------------------------------------------------------------------------

/*
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({result: 'error', message: 'Invalid JSON'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = data.action;
  
  // ---> TAMBAHKAN BLOK KODE INI: <---
  if (action === 'absen_bulk') {
    return handleAbsenBulk(data);
  }
  if (action === 'update_absen_status') {
    return updateAbsenStatus(data);
  }
  // ----------------------------------

  // (Kode Anda sebelumnya untuk action 'simpan_cuti', 'save_setup', dll)
  // ...
}
*/

// -------------------------------------------------------------------------
// 2. TAMBAHKAN LOGIKA INI DI DALAM FUNGSI doGet(e)
// Cari fungsi doGet(e) Anda yang sudah ada, lalu di dalam percabangan
// 'action', tambahkan kode untuk 'rekapAbsen':
// -------------------------------------------------------------------------

/*
function doGet(e) {
  const action = e.parameter.action;

  // ---> TAMBAHKAN BLOK KODE INI: <---
  if (action === 'rekapAbsen') {
    return getRekapAbsen();
  }
  // ----------------------------------

  // (Kode Anda sebelumnya untuk action 'init', 'riwayat', 'booked', dll)
  // ...
}
*/

// -------------------------------------------------------------------------
// 3. TAMBAHKAN FUNGSI-FUNGSI BARU INI DI BAGIAN PALING BAWAH FILE ANDA
// (Tinggal copy-paste fungsi di bawah ini)
// -------------------------------------------------------------------------

function getOrCreateAbsenSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Absen Briefing");
  
  // Jika sheet belum ada, buat baru dan setel header
  if (!sheet) {
    sheet = ss.insertSheet("Absen Briefing");
    sheet.appendRow(["Waktu (Timestamp)", "ID Karyawan", "Nama Karyawan", "Bagian/Divisi", "Status"]);
    // Mempercantik header
    const headerRange = sheet.getRange("A1:E1");
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f3f4f6");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 5);
  }
  
  return sheet;
}

function handleAbsenBulk(data) {
  try {
    const sheet = getOrCreateAbsenSheet();
    const records = data.records; // Array dari data absensi massal
    
    if (!records || !Array.isArray(records)) {
      throw new Error("Data records tidak valid");
    }

    // Persiapkan data array 2D untuk insert sekaligus agar lebih efisien
    const rows = [];
    records.forEach(function(rec) {
      rows.push([
        rec.waktu,
        rec.idKaryawan,
        rec.nama,
        rec.bagian,
        rec.status
      ]);
    });

    if (rows.length > 0) {
      // Dapatkan range baris baru yang akan ditambahkan
      const startRow = sheet.getLastRow() + 1;
      const targetRange = sheet.getRange(startRow, 1, rows.length, rows[0].length);
      // Set values sekaligus
      targetRange.setValues(rows);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success', 
      message: 'Berhasil merekam rekap absen massal'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', 
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getRekapAbsen() {
  try {
    const sheet = getOrCreateAbsenSheet();
    const lastRow = sheet.getLastRow();
    
    // Jika belum ada data absensi
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        data: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Ambil semua data (kecuali header)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5);
    const values = dataRange.getValues();
    
    const rekapData = [];
    
    // Urutkan dari yang terbaru (reverse loop)
    for (let i = values.length - 1; i >= 0; i--) {
      rekapData.push({
        waktu: values[i][0],
        idKaryawan: values[i][1],
        nama: values[i][2],
        bagian: values[i][3],
        status: values[i][4]
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      data: rekapData
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', 
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------------------------
// FUNGSI: Update status absen individual (dipanggil dari dashboard admin)
// Mencari baris berdasarkan idKaryawan + waktu (timestamp ISO), lalu
// mengupdate kolom Status (kolom ke-5).
// -------------------------------------------------------------------------
function updateAbsenStatus(data) {
  try {
    const sheet = getOrCreateAbsenSheet();
    const idKaryawan = String(data.idKaryawan);
    const waktuTarget = String(data.waktu);
    const newStatus = data.newStatus;

    if (!idKaryawan || !waktuTarget || !newStatus) {
      throw new Error('Parameter tidak lengkap: idKaryawan, waktu, newStatus diperlukan.');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      throw new Error('Sheet absen masih kosong, tidak ada data untuk diupdate.');
    }

    // Kolom: A=Waktu(1), B=ID(2), C=Nama(3), D=Bagian(4), E=Status(5)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5);
    const values = dataRange.getValues();

    let found = false;
    for (let i = 0; i < values.length; i++) {
      const rowWaktu = String(values[i][0]);
      const rowId    = String(values[i][1]);

      // Cocokkan berdasarkan ID karyawan dan waktu
      if (rowId === idKaryawan && rowWaktu === waktuTarget) {
        // Update kolom Status (kolom ke-5, index sheet = 5)
        sheet.getRange(i + 2, 5).setValue(newStatus);
        found = true;
        break;
      }
    }

    if (!found) {
      // Fallback: cari hanya berdasarkan ID karyawan dan ambil baris terbaru hari yang sama
      const targetDate = waktuTarget.substring(0, 10); // YYYY-MM-DD
      for (let i = values.length - 1; i >= 0; i--) {
        const rowWaktu = String(values[i][0]);
        const rowId    = String(values[i][1]);
        const rowDate  = rowWaktu.length >= 10 ? rowWaktu.substring(0, 10) : '';

        if (rowId === idKaryawan && rowDate === targetDate) {
          sheet.getRange(i + 2, 5).setValue(newStatus);
          found = true;
          break;
        }
      }
    }

    if (!found) {
      throw new Error('Data absen tidak ditemukan untuk karyawan dan tanggal tersebut.');
    }

    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      message: 'Status absen berhasil diperbarui.'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
