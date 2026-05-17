// ================================================================
// SISTEM CUTI KARYAWAN — Google Apps Script (v6 - ADMIN DASHBOARD PRO)
// ================================================================

const SPREADSHEET_ID  = "16FYB1fjjT7dQcOYMCw-Wict43FOEhBlIqTqyHRoIeS4";
const SHEET_DATA      = "Data Cuti";
const SHEET_REKAP     = "Rekap Harian";
const SHEET_KARYAWAN  = "Karyawan";
const SHEET_SETUP     = "Setup";

function getColMap(sheet) {
  if (!sheet || sheet.getLastColumn() === 0) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}

// ================================================================
// ROUTER (GET)
// ================================================================
function doGet(e) {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const action = e.parameter.action || 'init';

  if (action === 'init') {
    return jsonResponse(getInitData(ss));
  }
  if (action === 'booked') {
    return jsonResponse({ booked: getBookedDates(ss, e.parameter.bagian || '', e.parameter.idKaryawan || '') });
  }
  if (action === 'riwayat') {
    return jsonResponse({ riwayat: getRiwayat(ss, e.parameter.idKaryawan || '') });
  }
  if (action === 'dashboard_data') {
    return jsonResponse(getDashboardData(ss));
  }

  return jsonResponse({ error: 'Action tidak valid.' });
}

function getInitData(ss) {
  const shKaryawan = ss.getSheetByName(SHEET_KARYAWAN);
  const mapK = getColMap(shKaryawan);
  const dataK = shKaryawan.getDataRange().getValues();
  const karyawan = [];
  
  for (let i = 1; i < dataK.length; i++) {
    const status = (mapK['Status'] !== undefined) ? dataK[i][mapK['Status']] : 'Aktif';
    if (String(status).toLowerCase() === 'nonaktif') continue; // Hide from public UI
    
    const id = dataK[i][mapK['ID Karyawan']] || dataK[i][mapK['Nama Karyawan']] || '';
    const nama = dataK[i][mapK['Nama Karyawan']] || '';
    const bagian = dataK[i][mapK['Bagian']] || '';
    
    if (id || nama) karyawan.push({ id, nama, bagian });
  }

  const shSetup = ss.getSheetByName(SHEET_SETUP);
  let batasWaktu = "";
  if (shSetup) {
    const dataS = shSetup.getDataRange().getValues();
    for (let i = 1; i < dataS.length; i++) {
      if (dataS[i][0] === "Batas Akhir Penginputan") batasWaktu = dataS[i][1];
    }
  }

  return { karyawan, batasWaktu };
}

// ================================================================
// ROUTER (POST)
// ================================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'simpan_cuti';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (action === 'simpan_cuti') {
      return handleSimpanCuti(ss, payload);
    } else if (action === 'toggle_cuti') {
      return handleToggleCuti(ss, payload);
    } else if (action === 'save_karyawan') {
      return handleSaveKaryawan(ss, payload);
    } else if (action === 'delete_karyawan') {
      return handleDeleteKaryawan(ss, payload);
    } else if (action === 'save_setup') {
      return handleSaveSetup(ss, payload);
    }

    return jsonResponse({ result: 'error', message: 'Unknown Action' });
  } catch (err) {
    return jsonResponse({ result: 'error', message: 'Error: ' + err.message });
  }
}
function handleSaveSetup(ss, payload) {
  const { batasWaktu } = payload;
  const shSetup = ss.getSheetByName(SHEET_SETUP) || ss.insertSheet(SHEET_SETUP);
  const mapS = getColMap(shSetup);
  const data = shSetup.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][mapS['Pengaturan']] === 'Batas Akhir Penginputan') {
      rowIdx = i + 1;
      break;
    }
  }
  const nilaiCol = (mapS['Nilai'] !== undefined) ? mapS['Nilai'] + 1 : 2;
  if (rowIdx > 0) {
    shSetup.getRange(rowIdx, nilaiCol).setValue(batasWaktu || '');
  } else {
    shSetup.appendRow(['Batas Akhir Penginputan', batasWaktu || '']);
  }
  return jsonResponse({ result: 'success' });
}

function handleSimpanCuti(ss, payload) {
  const { idKaryawan, nama, bagian, tanggal, isAdmin } = payload;

  if (!idKaryawan || !nama || !bagian || !Array.isArray(tanggal)) {
    return jsonResponse({ result: 'error', message: 'Data tidak lengkap.' });
  }

  // Validasi Maks 3 Hari per BULAN
  const datesByMonth = {};
  for (let d of tanggal) {
    const ym = d.substring(0, 7);
    datesByMonth[ym] = (datesByMonth[ym] || 0) + 1;
    if (datesByMonth[ym] > 3) {
      return jsonResponse({ result: 'error', message: `Maksimal 3 hari cuti per bulan! Anda melebihi kuota di bulan ${ym}.` });
    }
  }
  
  // Validasi Batas Waktu (Dilewati jika yang edit adalah Admin via Dashboard)
  if (!isAdmin) {
    const { batasWaktu } = getInitData(ss);
    if (batasWaktu && new Date() > new Date(batasWaktu)) {
      return jsonResponse({ result: 'error', message: 'Batas waktu penginputan telah berakhir. Data dikunci.' });
    }
  }

  const sheet = ss.getSheetByName(SHEET_DATA);
  const map = getColMap(sheet);
  if (map['ID Karyawan'] === undefined) {
    return jsonResponse({ result: 'error', message: 'Struktur sheet Data Cuti salah. Lakukan Setup ulang.' });
  }

  // Validasi Konflik Divisi
  const booked = getBookedDates(ss, bagian, idKaryawan);
  const conflicts = tanggal.filter(d => booked.includes(d));
  if (conflicts.length > 0) {
    return jsonResponse({ result: 'error', message: `Tanggal ${conflicts.join(', ')} sudah diambil rekan di divisi ${bagian}.` });
  }

  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map['ID Karyawan']]) === String(idKaryawan)) {
      rowIdx = i + 1;
      break;
    }
  }

  const rawStr = tanggal.sort().join(',');
  
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, map['Nama Karyawan'] + 1).setValue(nama); 
    sheet.getRange(rowIdx, map['Bagian'] + 1).setValue(bagian);
    sheet.getRange(rowIdx, map['Tanggal (raw)'] + 1).setNumberFormat('@').setValue(rawStr);
    sheet.getRange(rowIdx, map['Total Hari'] + 1).setValue(tanggal.length);
    sheet.getRange(rowIdx, map['Update Terakhir'] + 1).setValue(new Date());
  } else {
    if (tanggal.length > 0) {
      const newRow = new Array(Object.keys(map).length).fill('');
      newRow[map['Timestamp']] = new Date();
      newRow[map['ID Karyawan']] = idKaryawan;
      newRow[map['Nama Karyawan']] = nama;
      newRow[map['Bagian']] = bagian;
      newRow[map['Tanggal (raw)']] = rawStr;
      newRow[map['Total Hari']] = tanggal.length;
      newRow[map['Update Terakhir']] = new Date();
      sheet.appendRow(newRow);
    }
  }

  buildRekap(ss);
  return jsonResponse({ result: 'success', tanggal });
}

function handleToggleCuti(ss, payload) {
  const { idKaryawan, tanggalToggle, isAdmin } = payload;
  if (!isAdmin || !idKaryawan || !tanggalToggle) return jsonResponse({ result: 'error', message: 'Invalid payload' });

  const sheet = ss.getSheetByName(SHEET_DATA);
  const map = getColMap(sheet);
  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  let tRaw = '';
  let nama = '';
  let bagian = '';

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map['ID Karyawan']]) === String(idKaryawan)) {
      rowIdx = i + 1;
      const val = data[i][map['Tanggal (raw)']];
      if (val instanceof Date) {
        tRaw = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        tRaw = String(val || '');
      }
      nama = data[i][map['Nama Karyawan']];
      bagian = data[i][map['Bagian']];
      break;
    }
  }

  // Jika belum punya baris, kita tidak bisa toggle dari Gantt chart dengan mudah tanpa nama/bagian, tapi dari Gantt pasti sudah ada nama/bagian.
  // Wait, di Dashboard, dbKaryawan punya data lengkap. Mari asumsikan payload juga mengirim nama dan bagian.
  nama = payload.nama || nama;
  bagian = payload.bagian || bagian;

  let arr = tRaw ? tRaw.split(',').map(d => d.trim()).filter(Boolean) : [];
  let added = false;
  
  if (arr.includes(tanggalToggle)) {
    // Remove
    arr = arr.filter(d => d !== tanggalToggle);
  } else {
    // Add (Check kuota bulan itu)
    const ym = tanggalToggle.substring(0, 7);
    const countBulanIni = arr.filter(d => d.startsWith(ym)).length;
    if (countBulanIni >= 3) {
      return jsonResponse({ result: 'error', message: `Kuota bulan ${ym} penuh (Maks 3 hari).` });
    }
    
    // Check conflicts
    const booked = getBookedDates(ss, bagian, idKaryawan);
    if (booked.includes(tanggalToggle)) {
      return jsonResponse({ result: 'error', message: 'Tanggal sudah diambil rekan di divisi yang sama.' });
    }

    arr.push(tanggalToggle);
    added = true;
  }

  const rawStr = arr.sort().join(',');

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, map['Tanggal (raw)'] + 1).setNumberFormat('@').setValue(rawStr);
    sheet.getRange(rowIdx, map['Total Hari'] + 1).setValue(arr.length);
    sheet.getRange(rowIdx, map['Update Terakhir'] + 1).setValue(new Date());
  } else {
    const newRow = new Array(Object.keys(map).length).fill('');
    newRow[map['Timestamp']] = new Date();
    newRow[map['ID Karyawan']] = idKaryawan;
    newRow[map['Nama Karyawan']] = nama;
    newRow[map['Bagian']] = bagian;
    newRow[map['Tanggal (raw)']] = rawStr;
    newRow[map['Total Hari']] = arr.length;
    newRow[map['Update Terakhir']] = new Date();
    sheet.appendRow(newRow);
  }

  buildRekap(ss);
  return jsonResponse({ result: 'success', added });
}

function handleSaveKaryawan(ss, payload) {
  const { idKaryawan, nama, bagian, status, isNew } = payload;
  const sheet = ss.getSheetByName(SHEET_KARYAWAN);
  
  // Pastikan kolom Status ada
  let map = getColMap(sheet);
  if (map['Status'] === undefined) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Status');
    map = getColMap(sheet);
  }

  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  
  if (!isNew) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][map['ID Karyawan']]) === String(idKaryawan)) {
        rowIdx = i + 1;
        break;
      }
    }
  }

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, map['ID Karyawan'] + 1).setNumberFormat('@');
    sheet.getRange(rowIdx, map['Nama Karyawan'] + 1).setValue(nama);
    sheet.getRange(rowIdx, map['Bagian'] + 1).setValue(bagian);
    sheet.getRange(rowIdx, map['Status'] + 1).setValue(status);
  } else {
    // Generate ID otomatis: cari nomor tertinggi dan tambah 1 agar unik
    let newId = idKaryawan && idKaryawan.trim();
    if (!newId) {
      let maxNum = 0;
      for (let i = 1; i < data.length; i++) {
        const existId = String(data[i][map['ID Karyawan']] || '');
        const match = existId.match(/K-(\d+)/);
        if (match) maxNum = Math.max(maxNum, parseInt(match[1]));
      }
      newId = 'K-' + String(maxNum + 1).padStart(3, '0');
    }
    const newRow = new Array(Object.keys(map).length).fill('');
    newRow[map['ID Karyawan']] = newId;
    newRow[map['Nama Karyawan']] = nama;
    newRow[map['Bagian']] = bagian;
    newRow[map['Status']] = status || 'Aktif';
    sheet.appendRow(newRow);
    // Ensure ID is text
    sheet.getRange(sheet.getLastRow(), map['ID Karyawan'] + 1).setNumberFormat('@');
  }
  
  // Update juga nama & bagian di sheet Data Cuti agar sinkron
  const shData = ss.getSheetByName(SHEET_DATA);
  const mapD = getColMap(shData);
  if(mapD['ID Karyawan'] !== undefined && !isNew) {
     const dataC = shData.getDataRange().getValues();
     for(let i=1; i<dataC.length; i++) {
       if(String(dataC[i][mapD['ID Karyawan']]) === String(idKaryawan)) {
         shData.getRange(i+1, mapD['Nama Karyawan']+1).setValue(nama);
         shData.getRange(i+1, mapD['Bagian']+1).setValue(bagian);
       }
     }
  }

  return jsonResponse({ result: 'success' });
}

function handleDeleteKaryawan(ss, payload) {
  const { idKaryawan } = payload;
  
  // Hapus dari Karyawan
  const shKaryawan = ss.getSheetByName(SHEET_KARYAWAN);
  let map = getColMap(shKaryawan);
  let data = shKaryawan.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][map['ID Karyawan']]) === String(idKaryawan)) {
      shKaryawan.deleteRow(i + 1);
      break;
    }
  }

  // Hapus dari Data Cuti (Riwayat)
  const shData = ss.getSheetByName(SHEET_DATA);
  map = getColMap(shData);
  if (map['ID Karyawan'] !== undefined) {
    data = shData.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][map['ID Karyawan']]) === String(idKaryawan)) {
        shData.deleteRow(i + 1);
        break; // Setiap ID cuma punya 1 baris di Data Cuti (struktur sekarang)
      }
    }
  }

  // Rekap ulang
  buildRekap(ss);

  return jsonResponse({ result: 'success' });
}

// ================================================================
// HELPERS
// ================================================================
function getBookedDates(ss, bagian, excludeId) {
  const sheet = ss.getSheetByName(SHEET_DATA);
  const map = getColMap(sheet);
  const data = sheet.getDataRange().getValues();
  const booked = [];
  if (map['Bagian'] === undefined) return booked;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map['Bagian']]).toLowerCase() === String(bagian).toLowerCase() && String(data[i][map['ID Karyawan']]) !== String(excludeId)) {
      const dates = String(data[i][map['Tanggal (raw)']] || '').split(',').map(d => d.trim()).filter(Boolean);
      dates.forEach(d => { if (!booked.includes(d)) booked.push(d); });
    }
  }
  return booked;
}

function getRiwayat(ss, idKaryawan) {
  const sheet = ss.getSheetByName(SHEET_DATA);
  const map = getColMap(sheet);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map['ID Karyawan']]) === String(idKaryawan)) {
      const val = data[i][map['Tanggal (raw)']];
      let tRaw = '';
      if (val instanceof Date) {
        tRaw = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        tRaw = String(val || '');
      }
      const tanggal = tRaw.split(',').map(d => d.trim()).filter(Boolean);
      return { ada: true, tanggal };
    }
  }
  return { ada: false, tanggal: [] };
}

function getDashboardData(ss) {
  const shRekap = ss.getSheetByName(SHEET_REKAP);
  const mapR = getColMap(shRekap);
  const data = shRekap.getDataRange().getValues();
  
  const tz = Session.getScriptTimeZone();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][mapR['ID Karyawan']]) {
      // Google Sheets bisa mengkonversi string tanggal ke Date object secara otomatis
      // Kita harus format ulang agar selalu jadi string YYYY-MM-DD
      const rawTgl = data[i][mapR['Tanggal Cuti']];
      let tanggalStr = '';
      if (rawTgl instanceof Date) {
        tanggalStr = Utilities.formatDate(rawTgl, tz, 'yyyy-MM-dd');
      } else {
        tanggalStr = String(rawTgl || '').trim();
      }

      const rawBulan = data[i][mapR['Bulan']];
      let bulanStr = '';
      if (rawBulan instanceof Date) {
        bulanStr = Utilities.formatDate(rawBulan, tz, 'MM');
      } else {
        bulanStr = String(rawBulan || '').padStart(2, '0');
      }

      const rawTahun = data[i][mapR['Tahun']];
      let tahunStr = '';
      if (rawTahun instanceof Date) {
        tahunStr = Utilities.formatDate(rawTahun, tz, 'yyyy');
      } else {
        tahunStr = String(rawTahun || '');
      }

      result.push({
        idKaryawan: String(data[i][mapR['ID Karyawan']]),
        nama: data[i][mapR['Nama Karyawan']],
        bagian: data[i][mapR['Bagian']],
        tanggal: tanggalStr,
        bulan: bulanStr,
        tahun: tahunStr
      });
    }
  }

  // Kirim daftar SEMUA karyawan (termasuk Nonaktif) untuk Dashboard
  const shKaryawan = ss.getSheetByName(SHEET_KARYAWAN);
  const mapK = getColMap(shKaryawan);
  const dataK = shKaryawan.getDataRange().getValues();
  const allKaryawan = [];
  for (let i = 1; i < dataK.length; i++) {
    const id = dataK[i][mapK['ID Karyawan']] || dataK[i][mapK['Nama Karyawan']] || '';
    if (!id) continue;
    allKaryawan.push({
      id: id,
      nama: dataK[i][mapK['Nama Karyawan']] || '',
      bagian: dataK[i][mapK['Bagian']] || '',
      status: (mapK['Status'] !== undefined) ? dataK[i][mapK['Status']] : 'Aktif'
    });
  }
  
  return { result: 'success', data: result, karyawan: allKaryawan };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function formatDateSafe(val, tz) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val || '').trim();
}

// ================================================================
// REKAP HARIAN
// ================================================================
function buildRekap(ss) {
  const datSheet = ss.getSheetByName(SHEET_DATA);
  const mapDat = getColMap(datSheet);
  let rekSheet = ss.getSheetByName(SHEET_REKAP);
  if (!rekSheet) rekSheet = ss.insertSheet(SHEET_REKAP);

  rekSheet.clearContents();
  const headers = ['ID Karyawan', 'Nama Karyawan', 'Bagian', 'Tanggal Cuti', 'Bulan', 'Tahun'];
  rekSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  rekSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  rekSheet.setFrozenRows(1);

  const data = datSheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const id = data[i][mapDat['ID Karyawan']];
    const nama = data[i][mapDat['Nama Karyawan']];
    const bagian = data[i][mapDat['Bagian']];
    const tRaw = String(data[i][mapDat['Tanggal (raw)']] || '');
    
    if (id && tRaw) {
      tRaw.split(',').forEach(d => {
        const t = d.trim();
        if (t) {
          const parts = t.split('-');
          rows.push([String(id), nama, bagian, t, parts[1], parts[0]]);
        }
      });
    }
  }

  if (rows.length > 0) {
    const range = rekSheet.getRange(2, 1, rows.length, headers.length);
    range.setValues(rows);
    // Paksa kolom Tanggal Cuti (kolom 4) sebagai teks agar Google Sheets tidak auto-convert ke Date
    rekSheet.getRange(2, 4, rows.length, 1).setNumberFormat('@STRING@');
    // Paksa kolom Bulan dan Tahun (kolom 5 & 6) tetap angka/teks
    rekSheet.getRange(2, 5, rows.length, 2).setNumberFormat('@STRING@');
  }
}

// ================================================================
// SETUP AWAL
// ================================================================
function setupSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function makeSheet(name, headers, sampleData) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (sampleData && sh.getLastRow() === 1) {
      sh.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    }
    return sh;
  }

  makeSheet(SHEET_DATA, ['Timestamp', 'ID Karyawan', 'Nama Karyawan', 'Bagian', 'Tanggal (raw)', 'Total Hari', 'Update Terakhir']);
  makeSheet(SHEET_KARYAWAN, ['ID Karyawan', 'Nama Karyawan', 'Bagian', 'Status'], [
    ['K-001', 'Fulan', 'Kasir', 'Aktif'], 
    ['K-002', 'Dwi', 'Gudang', 'Aktif']
  ]);
  
  const d = new Date(); d.setDate(d.getDate() + 7);
  makeSheet(SHEET_SETUP, ['Pengaturan', 'Nilai'], [
    ['Batas Akhir Penginputan', Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss")]
  ]);
  
  buildRekap(ss);
  Logger.log('Setup v6 PRO berhasil! Dashboard dilengkapi CRUD.');
}