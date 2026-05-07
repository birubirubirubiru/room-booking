// ============================================================
//  ROOM BOOKING — Google Apps Script Backend
//  Salin seluruh kode ini ke Apps Script di Google Spreadsheet
// ============================================================

// ── KONFIGURASI ─────────────────────────────────────────────
const SHEET_NAME   = 'Bookings';   // nama sheet untuk data booking
const COL_ID       = 1;  // A
const COL_NAME     = 2;  // B
const COL_DIV      = 3;  // C
const COL_ROOM     = 4;  // D
const COL_DATE     = 5;  // E
const COL_START    = 6;  // F
const COL_DUR      = 7;  // G
const COL_DESC     = 8;  // H
const COL_STATUS   = 9;  // I
const COL_CREATED  = 10; // J

// ── SETUP: buat sheet & header jika belum ada ────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Tulis header jika baris 1 kosong
  if (sheet.getRange('A1').getValue() === '') {
    const headers = ['ID', 'Nama', 'Divisi', 'Ruangan', 'Tanggal', 'Jam Mulai', 'Durasi (jam)', 'Agenda', 'Status', 'Dibuat Pada'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1D9E75');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);

    // Lebar kolom
    sheet.setColumnWidth(1, 80);   // ID
    sheet.setColumnWidth(2, 160);  // Nama
    sheet.setColumnWidth(3, 120);  // Divisi
    sheet.setColumnWidth(4, 120);  // Ruangan
    sheet.setColumnWidth(5, 110);  // Tanggal
    sheet.setColumnWidth(6, 90);   // Jam Mulai
    sheet.setColumnWidth(7, 100);  // Durasi
    sheet.setColumnWidth(8, 250);  // Agenda
    sheet.setColumnWidth(9, 100);  // Status
    sheet.setColumnWidth(10, 160); // Dibuat Pada

    sheet.setFrozenRows(1);

    SpreadsheetApp.getUi().alert('Sheet "Bookings" berhasil dibuat!');
  }

  return sheet;
}

// ── HANDLE REQUEST HTTP ──────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'getBookings') {
    return handleGetBookings();
  }
  if (action === 'cancelBooking') {
    return handleCancelBooking(e.parameter.id);
  }

  return jsonResponse({ status: 'error', message: 'Action tidak dikenal' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || '';

    if (action === 'addBooking') {
      return handleAddBooking(data);
    }
    if (action === 'cancelBooking') {
      return handleCancelBooking(data.id);
    }

    return jsonResponse({ status: 'error', message: 'Action tidak dikenal' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ── GET: ambil semua booking ─────────────────────────────────
function handleGetBookings() {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return jsonResponse({ status: 'ok', bookings: [] });
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const bookings = data
      .filter(row => row[0] !== '') // skip baris kosong
      .map(row => ({
        id:     row[COL_ID - 1],
        name:   row[COL_NAME - 1],
        div:    row[COL_DIV - 1],
        room:   row[COL_ROOM - 1],
        date:   formatDateToISO(row[COL_DATE - 1]),
        start:  Number(row[COL_START - 1]),
        dur:    Number(row[COL_DUR - 1]),
        desc:   row[COL_DESC - 1],
        status: row[COL_STATUS - 1],
      }));

    return jsonResponse({ status: 'ok', bookings });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ── POST: tambah booking baru ────────────────────────────────
function handleAddBooking(data) {
  try {
    const sheet = getOrCreateSheet();

    // Validasi input
    if (!data.name || !data.room || !data.date || data.start === undefined || !data.dur) {
      return jsonResponse({ status: 'error', message: 'Data tidak lengkap' });
    }

    // Cek konflik waktu
    const conflict = checkConflict(sheet, data.room, data.date, Number(data.start), Number(data.dur));
    if (conflict) {
      return jsonResponse({ status: 'conflict', message: `Waktu bentrok dengan booking ${conflict}` });
    }

    // Generate ID unik
    const id = 'BK-' + new Date().getTime();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

    // Tulis ke sheet
    sheet.appendRow([
      id,
      data.name,
      data.div || 'Staff',
      data.room,
      data.date,
      Number(data.start),
      Number(data.dur),
      data.desc || '',
      'ok',
      now
    ]);

    // Warnai baris baru berdasarkan ruangan
    const newRow = sheet.getLastRow();
    colorRow(sheet, newRow, data.room);

    return jsonResponse({ status: 'ok', id, message: 'Booking berhasil disimpan' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ── POST/GET: batalkan booking ───────────────────────────────
function handleCancelBooking(id) {
  try {
    if (!id) return jsonResponse({ status: 'error', message: 'ID tidak diberikan' });

    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'error', message: 'Tidak ada data' });

    const ids = sheet.getRange(2, COL_ID, lastRow - 1, 1).getValues();
    let found = false;

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        const rowNum = i + 2;
        sheet.getRange(rowNum, COL_STATUS).setValue('cancelled');
        // Abu-abu untuk baris yang dibatalkan
        sheet.getRange(rowNum, 1, 1, 10).setBackground('#F0F0F0').setFontColor('#999999');
        found = true;
        break;
      }
    }

    if (!found) return jsonResponse({ status: 'error', message: 'Booking tidak ditemukan' });
    return jsonResponse({ status: 'ok', message: 'Booking berhasil dibatalkan' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ── HELPER: cek konflik waktu ────────────────────────────────
function checkConflict(sheet, room, date, start, dur) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (const row of data) {
    if (row[COL_STATUS - 1] !== 'ok') continue;
    if (row[COL_ROOM - 1] !== room) continue;

    const rowDate = formatDateToISO(row[COL_DATE - 1]);
    if (rowDate !== date) continue;

    const rowStart = Number(row[COL_START - 1]);
    const rowDur   = Number(row[COL_DUR - 1]);

    // Overlap check
    if (!(start >= rowStart + rowDur || start + dur <= rowStart)) {
      return row[COL_NAME - 1]; // kembalikan nama yang konflik
    }
  }
  return null;
}

// ── HELPER: warnai baris berdasarkan ruangan ─────────────────
function colorRow(sheet, rowNum, room) {
  const colors = { 'Mawar': '#FDE8E8', 'Melati': '#FCE4EF', 'Kenanga': '#E0F7EE' };
  const bg = colors[room] || '#FFFFFF';
  sheet.getRange(rowNum, 1, 1, 10).setBackground(bg);
}

// ── HELPER: konversi Date object ke string YYYY-MM-DD ────────
function formatDateToISO(val) {
  if (!val) return '';
  if (typeof val === 'string') return val;
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(val);
}

// ── HELPER: get atau buat sheet ───────────────────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  return sheet;
}

// ── HELPER: response JSON dengan CORS ────────────────────────
function jsonResponse(obj) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── MENU DI SPREADSHEET ───────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Room Booking')
    .addItem('Setup Sheet', 'setupSheet')
    .addItem('Lihat Statistik', 'showStats')
    .addSeparator()
    .addItem('Hapus Semua Data (Hati-hati!)', 'clearAllData')
    .addToUi();
}

// ── STATISTIK ─────────────────────────────────────────────────
function showStats() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Belum ada data booking.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const active = data.filter(r => r[COL_STATUS - 1] === 'ok');
  const cancelled = data.filter(r => r[COL_STATUS - 1] === 'cancelled');

  const byRoom = {};
  active.forEach(r => {
    const room = r[COL_ROOM - 1];
    byRoom[room] = (byRoom[room] || 0) + 1;
  });

  let msg = `📊 STATISTIK BOOKING\n\n`;
  msg += `Total Booking Aktif : ${active.length}\n`;
  msg += `Total Dibatalkan    : ${cancelled.length}\n\n`;
  msg += `Per Ruangan:\n`;
  Object.entries(byRoom).forEach(([room, count]) => {
    msg += `  • ${room}: ${count} booking\n`;
  });

  SpreadsheetApp.getUi().alert(msg);
}

// ── HAPUS SEMUA DATA ──────────────────────────────────────────
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Konfirmasi',
    'Yakin ingin menghapus SEMUA data booking? Tindakan ini tidak bisa dibatalkan.',
    ui.ButtonSet.YES_NO
  );

  if (confirm === ui.Button.YES) {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 10).clearContent().setBackground('#FFFFFF').setFontColor('#000000');
    }
    ui.alert('Semua data berhasil dihapus.');
  }
}
