function debugSheetConnection() {
  try {
    var SPREADSHEET_ID = '1Zd4plVMj7Z_UczDz8enSMYKI3AgD505AuuQdhNdPGqo';
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pelaksanaan');
    
    if (!sheet) {
      return "❌ ERROR: Sheet 'Pelaksanaan' TIDAK DITEMUKAN!\nPastikan nama tab di Google Sheet persis 'Pelaksanaan'.";
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      return "⚠️ WARNING: Sheet 'Pelaksanaan' DITEMUKAN tapi KOSONG (Hanya Header atau Kosong Total).";
    }

    // Ambil Header
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Ambil Sample Data Row 2
    var row2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    
    // Cek Value Spesifik
    var sapEvent = row2[1]; // Kolom B
    var items = row2.map((val, idx) => `[${idx}] ${headers[idx]}: "${val}"`).join('\n');

    return "✅ KONEKSI SUKSES!\nSheet: Pelaksanaan\nTotal Baris: " + lastRow + "\n\nSample Baris 2:\n" + items;

  } catch (e) {
    return "❌ EXCEPTION: " + e.message + "\nStack: " + e.stack;
  }
}
