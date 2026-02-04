/* --- FUNGSI DEBUGGING (Jalankan ini di Editor, bukan di Browser) --- */
function debugSpreadsheet() {
  var sheetName = 'Perencanaan'; // Pastikan ini SAMA PERSIS dengan nama Tab di bawah
  var ss = SpreadsheetApp.openById('1Zd4plVMj7Z_UczDz8enSMYKI3AgD505AuuQdhNdPGqo'); // ID Spreadsheet Anda
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("❌ ERROR: Sheet dengan nama '" + sheetName + "' TIDAK DITEMUKAN!");
    var allSheets = ss.getSheets();
    Logger.log("   Daftar sheet yang ada: " + allSheets.map(s => s.getName()).join(', '));
    return;
  }

  Logger.log("✅ Sheet '" + sheetName + "' ditemukan.");
  
  var lastRow = sheet.getLastRow();
  Logger.log("📊 Jumlah Baris Terisi (LastRow): " + lastRow);

  if (lastRow < 2) {
    Logger.log("⚠️ PERINGATAN: Data tampaknya kosong (hanya header atau kosong total).");
    return;
  }

  // Cek Baris Kedua (Data Pertama)
  var firstDataRow = sheet.getRange(2, 1, 1, 16).getValues()[0];
  Logger.log("🔎 Isi Baris ke-2 (Index 1):");
  Logger.log("   - Kolom A (No): " + firstDataRow[0]);
  Logger.log("   - Kolom D (Item ID Cert): " + firstDataRow[3]);
  Logger.log("   - Kolom L (Item ID Lat): " + firstDataRow[11]);

  // Simulasi Logika getData
  if (firstDataRow[3] && firstDataRow[3].toString() !== "") {
    Logger.log("✅ Logika CERT: Data akan terbaca.");
  } else {
    Logger.log("❌ Logika CERT: Data TIDAK terbaca karena Kolom D kosong.");
  }

  // Simulasi Logika getLATData
  if (firstDataRow[11] && firstDataRow[11].toString() !== "") {
    Logger.log("✅ Logika LAT: Data akan terbaca.");
  } else {
    Logger.log("❌ Logika LAT: Data TIDAK terbaca karena Kolom L kosong.");
  }
}
