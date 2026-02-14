/**
 * CODE.GS - UPDATED HEADER VERSION
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistem Informasi Sertifikasi')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

var SPREADSHEET_ID = '1Zd4plVMj7Z_UczDz8enSMYKI3AgD505AuuQdhNdPGqo'; 
var MAIN_SHEET_NAME = 'Perencanaan';

function connect() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  return sheet;
}

/* --- 1. DATA SERTIFIKASI (KIRI) --- */
function getData() {
  try {
    var sheet = connect();
    if (!sheet) return []; // Safety check
    
    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        // SKIP HEADER: Jika kolom A bernilai "NO" (case-insensitive)
        if (String(r[0]).toUpperCase() === "NO") continue;
        
        // VALIDASI: Skip jika baris dianggap kosong (tidak ada SAP dan Nama)
        // Kita tidak wajibkan kolom A (No) terisi, agar data tetap muncul meski user lupa isi nomor
        if (!r[1] && !r[2]) continue;

        try {
            var certItemId = r[3]; // Kolom D
            
            // Ambil data jika ada ID atau Nama
            if ((certItemId && String(certItemId).trim() !== "") || (r[1] && r[2])) {
                // Jika ID (r[0]) kosong, gunakan index loop sebagai fallback ID sementara
                var id = r[0] ? String(r[0]) : "ROW_" + i;

                data.push({
                    id: id,          
                    sap: cleanString(r[1]), 
                    nama: String(r[2]),        
                    itemId: String(r[3]),      
                    judul: String(r[4]),       
                    periode: safeParseDate(r[5]), 
                    jumlah: String(r[6]),           // JUMLAH ANGGARAN
                    statusAnggaran: String(r[7]),   // TERSEDIA/TIDAK
                    mandatory: String(r[8]),        // MANDATORY DAN REGULASI
                    resiko: String(r[9]),           // RESIKO
                    type: 'cert'
                });
            }
        } catch (rowErr) {
            Logger.log("Error processing CERT row " + i + ": " + rowErr);
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getData: ' + e.message);
    return []; // Return empty array to keep frontend running
  }
}

/* --- 2. DATA LAT (KANAN - Kolom L ke kanan) --- */
function getLATData() {
  try {
    var sheet = connect();
    if (!sheet) return [];

    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (String(r[0]).toUpperCase() === "NO") continue;
        if (!r[1] && !r[2]) continue; // Skip empty rows

        try {
            // Cek Kolom L (Index 11) - Item ID LAT
            var latItemId = r[11];
            
            // Allow entry if valid LAT item OR if basic data exists (SAP/Nama)
            if ((latItemId && String(latItemId).trim() !== "") || (r[1] && r[2] && r[11])) {
                 var id = r[0] ? String(r[0]) : "ROW_" + i;
                 
                 data.push({
                    id: id + "_LAT", 
                    originalId: id,
                    sap: cleanString(r[1]),
                    nama: String(r[2]),
                    itemId: String(r[11]),     
                    judul: String(r[12]),      
                    instruktur: String(r[13]), 
                    periode: safeParseDate(r[14]),
                    resiko: String(r[15]),     
                    type: 'lat'
                });
            }
        } catch (rowErr) {
             Logger.log("Error processing LAT row " + i + ": " + rowErr);
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getLATData: ' + e.message);
    return [];
  }
}

// HELPER
function cleanString(val) {
  if (!val) return "";
  return String(val).trim().toUpperCase(); 
}

// SAFE PARSE DATE - Handles Indonesian format and returns YYYY-MM-DD
function safeParseDate(dateVal) {
  try {
      if (!dateVal) return "";
      
      // 1. Jika object Date (dari Excel date cell)
      if (Object.prototype.toString.call(dateVal) === '[object Date]') {
        var yyyy = dateVal.getFullYear();
        var mm = String(dateVal.getMonth() + 1).padStart(2, '0');
        var dd = String(dateVal.getDate()).padStart(2, '0');
        return yyyy + "-" + mm + "-" + dd;
      }
      
      var str = String(dateVal).trim();

      // 2. Handle Format "Bulan Tahun" (Contoh: "Maret 2026")
      var monthMap = {
        'JANUARI': '01', 'FEBRUARI': '02', 'MARET': '03', 'APRIL': '04', 'MEI': '05', 'JUNI': '06',
        'JULI': '07', 'AGUSTUS': '08', 'SEPTEMBER': '09', 'OKTOBER': '10', 'NOVEMBER': '11', 'DESEMBER': '12',
        'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'JUN': '06', 'JUL': '07', 'AGU': '08', 'SEP': '09', 'OKT': '10', 'NOV': '11', 'DES': '12'
      };
      
      // Cek apakah format "NamaBulan Tahun"
      var parts = str.split(' ');
      if (parts.length === 2) {
        var mName = parts[0].toUpperCase();
        var yName = parts[1];
        if (monthMap[mName] && !isNaN(yName)) {
           return yName + "-" + monthMap[mName] + "-01";
        }
      }
      
      // 3. Handle Format "D/M/YYYY" atau "M/D/YYYY" (Excel text format kadang begini)
      // Asumsi default Spreadsheet Indonesia: DD/MM/YYYY
      if (str.includes('/')) {
         var p = str.split('/');
         if (p.length === 3) {
            // Cek mana yang tahun (biasanya 4 digit)
            if (p[2].length === 4) return p[2] + "-" + String(p[1]).padStart(2,'0') + "-" + String(p[0]).padStart(2,'0');
            // Jika format english M/D/Y
            if (p[2].length === 2 && p[0].length === 4) return p[0] + "-" + String(p[1]).padStart(2,'0') + "-" + String(p[2]).padStart(2,'0'); 
         }
      }

      return str; 
  } catch (e) {
      return String(dateVal);
  }
}

function parseDate(d) { return safeParseDate(d); }

/* --- 3. CRUD (Update Mapping Save) --- */
/* --- 3. CRUD (DIPERBAIKI AGAR NOMOR BERURUTAN) --- */

// Helper untuk mendapatkan nomor urut selanjutnya
function getNextId(sheet) {
  var lastRow = sheet.getLastRow();
  
  // Jika baris hanya 1 (hanya header), mulai dari 1
  if (lastRow <= 1) return 1;

  // Ambil nilai dari kolom A baris terakhir
  var lastVal = sheet.getRange(lastRow, 1).getValue();

  // Pastikan nilainya angka, jika tidak (misal error), gunakan nomor baris
  var nextNum = parseInt(lastVal);
  if (isNaN(nextNum)) {
    return lastRow; // Fallback jika data berantakan
  }
  
  return nextNum + 1; // Nomor terakhir + 1
}

function addData(formObject) {
  var sheet = connect();
  
  // UBAH DISINI: Pakai getNextId bukan Date().getTime()
  var id = getNextId(sheet); 
  
  var newRow = [
      id, 
      formObject.sap, 
      formObject.nama,
      formObject.itemId, 
      formObject.judul, 
      formObject.periode, 
      formObject.jumlah,         
      formObject.statusAnggaran, 
      formObject.mandatory,      
      formObject.resiko,         
      "", "", "", "", "", "" 
  ];
  sheet.appendRow(newRow);
  return { success: true };
}

function addLATData(formObject) {
    var sheet = connect();
    
    // UBAH DISINI JUGA
    var id = getNextId(sheet);

    var newRow = [
        id, formObject.sap, formObject.nama,
        "", "", "", "", "", "", "", 
        "", 
        formObject.itemId, formObject.judul, formObject.instruktur, 
        formObject.periode, formObject.resiko
    ];
    sheet.appendRow(newRow);
    return { success: true };
}


/* --- 4. DATA PELAKSANAAN (UPDATED: SESUAI USER HEADERS) --- */
/**
 * Membaca data dari sheet Pelaksanaan.
 * Kolom: NO, SAP, Start, End, Bulan, Tahun, Item ID, Sap Instruktur, Nama Instruktur, 
 * Course Title, SAP, Nama Partisipan, Room, Pesona, Kel, Departemen, Unit Kerja, 
 * Jumlah Hadir, Count Pelatihan, Durasi, Kehadiran, Durasi Peserta, Durasi Instruktur
 */
function getRealizationData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pelaksanaan'); 
    
    if (!sheet) {
      Logger.log('Sheet Pelaksanaan tidak ditemukan');
      return [];
    }

    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    // Skip header row, mulai dari index 1
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        
        try {
            // Ambil NO dari kolom A (index 0)
            var no = r[0] ? String(r[0]) : "";
            
            // Skip jika ini adalah header "NO" (just in case)
            if (no.trim().toUpperCase() === "NO") continue;

            // VALIDASI: Skip jika baris dianggap kosong total (Cek SAP, Nama, Judul, SAP Peserta, Nama Peserta)
            // r[1]=SAP Event, r[9]=Course Title, r[6]=Item ID, r[10]=SAP Peserta, r[11]=Nama Peserta
            if (!r[1] && !r[9] && !r[6] && !r[10] && !r[11]) continue;
            
            // Fallback ID jika kosong
            if (!no || no.trim() === "") no = "REAL_" + i;

            // Safe Date Parsing
            var dateStart = safeParseDate(r[2]);
            var dateEnd = safeParseDate(r[3]);
            
            // Robust Year Extraction
            // 1. Cek kolom Tahun (Index 5)
            var rawTahun = r[5];
            var tahun = "";
            
            if (rawTahun) {
               if (Object.prototype.toString.call(rawTahun) === '[object Date]') {
                  tahun = String(rawTahun.getFullYear());
               } else {
                  var strTahun = String(rawTahun).trim();
                  var match = strTahun.match(/20\d{2}/);
                  if (match) tahun = match[0];
                  else tahun = strTahun;
               }
            }
            
            
            // 2. Fallback: Jika kolom Tahun kosong, ambil dari Start Date (Index 2)
            if ((!tahun || tahun === "") && dateStart) {
                var d = new Date(dateStart);
                if (!isNaN(d.getTime())) {
                    tahun = String(d.getFullYear());
                }
            }

            // 3. Fallback: Ambil dari End Date (Index 3)
            if ((!tahun || tahun === "") && dateEnd) {
                 var d = new Date(dateEnd);
                 if (!isNaN(d.getTime())) {
                     tahun = String(d.getFullYear());
                 }
            }

            // 4. Last Resort: "Uncategorized" atau empty string (biar masuk card "Semua Data")
            if (!tahun) tahun = ""; 

            // --- CRITICAL FIX FOR FRONTEND GROUPING ---
            // Frontend groups by SAP. If Participant SAP (Col K / Index 10) is missing, 
            // we MUST provide a fallback, otherwise it might be grouped under "undefined" or lost.
            
            // PRIORITAS SAP: 1. SAP Peserta (K) -> 2. SAP Event (B) -> 3. "NO_SAP"
            var finalSap = r[10] ? String(r[10]) : (r[1] ? String(r[1]) : "NO_SAP");
            
            // PRIORITAS NAMA: 1. Nama Peserta (L) -> 2. Course Title (J) -> 3. Nama Instruktur -> 4. "No Name"
            var finalNama = r[11] ? String(r[11]) : (r[9] ? String(r[9]) : (r[8] ? String(r[8]) : "No Name"));

            data.push({
                id: no,                                         
                sapEvent: r[1] ? String(r[1]) : "",            
                sapStart: dateStart,         
                end: dateEnd,              
                bulan: r[4] ? String(r[4]) : "",               
                tahun: tahun,                                   
                itemId: r[6] ? String(r[6]) : "",              
                sapInstruktur: r[7] ? String(r[7]) : "",       
                namaInstruktur: r[8] ? String(r[8]) : "",      
                courseTitle: r[9] ? String(r[9]) : "",         
                sapPeserta: r[10] ? String(r[10]) : "",        
                namaPeserta: r[11] ? String(r[11]) : "",       
                room: r[12] ? String(r[12]) : "",              
                
                // MAPPING BARU SESUAI GAMBAR USER
                pesona: r[13] ? String(r[13]) : "",          // Ex: Presensi
                kel: r[14] ? String(r[14]) : "",             // Ex: Ket
                
                departemen: r[15] ? String(r[15]) : "",        
                unitKerja: r[16] ? String(r[16]) : "",         
                jumlahHadir: r[17] != null ? String(r[17]) : "",       
                countPelatihan: r[18] != null ? String(r[18]) : "",    
                durasi: r[19] != null ? String(r[19]) : "",            
                kehadiran: r[20] != null ? String(r[20]) : "",         
                durasiPeserta: r[21] != null ? String(r[21]) : "",   
                durasiInstruktur: r[22] != null ? String(r[22]) : "",    
                
                // Compatibility Fields (untuk frontend existing agar tidak error)
                sap: finalSap,       // CRITICAL: Used for grouping in renderRealizationList
                nama: finalNama      // CRITICAL: Used for grouping name
            });
        } catch (errRow) {
            Logger.log("Error processing row " + i + ": " + errRow.message);
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getRealizationData: ' + e.message);
    return []; 
  }
}

/* ─────────────────────────────────────────────────────────────────────────────
 * REALIZATION DATA CRUD OPERATIONS
 * ───────────────────────────────────────────────────────────────────────────── */

/**
 * Add new realization data to Pelaksanaan sheet
 */
function addRealizationData(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pelaksanaan');
    
    if (!sheet) {
      return { success: false, error: 'Sheet Pelaksanaan not found' };
    }

    // Generate new ID
    var nextId = getNextId(sheet);
    
    // Map formData to row array matching column order
    var newRow = [
      nextId,                            // A: NO
      formData.sapEvent || '',          // B: SAP Event
      formData.sapStart || '',          // C: Start Date
      formData.end || '',               // D: End Date
      formData.bulan || '',             // E: Bulan
      formData.tahun || '',             // F: Tahun
      formData.itemId || '',            // G: Item ID
      formData.sapInstruktur || '',     // H: SAP Instruktur
      formData.namaInstruktur || '',    // I: Nama Instruktur
      formData.courseTitle || '',       // J: Course Title
      formData.sapPeserta || '',        // K: SAP Peserta
      formData.namaPeserta || '',       // L: Nama Peserta
      formData.room || '',              // M: Room
      '',                                // N: Pesona (not in form)
      '',                                // O: Kel (not in form)
      formData.departemen || '',        // P: Departemen
      formData.unitKerja || '',         // Q: Unit Kerja
      formData.jumlahHadir || '',       // R: Jumlah Hadir
      '',                                // S: Count Pelatihan (not in form)
      formData.durasi || '',            // T: Durasi
      formData.kehadiran || '',         // U: Kehadiran
      '',                                // V: Durasi Peserta (not in form)
      ''                                 // W: Durasi Instruktur (not in form)
    ];
    
    sheet.appendRow(newRow);
    
    // Return updated data
    var updatedData = getRealizationData();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addRealizationData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Update existing realization data
 */
function updateRealizationData(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pelaksanaan');
    
    if (!sheet) {
      return { success: false, error: 'Sheet Pelaksanaan not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // Find row by ID (column A)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1; // Row number (1-indexed)
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found with ID: ' + formData.id };
    }
    
    // Update row with new data
    sheet.getRange(rowIndex, 2).setValue(formData.sapEvent || '');      // B
    sheet.getRange(rowIndex, 3).setValue(formData.sapStart || '');      // C
    sheet.getRange(rowIndex, 4).setValue(formData.end || '');           // D
    sheet.getRange(rowIndex, 5).setValue(formData.bulan || '');         // E
    sheet.getRange(rowIndex, 6).setValue(formData.tahun || '');         // F
    sheet.getRange(rowIndex, 7).setValue(formData.itemId || '');        // G
    sheet.getRange(rowIndex, 8).setValue(formData.sapInstruktur || ''); // H
    sheet.getRange(rowIndex, 9).setValue(formData.namaInstruktur || ''); // I
    sheet.getRange(rowIndex, 10).setValue(formData.courseTitle || '');  // J
    sheet.getRange(rowIndex, 11).setValue(formData.sapPeserta || '');   // K
    sheet.getRange(rowIndex, 12).setValue(formData.namaPeserta || '');  // L
    sheet.getRange(rowIndex, 13).setValue(formData.room || '');         // M
    sheet.getRange(rowIndex, 16).setValue(formData.departemen || '');   // P
    sheet.getRange(rowIndex, 17).setValue(formData.unitKerja || '');    // Q
    sheet.getRange(rowIndex, 18).setValue(formData.jumlahHadir || '');  // R
    sheet.getRange(rowIndex, 20).setValue(formData.durasi || '');       // T
    sheet.getRange(rowIndex, 21).setValue(formData.kehadiran || '');    // U
    
    // Return updated data
    var updatedData = getRealizationData();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateRealizationData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Delete realization data by ID
 */
function deleteRealizationData(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pelaksanaan');
    
    if (!sheet) {
      return { success: false, error: 'Sheet Pelaksanaan not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // Find row by ID (column A)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1; // Row number (1-indexed)
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found with ID: ' + id };
    }
    
    sheet.deleteRow(rowIndex);
    
    // Return updated data
    var updatedData = getRealizationData();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteRealizationData: ' + e.message);
    return { success: false, error: e.message };
  }
}
