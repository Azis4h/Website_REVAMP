/**
 * CODE.GS - RESTORED & MERGED VERSION
 */

function doGet() {
  // Always serve Index since authentication checks are handled client-side
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistem Informasi Sertifikasi')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns the email of the active user (Legacy / local dev fallback)
 */
function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  } catch (e) {
    return "Admin Dashboard";
  }
}

/**
 * Legacy access check based on Google Session. Kept for backwards compatibility.
 */
function getUserAccess() {
  var email = getUserEmail();
  var result = {
    authorized: false,
    role: 'Viewer',
    email: email
  };
  
  if (!email || email === "Admin Dashboard") {
    var devEmail = Session.getEffectiveUser().getEmail();
    if (devEmail) {
      email = devEmail;
      result.email = devEmail;
    } else {
      return result;
    }
  }
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Akses_User');
    if (!sheet) {
      initializeAuthSheets();
      sheet = ss.getSheetByName('Akses_User');
    }
    
    var data = sheet.getDataRange().getValues();
    var cleanEmail = email.toLowerCase().trim();
    
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var emailIndex = headers.indexOf('Email');
    var roleIndex = headers.indexOf('Role');
    
    for (var i = 1; i < data.length; i++) {
      var dbEmail = String(data[i][emailIndex]).toLowerCase().trim();
      if (dbEmail === cleanEmail) {
        result.authorized = true;
        result.role = String(data[i][roleIndex]).trim();
        return result;
      }
    }
  } catch (e) {
    Logger.log('Error in getUserAccess: ' + e.message);
  }
  return result;
}

// Hashing Password using SHA-256
function hashPassword(password) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.CharSet.UTF_8);
  var signature = [];
  for (var i = 0; i < digest.length; i++) {
    var byteVal = digest[i];
    if (byteVal < 0) byteVal += 256;
    var byteString = byteVal.toString(16);
    if (byteString.length == 1) byteString = "0" + byteString;
    signature.push(byteString);
  }
  return signature.join("");
}

// Initialize Akses_User and User_Sessions sheet if they don't exist
function initializeAuthSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Check or create Akses_User
  var userSheet = ss.getSheetByName('Akses_User');
  if (!userSheet) {
    userSheet = ss.insertSheet('Akses_User');
    userSheet.appendRow(['Email', 'Password', 'Nama', 'Role']);
    var ownerEmail = Session.getEffectiveUser().getEmail() || "admin@example.com";
    var defaultPasswordHash = hashPassword('admin123');
    userSheet.appendRow([ownerEmail, defaultPasswordHash, 'Owner / Super Admin', 'Super Admin']);
  } else {
    var headers = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0];
    var emailIndex = headers.indexOf('Email');
    var passwordIndex = headers.indexOf('Password');
    var namaIndex = headers.indexOf('Nama');
    var roleIndex = headers.indexOf('Role');
    
    // Auto-migrate if password column does not exist
    if (passwordIndex === -1) {
      userSheet.insertColumnAfter(1); // Column 2
      userSheet.getRange(1, 2).setValue('Password');
      
      var lastRow = userSheet.getLastRow();
      if (lastRow > 1) {
        var defaultPasswordHash = hashPassword('admin123');
        var passwordValues = [];
        for (var i = 2; i <= lastRow; i++) {
          passwordValues.push([defaultPasswordHash]);
        }
        userSheet.getRange(2, 2, lastRow - 1, 1).setValues(passwordValues);
      }
    }
  }
  
  // 2. Check or create User_Sessions
  var sessionSheet = ss.getSheetByName('User_Sessions');
  if (!sessionSheet) {
    sessionSheet = ss.insertSheet('User_Sessions');
    sessionSheet.appendRow(['Token', 'Email', 'ExpiresAt']);
  }
}

// Validate session token
function validateSession(token) {
  if (!token) return { isValid: false };
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sessionSheet = ss.getSheetByName('User_Sessions');
    if (!sessionSheet) {
      initializeAuthSheets();
      sessionSheet = ss.getSheetByName('User_Sessions');
    }
    
    var sessionData = sessionSheet.getDataRange().getValues();
    var now = new Date().getTime();
    var tokenRowIndex = -1;
    var email = "";
    
    for (var i = 1; i < sessionData.length; i++) {
      if (String(sessionData[i][0]) === token) {
        var expiresAt = new Date(sessionData[i][2]).getTime();
        if (expiresAt > now) {
          tokenRowIndex = i + 1;
          email = String(sessionData[i][1]);
          break;
        } else {
          // Clean up expired session
          sessionSheet.deleteRow(i + 1);
          sessionData.splice(i, 1);
          i--;
        }
      }
    }
    
    if (tokenRowIndex === -1) return { isValid: false };
    
    // Fetch user details from Akses_User
    var userSheet = ss.getSheetByName('Akses_User');
    if (!userSheet) return { isValid: false };
    var userData = userSheet.getDataRange().getValues();
    var cleanEmail = email.toLowerCase().trim();
    
    var userHeaders = userData[0].map(function(h) { return String(h).trim(); });
    var emailIdx = userHeaders.indexOf('Email');
    var namaIdx = userHeaders.indexOf('Nama');
    var roleIdx = userHeaders.indexOf('Role');
    
    for (var j = 1; j < userData.length; j++) {
      var dbEmail = String(userData[j][emailIdx]).toLowerCase().trim();
      if (dbEmail === cleanEmail) {
        return {
          isValid: true,
          email: email,
          nama: String(userData[j][namaIdx]),
          role: String(userData[j][roleIdx])
        };
      }
    }
  } catch (e) {
    Logger.log("Error in validateSession: " + e.message);
  }
  return { isValid: false };
}

// User Login Endpoint
function loginUser(email, password) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Initialize auth sheets if needed
    initializeAuthSheets();
    
    var userSheet = ss.getSheetByName('Akses_User');
    var data = userSheet.getDataRange().getValues();
    var cleanEmail = String(email).toLowerCase().trim();
    var hashedPassword = hashPassword(String(password));
    
    var userHeaders = data[0].map(function(h) { return String(h).trim(); });
    var emailIdx = userHeaders.indexOf('Email');
    var passwordIdx = userHeaders.indexOf('Password');
    var namaIdx = userHeaders.indexOf('Nama');
    var roleIdx = userHeaders.indexOf('Role');
    
    var foundUser = null;
    for (var i = 1; i < data.length; i++) {
      var dbEmail = String(data[i][emailIdx]).toLowerCase().trim();
      var dbPassword = String(data[i][passwordIdx]);
      
      if (dbEmail === cleanEmail && dbPassword === hashedPassword) {
        foundUser = {
          email: String(data[i][emailIdx]),
          nama: String(data[i][namaIdx]),
          role: String(data[i][roleIdx])
        };
        break;
      }
    }
    
    if (!foundUser) {
      return { success: false, error: 'Email atau password salah.' };
    }
    
    // Create new session
    var sessionSheet = ss.getSheetByName('User_Sessions');
    var token = Utilities.getUuid();
    var expiresAt = new Date(new Date().getTime() + 24 * 60 * 60 * 1000); // 24 hours
    
    sessionSheet.appendRow([token, foundUser.email, expiresAt.toISOString()]);
    
    return {
      success: true,
      token: token,
      user: foundUser
    };
  } catch (e) {
    Logger.log("Error in loginUser: " + e.message);
    return { success: false, error: 'Terjadi kesalahan sistem: ' + e.message };
  }
}

// User Logout Endpoint
function logoutUser(token) {
  if (!token) return { success: true };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sessionSheet = ss.getSheetByName('User_Sessions');
    if (!sessionSheet) return { success: true };
    
    var data = sessionSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === token) {
        sessionSheet.deleteRow(i + 1);
        break;
      }
    }
  } catch (e) {
    Logger.log("Error in logoutUser: " + e.message);
  }
  return { success: true };
}

// Helper to check write access with token
function hasWriteAccessSession(token) {
  var session = validateSession(token);
  return session.isValid && (session.role === 'Super Admin' || session.role === 'Admin' || session.role === 'Admin LND');
}

// Legacy helper function (will check default developer session if token is omitted)
function hasWriteAccess() {
  var access = getUserAccess();
  return access.authorized && (access.role === 'Super Admin' || access.role === 'Admin' || access.role === 'Admin LND');
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

var SPREADSHEET_ID = '1XE_Utc16DPIdn_yc0Ic2nurtPKkjndJu64sAz7RAz5Q'; // MASTER
var VENDOR_BACKUP_ID = '1sASC3Me80zWBGPfqIUhaPRKZ4jpK359m-V9a0kuwOyo'; // BACKUP
var PESERTA_SPREADSHEET_ID = '1XE_Utc16DPIdn_yc0Ic2nurtPKkjndJu64sAz7RAz5Q';
var MAIN_SHEET_NAME_CAP = 'Perencanaan';
var MAIN_SHEET_NAME_LOWER = 'perencanaan';

function connect() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME_CAP) || ss.getSheetByName(MAIN_SHEET_NAME_LOWER);
  return sheet;
}

/* --- 1. DATA SERTIFIKASI (KIRI) --- */
function getData(token) {
  if (!validateSession(token).isValid) return [];
  try {
    var sheet = connect();
    if (!sheet) return []; // Safety check
    
    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    var lastSAP = "";
    var lastNama = "";

    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (String(r[0]).toUpperCase() === "NO") continue;
        
        // Update last SAP/Nama if present
        if (r[1]) lastSAP = cleanString(r[1]);
        if (r[2]) lastNama = String(r[2]);

        // SKIP: Jika baris dianggap kosong total (tidak ada SAP/Nama dan tidak ada Item ID)
        if (!r[1] && !r[2] && !r[3]) continue;

        try {
            var certItemId = r[3]; // Kolom D
            
            if ((certItemId && String(certItemId).trim() !== "") || (r[1] && r[2])) {
                var id = r[0] ? String(r[0]) : "ROW_" + i;

                data.push({
                    id: id,          
                    sap: lastSAP, 
                    nama: lastNama,        
                    itemId: String(r[3]),      
                    judul: String(r[4]),       
                    periode: safeParseDate(r[5]), 
                    jumlah: String(r[6]),           
                    statusAnggaran: String(r[7]),   
                    mandatory: String(r[8]),        
                    resiko: String(r[9]),           
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
function getLATData(token) {
  if (!validateSession(token).isValid) return [];
  try {
    var sheet = connect();
    if (!sheet) return [];

    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    var lastSAP = "";
    var lastNama = "";

    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (String(r[0]).toUpperCase() === "NO") continue;

        if (r[1]) lastSAP = cleanString(r[1]);
        if (r[2]) lastNama = String(r[2]);

        if (!r[1] && !r[2] && !r[11]) continue; 

        try {
            var latItemId = r[11];
            
            if ((latItemId && String(latItemId).trim() !== "") || (r[1] && r[2] && r[11])) {
                 var id = r[0] ? String(r[0]) : "ROW_" + i;
                 
                 data.push({
                    id: id + "_LAT", 
                    originalId: id,
                    sap: lastSAP,
                    nama: lastNama,
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

function addData(token, formObject) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menambah data.' };
  var sheet = connect();
  
  var id = "=ROW()-1"; 
  
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
  copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
  return { success: true };
}

function addLATData(token, formObject) {
    if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menambah data.' };
    var sheet = connect();
    
    var id = "=ROW()-1";

    var newRow = [
        id, formObject.sap, formObject.nama,
        "", "", "", "", "", "", "", 
        "", 
        formObject.itemId, formObject.judul, formObject.instruktur, 
        formObject.periode, formObject.resiko
    ];
    sheet.appendRow(newRow);
    copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
    return { success: true };
}

/* --- UPDATE DATA SERTIFIKASI --- */
function updateData(token, formData) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk mengubah data.' };
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(formData.id)) {
        rowIndex = i + 1; // 1-indexed untuk getRange
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data tidak ditemukan dengan ID: ' + formData.id };

    // Update kolom Cert: B(SAP), C(Nama), D(ItemId), E(Judul), F(Periode), G(Jumlah), H(StatusAnggaran), I(Mandatory), J(Resiko)
    sheet.getRange(rowIndex, 2).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 3).setValue(formData.nama || '');
    sheet.getRange(rowIndex, 4).setValue(formData.itemId || '');
    sheet.getRange(rowIndex, 5).setValue(formData.judul || '');
    sheet.getRange(rowIndex, 6).setValue(formData.periode || '');
    sheet.getRange(rowIndex, 7).setValue(formData.jumlah || '');
    sheet.getRange(rowIndex, 8).setValue(formData.statusAnggaran || '');
    sheet.getRange(rowIndex, 9).setValue(formData.mandatory || '');
    sheet.getRange(rowIndex, 10).setValue(formData.resiko || '');

    return { success: true };
  } catch (e) {
    Logger.log('Error in updateData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- UPDATE DATA LAT --- */
function updateLATData(token, formData) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk mengubah data.' };
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    // ID LAT format: "X_LAT" — ambil original ID dengan hapus "_LAT"
    var originalId = String(formData.id).replace('_LAT', '');

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === originalId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data LAT tidak ditemukan dengan ID: ' + originalId };

    // Update kolom LAT: B(SAP), C(Nama), L(ItemId), M(Judul), N(Instruktur), O(Periode), P(Resiko)
    sheet.getRange(rowIndex, 2).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 3).setValue(formData.nama || '');
    sheet.getRange(rowIndex, 12).setValue(formData.itemId || '');
    sheet.getRange(rowIndex, 13).setValue(formData.judul || '');
    sheet.getRange(rowIndex, 14).setValue(formData.instruktur || '');
    sheet.getRange(rowIndex, 15).setValue(formData.periode || '');
    sheet.getRange(rowIndex, 16).setValue(formData.resiko || '');

    return { success: true };
  } catch (e) {
    Logger.log('Error in updateLATData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- DELETE DATA SERTIFIKASI --- */
function deleteData(token, id) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menghapus data.' };
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data tidak ditemukan dengan ID: ' + id };

    // Cek apakah baris ini juga punya data LAT (kolom L / index 11)
    var hasLAT = rows[rowIndex - 1][11] && String(rows[rowIndex - 1][11]).trim() !== '';

    if (hasLAT) {
      // Baris punya LAT juga — hanya kosongkan kolom Cert (D-J) agar data LAT aman
      sheet.getRange(rowIndex, 4, 1, 7).clearContent(); // D=4 sampai J=10 (7 kolom)
    } else {
      // Baris hanya Cert — hapus seluruh baris
      sheet.deleteRow(rowIndex);
    }

    return { success: true };
  } catch (e) {
    Logger.log('Error in deleteData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- DELETE DATA LAT --- */
function deleteLATData(token, id) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menghapus data.' };
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    // ID LAT format: "X_LAT"
    var originalId = String(id).replace('_LAT', '');

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === originalId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data LAT tidak ditemukan dengan ID: ' + originalId };

    // Cek apakah baris ini juga punya data Cert (kolom D / index 3)
    var hasCert = rows[rowIndex - 1][3] && String(rows[rowIndex - 1][3]).trim() !== '';

    if (hasCert) {
      // Baris punya Cert juga — hanya kosongkan kolom LAT (L-P) agar data Cert aman
      sheet.getRange(rowIndex, 12, 1, 5).clearContent(); // L=12 sampai P=16 (5 kolom)
    } else {
      // Baris hanya LAT — hapus seluruh baris
      sheet.deleteRow(rowIndex);
    }

    return { success: true };
  } catch (e) {
    Logger.log('Error in deleteLATData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- DATA PELAKSANAAN (UPDATED: SESUAI USER HEADERS) --- */
function getRealizationData(token) {
  if (!validateSession(token).isValid) return [];
  try {
    var ss = SpreadsheetApp.openById(PESERTA_SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Peserta'); 
    
    if (!sheet) {
      Logger.log('Sheet peserta tidak ditemukan di Spreadsheet baru');
      return [];
    }

    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        
        try {
            // Skip jika baris kosong (cek SAP atau Course Title)
            if (!r[9] && !r[8]) continue;

            var dateStart = safeParseDate(r[1]);
            var dateEnd = safeParseDate(r[2]);
            
            var tahun = r[4] ? String(r[4]).trim() : "";
            
            // Fallback tahun jika kosong
            if (!tahun && dateStart) {
                var d = new Date(dateStart);
                if (!isNaN(d.getTime())) tahun = String(d.getFullYear());
            }

            data.push({
                id: r[0] ? String(r[0]) : "P_" + i,
                rowIndex: i + 1,
                sapStart: dateStart,         
                end: dateEnd,              
                bulan: r[3] ? String(r[3]) : "",               
                tahun: tahun,                                   
                itemId: r[5] ? String(r[5]) : "",              
                sapInstruktur: r[6] ? String(r[6]) : "",       
                namaInstruktur: r[7] ? String(r[7]) : "",      
                courseTitle: r[8] ? String(r[8]) : "",
                judulPelatihan: r[8] ? String(r[8]) : "", 
                sapPeserta: r[9] ? String(r[9]) : "",        
                namaPeserta: r[10] ? String(r[10]) : "",       
                room: r[11] ? String(r[11]) : "",              
                presensi: r[12] ? String(r[12]) : "",
                ket: r[13] ? String(r[13]) : "",
                departemen: r[14] ? String(r[14]) : "",
                
                // NEW COLUMNS FROM SCREENSHOT
                unitKerja: r[15] ? String(r[15]) : "",
                jumlahHadir: r[16] ? String(r[16]) : "",
                countPelatihan: r[17] ? String(r[17]) : "",
                durasi: r[18] ? String(r[18]) : "",
                kehadiran: r[19] ? String(r[19]) : "",
                durasiPelatihan: r[20] ? String(r[20]) : "",
                durasiIndividu: r[21] ? String(r[21]) : "",
                
                // Compatibility Fields
                sap: r[9] ? String(r[9]) : "NO_SAP",
                nama: r[10] ? String(r[10]) : "No Name"
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
// Removed realization CRUD functions (addRealizationData, updateRealizationData, deleteRealizationData)

/* ─────────────────────────────────────────────────────────────────────────────
 * EVALUASI L1 DATA OPERATIONS
 * ───────────────────────────────────────────────────────────────────────────── */

/**
 * Get all L1 evaluation data from sheet "L1"
 */
function getEvaluasiL1Data(token) {
  if (!validateSession(token).isValid) return [];
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      Logger.log('Sheet L1 not found');
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) {
      Logger.log('No data in L1 sheet');
      return [];
    }

    var data = [];
    
    // Start from row 2 (skip header)
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      
      // Skip empty rows
      if (!r[0] && !r[1]) continue;
      
        data.push({
        id: r[0] ? String(r[0]) : '',                           
        judulPelatihan: r[1] ? String(r[1]) : '',               
        pelaksanaanId: r[2] ? safeParseDate(r[2]) : '',               
        sap: r[3] ? String(r[3]) : '',                          
        namaPeserta: r[4] ? String(r[4]) : '',                  
        tempatPembelajaran: r[5] ? String(r[5]) : '',           
        fasilitasMedia: r[6] ? String(r[6]) : '',               
        pelayananUmum: r[7] ? String(r[7]) : '',                
        ratapenyelenggaraan: r[8] ? String(r[8]) : '',         
        materi: r[9] ? String(r[9]) : '',                       
        tujuanTercapai: r[10] ? String(r[10]) : '',              
        penyajian: r[11] ? String(r[11]) : '',                   
        disiplin: r[12] ? String(r[12]) : '',                    
        rataPembelajaran: r[13] ? String(r[13]) : '',            
        pengetahuan: r[14] ? String(r[14]) : '',                 
        presentasi: r[15] ? String(r[15]) : '',                  
        perilaku: r[16] ? String(r[16]) : '',                    
        waktu: r[17] ? String(r[17]) : '',                       
        rataInstruktur: r[18] ? String(r[18]) : '',              
        rataKeseluruhan: r[19] ? String(r[19]) : '',             
        komentarPeserta: r[20] ? String(r[20]) : ''              
      });
    }
    
    Logger.log('L1 data loaded: ' + data.length + ' records');
    return data;
    
  } catch (e) {
    Logger.log('ERROR getEvaluasiL1Data: ' + e.message);
    return [];
  }
}

/**
 * Add new L1 evaluation
 */
function addEvaluasiL1(token, formData) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menambah data evaluasi L1.' };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      return { success: false, error: 'Sheet L1 not found' };
    }

    var nextId = "=ROW()-1";
    
    var newRow = [
      nextId,                                     // A
      formData.judulPelatihan || '',              // B
      formData.pelaksanaanId || '',               // C
      formData.sap || '',                         // D
      formData.namaPeserta || '',                 // E
      formData.tempatPembelajaran || '',          // F
      formData.fasilitasMedia || '',              // G
      formData.pelayananUmum || '',               // H
      formData.ratapenyelenggaraan || '',         // I (Manual Input)
      formData.materi || '',                      // J
      formData.tujuanTercapai || '',              // K
      formData.penyajian || '',                   // L
      formData.disiplin || '',                    // M
      formData.rataPembelajaran || '',            // N (Manual Input)
      formData.pengetahuan || '',                 // O
      formData.presentasi || '',                  // P
      formData.perilaku || '',                    // Q
      formData.waktu || '',                       // R
      formData.rataInstruktur || '',              // S (Manual Input)
      formData.rataKeseluruhan || '',             // T (Manual Input)
      formData.komentarPeserta || ''              // U
    ];
    
    sheet.appendRow(newRow);
    copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
    
    var updatedData = getEvaluasiL1Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addEvaluasiL1: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Update L1 evaluation
 */
function updateEvaluasiL1(token, formData) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk mengubah data evaluasi L1.' };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      return { success: false, error: 'Sheet L1 not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found' };
    }
    
    // Update all fields
    sheet.getRange(rowIndex, 2).setValue(formData.judulPelatihan || '');
    sheet.getRange(rowIndex, 3).setValue(formData.pelaksanaanId || '');
    sheet.getRange(rowIndex, 4).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 5).setValue(formData.namaPeserta || '');
    sheet.getRange(rowIndex, 6).setValue(formData.tempatPembelajaran || '');
    sheet.getRange(rowIndex, 7).setValue(formData.fasilitasMedia || '');
    sheet.getRange(rowIndex, 8).setValue(formData.pelayananUmum || '');
    sheet.getRange(rowIndex, 9).setValue(formData.ratapenyelenggaraan || ''); // Manual Input
    sheet.getRange(rowIndex, 10).setValue(formData.materi || '');
    sheet.getRange(rowIndex, 11).setValue(formData.tujuanTercapai || '');
    sheet.getRange(rowIndex, 12).setValue(formData.penyajian || '');
    sheet.getRange(rowIndex, 13).setValue(formData.disiplin || '');
    sheet.getRange(rowIndex, 14).setValue(formData.rataPembelajaran || ''); // Manual Input
    sheet.getRange(rowIndex, 15).setValue(formData.pengetahuan || '');
    sheet.getRange(rowIndex, 16).setValue(formData.presentasi || '');
    sheet.getRange(rowIndex, 17).setValue(formData.perilaku || '');
    sheet.getRange(rowIndex, 18).setValue(formData.waktu || '');
    sheet.getRange(rowIndex, 19).setValue(formData.rataInstruktur || ''); // Manual Input
    sheet.getRange(rowIndex, 20).setValue(formData.rataKeseluruhan || ''); // Manual Input
    sheet.getRange(rowIndex, 21).setValue(formData.komentarPeserta || '');
    
    var updatedData = getEvaluasiL1Data(token);
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateEvaluasiL1: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Delete L1 evaluation
 */
function deleteEvaluasiL1(token, id) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menghapus data evaluasi L1.' };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      return { success: false, error: 'Sheet L1 not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found' };
    }
    
    sheet.deleteRow(rowIndex);
    
    var updatedData = getEvaluasiL1Data(token);
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteEvaluasiL1: ' + e.message);
    return { success: false, error: e.message };
  }
}

/** 
 * =================================================================================
 * EVALUASI L2 (LEARNING) - CRUD
 * =================================================================================
 */

function getEvaluasiL2Data(token) {
  if (!validateSession(token).isValid) return [];
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    
    if (!sheet) {
      // Auto-create if not exists
      sheet = ss.insertSheet('L2');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan', 'SAP', 'Nama Peserta', 
        'Pre Test', 'Post Test', 'Increase', 'Ket.'
      ]);
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) return [];

    var data = [];
    
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r[0] && !r[1] && !r[4]) continue; // Check ID, Judul, or Nama
      
      data.push({
        id: r[0] ? String(r[0]) : '',
        judulPelatihan: r[1] ? String(r[1]) : '',
        pelaksanaanId: r[2] ? safeParseDate(r[2]) : '',
        sap: r[3] ? String(r[3]) : '',
        namaPeserta: r[4] ? String(r[4]) : '',
        preTest: r[5] ? String(r[5]) : '0',
        postTest: r[6] ? String(r[6]) : '0',
        increase: r[7] ? String(r[7]) : '0',       
        ket: r[8] ? String(r[8]) : ''            
      });
    }
    
    return data;
    
  } catch (e) {
    Logger.log('ERROR getEvaluasiL2Data: ' + e.message);
    return [];
  }
}

function addEvaluasiL2(token, formData) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menambah data evaluasi L2.' };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    
    if (!sheet) {
      sheet = ss.insertSheet('L2');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan', 'SAP', 'Nama Peserta', 
        'Pre Test', 'Post Test', 'Increase', 'Ket.'
      ]);
    }

    var nextId = "=ROW()-1";
    var increase = (parseFloat(formData.postTest) || 0) - (parseFloat(formData.preTest) || 0);
    
    var newRow = [
      nextId,
      formData.judulPelatihan || '',
      formData.pelaksanaanId || '',
      formData.sap || '',
      formData.namaPeserta || '',
      formData.preTest || 0,
      formData.postTest || 0,
      increase.toFixed(2),
      formData.ket || ''
    ];
    
    sheet.appendRow(newRow);
    copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
    
    var updatedData = getEvaluasiL2Data(token);
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addEvaluasiL2: ' + e.message);
    return { success: false, error: e.message };
  }
}

function updateEvaluasiL2(token, formData) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk mengubah data evaluasi L2.' };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    if (!sheet) return { success: false, error: 'Sheet L2 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    var increase = (parseFloat(formData.postTest) || 0) - (parseFloat(formData.preTest) || 0);

    // Update fields (Columns 2-9)
    sheet.getRange(rowIndex, 2).setValue(formData.judulPelatihan || '');
    sheet.getRange(rowIndex, 3).setValue(formData.pelaksanaanId || '');
    sheet.getRange(rowIndex, 4).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 5).setValue(formData.namaPeserta || '');
    sheet.getRange(rowIndex, 6).setValue(formData.preTest || 0);
    sheet.getRange(rowIndex, 7).setValue(formData.postTest || 0);
    sheet.getRange(rowIndex, 8).setValue(increase.toFixed(2));
    sheet.getRange(rowIndex, 9).setValue(formData.ket || '');
    
    var updatedData = getEvaluasiL2Data(token);
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateEvaluasiL2: ' + e.message);
    return { success: false, error: e.message };
  }
}

function deleteEvaluasiL2(token, id) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menghapus data evaluasi L2.' };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    if (!sheet) return { success: false, error: 'Sheet L2 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    sheet.deleteRow(rowIndex);
    
    var updatedData = getEvaluasiL2Data(token);
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteEvaluasiL2: ' + e.message);
    return { success: false, error: e.message };
  }
}

/** 
 * =================================================================================
 * EVALUASI L3 (BEHAVIOR) - CRUD
 *Headers: No, Judul Pelatihan, Pelaksanaan Learning, SAP, Nama Peserta, Nilai Evaluasi, Ket., Key Behaviour, Tanggal Eval
 * =================================================================================
 */

function getEvaluasiL3Data(token) {
  if (!validateSession(token).isValid) return [];
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    
    if (!sheet) {
      sheet = ss.insertSheet('L3');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan Learning', 'SAP', 'Nama Peserta', 
        'Nilai Evaluasi', 'Ket.', 'Key Behaviour', 'Tanggal Eval'
      ]);
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) return [];

    var data = [];
    
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r[0] && !r[1] && !r[4]) continue;
      
      data.push({
        id: r[0] ? String(r[0]) : '',
        judulPelatihan: r[1] ? String(r[1]) : '',
        pelaksanaanId: r[2] ? String(r[2]) : '',
        sap: r[3] ? String(r[3]) : '',
        namaPeserta: r[4] ? String(r[4]) : '',
        nilaiEvaluasi: r[5] ? String(r[5]) : '',
        ket: r[6] ? String(r[6]) : '',
        keyBehaviour: r[7] ? String(r[7]) : '',
        tanggalEval: r[8] ? Utilities.formatDate(new Date(r[8]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : ''
      });
    }
    
    return data;
    
  } catch (e) {
    Logger.log('ERROR getEvaluasiL3Data: ' + e.message);
    return [];
  }
}

function addEvaluasiL3(token, formData) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menambah data evaluasi L3.' };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    
    if (!sheet) {
      sheet = ss.insertSheet('L3');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan Learning', 'SAP', 'Nama Peserta', 
        'Nilai Evaluasi', 'Ket.', 'Key Behaviour', 'Tanggal Eval'
      ]);
    }

    var nextId = "=ROW()-1";
    
    var newRow = [
      nextId,
      formData.judulPelatihan || '',
      formData.pelaksanaanId || '',
      formData.sap || '',
      formData.namaPeserta || '',
      formData.nilaiEvaluasi || '',
      formData.ket || '',
      formData.keyBehaviour || '',
      formData.tanggalEval || ''
    ];
    
    sheet.appendRow(newRow);
    copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
    
    var updatedData = getEvaluasiL3Data(token);
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addEvaluasiL3: ' + e.message);
    return { success: false, error: e.message };
  }
}

function updateEvaluasiL3(token, formData) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk mengubah data evaluasi L3.' };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    if (!sheet) return { success: false, error: 'Sheet L3 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    // Update fields (Columns 2-9)
    sheet.getRange(rowIndex, 2).setValue(formData.judulPelatihan || '');
    sheet.getRange(rowIndex, 3).setValue(formData.pelaksanaanId || '');
    sheet.getRange(rowIndex, 4).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 5).setValue(formData.namaPeserta || '');
    sheet.getRange(rowIndex, 6).setValue(formData.nilaiEvaluasi || '');
    sheet.getRange(rowIndex, 7).setValue(formData.ket || '');
    sheet.getRange(rowIndex, 8).setValue(formData.keyBehaviour || '');
    sheet.getRange(rowIndex, 9).setValue(formData.tanggalEval || '');
    
    var updatedData = getEvaluasiL3Data(token);
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateEvaluasiL3: ' + e.message);
    return { success: false, error: e.message };
  }
}

function deleteEvaluasiL3(token, id) {
  if (!hasWriteAccessSession(token)) return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menghapus data evaluasi L3.' };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    if (!sheet) return { success: false, error: 'Sheet L3 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    sheet.deleteRow(rowIndex);
    
    var updatedData = getEvaluasiL3Data(token);
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteEvaluasiL3: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* ─────────────────────────────────────────────────────────────────────────────
 * VENDOR DATA OPERATIONS (NEW)
 * ───────────────────────────────────────────────────────────────────────────── */

function getVendorData(token) {
  if (!validateSession(token).isValid) return [];
  try {
    // Read from MASTER
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Ajuan Vendor");
    if (!sheet) {
      Logger.log("Sheet 'Ajuan Vendor' not found");
      return [];
    }
    
    var rows = sheet.getDataRange().getValues();
    if (rows.length < 1) return [];

    // Map headers to indices
    var headers = rows[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var col = {
      timestamp: headers.indexOf("timestamp"),
      vendor: headers.indexOf("nama vendor"),
      lsp: headers.indexOf("nama lsp"),
      pic: headers.indexOf("pic vendor"),
      sertifikasi: headers.indexOf("nama sertifikasi"),
      jenis: headers.indexOf("jenis sertifikasi"),
      silabus: headers.indexOf("silabus"),
      biaya: headers.indexOf("biaya"),
      metode: headers.indexOf("metode pelaksanaan"),
      tempat: headers.indexOf("tempat pelaksanaan"),
      tanggal: headers.indexOf("tanggal"),
      file: headers.indexOf("file brosur/proposal/screenshot lsp dan pjk3"),
      status: headers.indexOf("status"),
      approver: headers.indexOf("yang menyetujui"),
      approvalTime: headers.indexOf("waktu penyetujuan"),
      emailPic: headers.indexOf("email pic")
    };

    // Fallbacks based on your specific spreadsheet structure (I=8, J=9)
    if (col.metode === -1) col.metode = 8;
    if (col.tempat === -1) col.tempat = 9;

    // Fallbacks if header names don't match exactly
    if (col.vendor === -1) col.vendor = 1;
    if (col.status === -1) col.status = 12;
    if (col.approver === -1) col.approver = 13;
    if (col.approvalTime === -1) col.approvalTime = 14;
    if (col.emailPic === -1) col.emailPic = 15;

    Logger.log("Detected Columns: " + JSON.stringify(col));

    var richTextValues = sheet.getDataRange().getRichTextValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r || r.length < 2) continue;
      
      var vendorName = String(r[col.vendor] || "").trim();
      if (!vendorName && !r[col.sertifikasi]) continue; 
      
      var fileUrl = "-";
      if (col.file !== -1 && richTextValues[i][col.file]) {
        fileUrl = richTextValues[i][col.file].getLinkUrl() || r[col.file] || "-";
      }

      var biayaRaw = col.biaya !== -1 ? r[col.biaya] : "";
      var biayaFormatted = "-";
      if (biayaRaw) {
        if (typeof biayaRaw === "number") {
          biayaFormatted = "Rp " + biayaRaw.toLocaleString('id-ID');
        } else {
          var cleanNum = Number(String(biayaRaw).replace(/[^\d]/g, ''));
          if (!isNaN(cleanNum) && cleanNum > 0) {
            biayaFormatted = "Rp " + cleanNum.toLocaleString('id-ID');
          } else {
            biayaFormatted = String(biayaRaw);
          }
        }
      }
      
      data.push({
        rowIndex: i + 1,
        timestamp: (col.timestamp !== -1 && r[col.timestamp] instanceof Date) ? r[col.timestamp].getTime() : 0,
        namaVendor: vendorName || "-",
        namaLsp: col.lsp !== -1 ? String(r[col.lsp] || "-") : "-",
        picVendor: col.pic !== -1 ? String(r[col.pic] || "-") : "-",
        namaSertifikasi: col.sertifikasi !== -1 ? String(r[col.sertifikasi] || "-") : "-",
        silabus: col.silabus !== -1 ? String(r[col.silabus] || "-") : "-",
        biaya: biayaFormatted,
        metode: col.metode !== -1 ? String(r[col.metode] || "-") : "-",
        tempat: col.tempat !== -1 ? String(r[col.tempat] || "-") : "-",
        tanggal: (col.tanggal !== -1 && r[col.tanggal]) ? safeParseDate(r[col.tanggal]) : "-",
        file: fileUrl,
        status: col.status !== -1 ? (String(r[col.status] || "Pending").trim() || "Pending") : "Pending",
        approver: col.approver !== -1 ? String(r[col.approver] || "-") : "-",
        approvalTime: (col.approvalTime !== -1 && r[col.approvalTime]) ? safeParseDate(r[col.approvalTime]) : "-",
        emailPic: col.emailPic !== -1 ? String(r[col.emailPic] || "-") : "-"
      });
    }
    Logger.log("getVendorData found " + data.length + " records");
    return data;
  } catch (e) {
    Logger.log('ERROR getVendorData: ' + e.message);
    return [];
  }
}

function updateVendorStatus(token, rowIndex, status) {
  var session = validateSession(token);
  if (!session.isValid || !(session.role === 'Super Admin' || session.role === 'Admin' || session.role === 'Admin LND')) {
    return { success: false, error: 'Unauthorized: Anda tidak memiliki akses untuk menyetujui/menolak pengajuan vendor.' };
  }
  try {
    var adminEmail = session.email;
    var timestamp = new Date();

    // --- 1. UPDATE MASTER ---
    var ssMaster = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetMaster = ssMaster.getSheetByName("Ajuan Vendor");
    if (sheetMaster) {
      sheetMaster.getRange(rowIndex, 13).setValue(status); // M
      sheetMaster.getRange(rowIndex, 14).setValue(adminEmail); // N
      sheetMaster.getRange(rowIndex, 15).setValue(timestamp); // O
    }

    // --- 2. UPDATE BACKUP ---
    try {
      var ssBackup = SpreadsheetApp.openById(VENDOR_BACKUP_ID);
      var sheetBackup = ssBackup.getSheetByName("Ajuan Vendor");
      if (sheetBackup) {
        sheetBackup.getRange(rowIndex, 13).setValue(status); // M
        sheetBackup.getRange(rowIndex, 14).setValue(adminEmail); // N
        sheetBackup.getRange(rowIndex, 15).setValue(timestamp); // O
      }
    } catch (e) {
      Logger.log("Gagal update backup: " + e.message);
    }
    
    var sheet = sheetMaster; // Use master for data retrieval below
    SpreadsheetApp.flush(); // Ensure data is written before reading back
    
    // Ambil data untuk keperluan email (ambil 16 kolom untuk mencapai P)
    var rowData = sheet.getRange(rowIndex, 1, 1, 16).getValues()[0];
    var vendorName = rowData[1] || "-";
    var certName = rowData[4] || "-";
    var picName = rowData[3] || "PIC Vendor";
    var picEmail = String(rowData[15] || "").trim();
    
    // 2. Kirim Notifikasi Email jika DISETUJUI
    if (status.toLowerCase() === "disetujui" && picEmail && picEmail.indexOf("@") !== -1) {
      try {
        var subject = "KONFIRMASI PENYETUJUAN AJUAN VENDOR - " + certName.toUpperCase();
        var formLink = "https://script.google.com/macros/s/AKfycbw_3LGkq38oE-4h1l9CZVF5dKqyejG2z0bBPk1HXVU/dev";
        var logoUrl = "https://upload.wikimedia.org/wikipedia/id/d/de/Semen_Tonasa_logo.png";
        
        var htmlBody = `
          <div style="font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 650px; margin: auto; background-color: #f8fafc; padding: 20px;">
            <div style="background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); border: 1px solid #e2e8f0;">
              
              <!-- HEADER -->
              <div style="background-color: #b91c1c; padding: 15px 30px 25px; color: white;">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="75" style="vertical-align: middle;">
                      <div style="background-color: #ffffff; width: 65px; height: 65px; border-radius: 12px; text-align: center; overflow: hidden; display: block; padding: 5px;">
                        <img src="${logoUrl}" alt="Logo" style="max-height: 55px; max-width: 55px; width: auto; display: inline-block; vertical-align: middle;">
                      </div>
                    </td>
                    <td style="padding-left: 20px; vertical-align: middle;">
                      <h1 style="margin: 0; font-size: 20px; font-weight: 700; letter-spacing: 0.5px;">PT SEMEN TONASA</h1>
                      <p style="margin: 2px 0 0; font-size: 12px; opacity: 0.9; font-weight: 400; text-transform: uppercase;">CORPORATE LEARNING & DEVELOPMENT (CLD)</p>
                    </td>
                  </tr>
                </table>
              </div>

              <!-- CONTENT -->
              <div style="padding: 40px; color: #334155; line-height: 1.6;">
                <p style="font-size: 16px; margin-bottom: 25px;">Halo <b>${picName.toUpperCase()}</b>,</p>
                <p style="margin-bottom: 30px; font-size: 14px;">Kami menginformasikan bahwa pengajuan vendor untuk program sertifikasi telah <b style="color: #16a34a;">DISETUJUI</b> oleh departemen terkait dengan rincian sebagai berikut:</p>

                <!-- DETAIL CARD -->
                <div style="background-color: #ffffff; border: 1px solid #f1f5f9; border-radius: 16px; padding: 30px; margin-bottom: 35px; box-shadow: inset 0 2px 4px 0 rgba(0, 0, 0, 0.05);">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <!-- Vendor -->
                    <tr>
                      <td style="padding-bottom: 15px; font-size: 13px; color: #64748b; font-weight: 600; text-transform: uppercase; width: 140px;">Nama Vendor</td>
                      <td style="padding-bottom: 15px; font-size: 14px; font-weight: 700; color: #1e293b;">: ${vendorName}</td>
                    </tr>
                    <!-- Sertifikasi -->
                    <tr>
                      <td style="padding-bottom: 15px; font-size: 13px; color: #64748b; font-weight: 600; text-transform: uppercase;">Sertifikasi</td>
                      <td style="padding-bottom: 15px; font-size: 14px; font-weight: 700; color: #1e293b;">: ${certName}</td>
                    </tr>
                    <!-- Penyetuju -->
                    <tr>
                      <td style="padding-bottom: 15px; font-size: 13px; color: #64748b; font-weight: 600; text-transform: uppercase;">Penyetuju</td>
                      <td style="padding-bottom: 15px; font-size: 14px; color: #0284c7;">: ${adminEmail}</td>
                    </tr>
                    <!-- Waktu -->
                    <tr>
                      <td style="font-size: 13px; color: #64748b; font-weight: 600; text-transform: uppercase;">Waktu</td>
                      <td style="font-size: 14px; color: #1e293b;">: ${Utilities.formatDate(timestamp, "GMT+7", "dd/MM/yyyy HH:mm")} WIB</td>
                    </tr>
                  </table>
                </div>

                <!-- INSTRUCTION BOX -->
                <div style="background-color: #fef2f2; border-left: 5px solid #b91c1c; border-radius: 4px 12px 12px 4px; padding: 25px; margin-bottom: 25px;">
                  <p style="margin: 0; font-weight: 700; color: #b91c1c; font-size: 15px; text-transform: uppercase;">Instruksi Lanjutan</p>
                  <p style="margin: 5px 0 15px; font-size: 13px; color: #450a0a;">Mohon agar <b>PIC Vendor</b> segera mengisi formulir <b>Pretest & Posttest</b> melalui tautan resmi di bawah ini sebagai bagian dari administrasi pelaksanaan program:</p>
                  <a href="${formLink}" style="display: inline-block; padding: 14px 32px; background-color: #b91c1c; color: #ffffff; text-decoration: none; border-radius: 10px; font-weight: 700; font-size: 14px; text-transform: uppercase;">ISI FORM PRETEST & POSTTEST &nbsp; &rarr;</a>
                </div>

                <!-- INFO BOX -->
                <div style="background-color: #eff6ff; border-radius: 12px; padding: 20px; margin-bottom: 35px; border: 1px solid #dbeafe;">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="45" style="vertical-align: middle;">
                        <div style="width: 32px; height: 32px; background-color: #3b82f6; border-radius: 50%; text-align: center; line-height: 32px; color: white; font-weight: bold; font-style: italic; font-size: 18px;">i</div>
                      </td>
                      <td>
                        <p style="margin: 0; font-weight: 700; color: #1d4ed8; font-size: 14px;">INFORMASI</p>
                        <p style="margin: 2px 0 0; font-size: 12px; color: #1e40af;">Email ini dikirim secara otomatis. Mohon tidak membalas email ini. Jika Anda memiliki pertanyaan, silakan <a href="mailto:stdiklat@gmail.com" style="color: #1d4ed8; font-weight: 600; text-decoration: underline;">hubungi tim CLD</a>.</p>
                      </td>
                    </tr>
                  </table>
                </div>

                <p style="font-size: 13px; margin: 0; color: #64748b;">Terima kasih atas kerja samanya dalam mendukung pengembangan kompetensi SDM di PT Semen Tonasa.</p>
                
                <div style="margin-top: 30px; border-top: 1px solid #f1f5f9; padding-top: 25px;">
                  <p style="margin: 0; font-size: 14px; color: #334155;">Salam,</p>
                  <p style="margin: 5px 0 0; font-size: 14px; font-weight: 700; color: #b91c1c;">Corporate Learning & Development (CLD)</p>
                </div>
              </div>

              <!-- FOOTER -->
              <div style="background-color: #f1f5f9; padding: 25px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0;">
                &copy; ${new Date().getFullYear()} <b>Admin CLD PT. Semen Tonasa</b>. Seluruh hak cipta dilindungi undang-undang.<br>
                <span style="margin-top: 5px; display: block; opacity: 0.7;">Pesan ini dihasilkan secara otomatis oleh Sistem Informasi Sertifikasi & LAT.</span>
              </div>
            </div>
          </div>
        `;

        GmailApp.sendEmail(picEmail, subject, "", {
          name: "Admin CLD PT. Semen Tonasa",
          htmlBody: htmlBody
        });
        
        Logger.log("Email HTML berhasil dikirim ke: " + picEmail);
      } catch (mailErr) {
        Logger.log("Gagal kirim email HTML ke " + picEmail + ": " + mailErr.message);
      }
    }

    return { success: true, data: getVendorData(token) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function debugHeaders(token) {
  if (!validateSession(token).isValid) return "Unauthorized: Sesi tidak valid.";
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var p = ss.getSheetByName("Pelaksanaan");
    var l1 = ss.getSheetByName("L1");
    var res = "";
    if (p) res += "PELAKSANAAN HEADERS: " + JSON.stringify(p.getRange(1, 1, 1, 20).getValues()[0]) + "\n";
    if (l1) res += "L1 HEADERS: " + JSON.stringify(l1.getRange(1, 1, 1, 30).getValues()[0]) + "\n";
    return res;
  } catch (e) { return e.message; }
}

/**
 * Helper to copy format from one row to another
 */
function copyRowFormat(sheet, sourceRow, targetRow) {
  try {
    if (sourceRow < 1) return;
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    var sourceRange = sheet.getRange(sourceRow, 1, 1, lastCol);
    var targetRange = sheet.getRange(targetRow, 1, 1, lastCol);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  } catch (e) {
    Logger.log("Error copying format: " + e.message);
  }
}

/* ─────────────────────────────────────────────────────────────────────────────
 * AJUAN KARYAWAN DATA OPERATIONS (NEW)
 * ───────────────────────────────────────────────────────────────────────────── */

function getKaryawanData(token) {
  if (!validateSession(token).isValid) return [];
  try {
    var ss = SpreadsheetApp.openById('1ptbh5lMR9R0Hi5Gc49lF-q_rIeisjFXqSgOv5hFrIr4');
    var sheet = ss.getSheets()[0];
    if (!sheet) {
      Logger.log("Sheet not found in Karyawan Spreadsheet");
      return [];
    }
    var rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    var data = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      // Skip empty row
      if (!r[0] && !r[1] && !r[2]) continue;

      data.push({
        id: "K_" + i,
        rowIndex: i + 1,
        timestamp: r[0] ? safeParseDate(r[0]) : "-",
        namaPemateri: r[1] ? String(r[1]).trim() : "-",
        sap: r[2] ? String(r[2]).trim() : "-",
        judulPemateri: r[3] ? String(r[3]).trim() : "-",
        pretest: r[4] !== undefined && r[4] !== "" ? String(r[4]).trim() : "-",
        postest: r[5] !== undefined && r[5] !== "" ? String(r[5]).trim() : "-"
      });
    }
    return data;
  } catch (e) {
    Logger.log("ERROR getKaryawanData: " + e.message);
    return [];
  }
}

function diagnoseSheets() {
  var report = [];
  report.push("=== DIAGNOSIS SPREADSHEET ===");
  report.push("SPREADSHEET_ID: " + SPREADSHEET_ID);
  report.push("PESERTA_SPREADSHEET_ID: " + PESERTA_SPREADSHEET_ID);
  
  // 1. Cek Spreadsheet Master
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    report.push("✅ Master Spreadsheet Berhasil Dibuka: '" + ss.getName() + "'");
    
    var sheets = ss.getSheets().map(function(s) { return s.getName(); });
    report.push("Tab yang ada: " + JSON.stringify(sheets));
    
    // Cek Perencanaan
    var pSheet = ss.getSheetByName(MAIN_SHEET_NAME_CAP) || ss.getSheetByName(MAIN_SHEET_NAME_LOWER);
    if (pSheet) {
      report.push("✅ Tab '" + pSheet.getName() + "' ditemukan. Baris data: " + pSheet.getLastRow());
    } else {
      report.push("❌ Tab '" + MAIN_SHEET_NAME_CAP + "' TIDAK ditemukan! Periksa apakah nama tab sudah sesuai.");
    }
  } catch (e) {
    report.push("❌ Gagal membuka Master Spreadsheet: " + e.message);
  }
  
  // 2. Cek Spreadsheet Peserta
  try {
    var ssPeserta = SpreadsheetApp.openById(PESERTA_SPREADSHEET_ID);
    report.push("✅ Peserta Spreadsheet Berhasil Dibuka: '" + ssPeserta.getName() + "'");
    var sheetsPeserta = ssPeserta.getSheets().map(function(s) { return s.getName(); });
    report.push("Tab yang ada (Peserta): " + JSON.stringify(sheetsPeserta));
  } catch (e) {
    report.push("❌ Gagal membuka Peserta Spreadsheet: " + e.message);
  }
  
  var result = report.join("\n");
  Logger.log(result);
  return result;
}
