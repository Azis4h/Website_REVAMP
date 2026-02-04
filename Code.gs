/**
 * CODE.GS - ULTRA ROBUST VERSION
 * Memastikan SAP bersih dari spasi dan konsisten.
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

// ID SPREADSHEET (Pastikan akses dibuka untuk Anyone with Link -> Viewer/Editor jika perlu)
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
    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    // Loop mulai baris ke-2 (Index 1)
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (!r[0] || String(r[0]).toUpperCase() === "NO") continue;

        // Cek Kolom D (Index 3) - Item ID Cert
        var certItemId = r[3]; 
        
        // Ambil data jika Item ID ada, ATAU Nama ada (agar tidak skip baris manual)
        if ((certItemId && String(certItemId).trim() !== "") || (r[1] && r[2])) {
            data.push({
                id: String(r[0]),          
                sap: cleanString(r[1]), // BERSIHKAN SAP
                nama: String(r[2]),        
                itemId: String(r[3]),      
                judul: String(r[4]),       
                periode: parseDate(r[5]), 
                jumlah: String(r[6]),      
                statusAnggaran: String(r[7]), 
                mandatory: String(r[8]),   
                resiko: String(r[9]),      
                type: 'cert'
            });
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getData: ' + e.message);
    throw e;
  }
}

/* --- 2. DATA LAT (KANAN) --- */
function getLATData() {
  try {
    var sheet = connect();
    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (!r[0] || String(r[0]).toUpperCase() === "NO") continue;

        // Cek Kolom L (Index 11) - Item ID LAT
        var latItemId = r[11];

        if (latItemId && String(latItemId).trim() !== "") {
             data.push({
                id: String(r[0]) + "_LAT", 
                originalId: String(r[0]),
                sap: cleanString(r[1]), // BERSIHKAN SAP
                nama: String(r[2]),
                itemId: String(r[11]),     
                judul: String(r[12]),      
                instruktur: String(r[13]), 
                periode: parseDate(r[14]),
                resiko: String(r[15]),     
                type: 'lat'
            });
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getLATData: ' + e.message);
     throw e;
  }
}

// HELPER: Bersihkan String (Hapus spasi depan/belakang, uppercase)
function cleanString(val) {
  if (!val) return "";
  return String(val).trim().toUpperCase(); 
}

// HELPER: Format Tanggal
function parseDate(dateVal) {
  if (!dateVal) return "";
  if (Object.prototype.toString.call(dateVal) === '[object Date]') {
    var year = dateVal.getFullYear();
    var months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
    var month = months[dateVal.getMonth()];
    return month + " " + year; 
  }
  return String(dateVal); 
}

/* --- 3. CRUD --- */
function addData(formObject) {
  var sheet = connect();
  var id = new Date().getTime(); 
  var newRow = [
      id, formObject.sap, formObject.nama,
      formObject.itemId, formObject.judul, formObject.periode, 
      formObject.jumlah, formObject.statusAnggaran, formObject.mandatory, formObject.resiko,
      "", "", "", "", "", "" 
  ];
  sheet.appendRow(newRow);
  return { success: true };
}

function addLATData(formObject) {
    var sheet = connect();
    var id = new Date().getTime();
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
