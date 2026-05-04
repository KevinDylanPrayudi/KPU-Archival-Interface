// This function handles the "Independent Page" view
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('KPU\'s Archival')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Management')
    .addItem('Open Dashboard', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Main')
    .setTitle('Company Dashboard')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getArsipPageInitializationData() {
  return {
    arsip: getArsipData(),
    klasifikasi: getKlasifikasiData()
  };
}

function formatRawData(rawData) {
  return rawData.map(row => row.map(cell => {
    if (Object.prototype.toString.call(cell) === '[object Date]') {
      return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    return cell !== null && cell !== undefined ? cell.toString().trim() : "";
  }));
}

// ==========================================
// AUTHENTICATION & LOGIN SYSTEM
// ==========================================
function verifyLogin(email, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pegawai");
  var data = formatRawData(sheet.getDataRange().getValues());
  var cleanEmail = email.toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toLowerCase() === cleanEmail && data[i][2] === password.toString().trim()) {
      return { success: true, name: data[i][0], email: data[i][1] };
    }
  }
  return { success: false, message: "Invalid email or password!" };
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pegawai");
  var rawData = sheet.getDataRange().getValues().slice(1);
  return formatRawData(rawData).map((value, index) => {
    return { rowId : index + 2, name: value[0], email: value[1], password: value[2] || "" }
  });
}

function processForm(name, email, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pegawai");
  var cleanEmail = email.toString().trim().toLowerCase();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim().toLowerCase() === cleanEmail) return { success: false, message: "Error: Email '" + email + "' is already registered!" };
  }
  sheet.appendRow([name, email, password, new Date()]);
  return { success: true, message: "Data Saved!" };
}

function editForm(name, email, password, row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pegawai");
  var cleanEmail = email.toString().trim().toLowerCase();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((i + 1) != row && data[i][1].toString().trim().toLowerCase() === cleanEmail) return { success: false, message: "Error: Email is already in use by another user!" };
  }
  sheet.getRange(row, 1).setValue(name);
  sheet.getRange(row, 2).setValue(email);
  sheet.getRange(row, 3).setValue(password);
  return { success: true, message: "Success! Updated row " + row };
}

function deleteRowFromSheet(row) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pegawai").deleteRow(row);
  return "Success!";
}

// ==========================================
// KLASIFIKASI MODULE FUNCTIONS
// ==========================================
function getKlasifikasiData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Klasifikasi");
  var rawData = sheet.getDataRange().getValues().slice(1);
  return formatRawData(rawData).map((value, index) => {
    return { rowId : index + 2, kode: value[0], nama: value[1], folderLink: value[2] }
  });
}

function processKlasifikasiForm(kode, nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Klasifikasi");
  var cleanKode = kode.toString().trim().toUpperCase();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { 
    if (data[i][0].toString().trim().toUpperCase() === cleanKode) return { success: false, message: "Error: Kode Klasifikasi already exists!" };
  }
  var newFolder = DriveApp.createFolder(cleanKode + " - " + nama.toString().trim());
  sheet.appendRow([cleanKode, nama.toString().trim(), newFolder.getUrl(), new Date()]);
  return { success: true, message: "Success!" };
}

function editKlasifikasiForm(nama, row) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Klasifikasi").getRange(row, 2).setValue(nama);
  return "Success!";
}

function deleteKlasifikasiRow(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Klasifikasi");
  const folderUrl = sheet.getRange(row, 3).getValue();
  if (folderUrl && folderUrl.toString().trim() !== "") {
    try {
      const match = folderUrl.match(/[-\w]{25,}/);
      if (match) DriveApp.getFolderById(match[0]).setTrashed(true); 
    } catch (e) {}
  }
  sheet.deleteRow(row);
  return "Success!";
}

// ==========================================
// ARSIP MODULE FUNCTIONS 
// ==========================================
function getArsipData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arsip");
  if (sheet.getRange(1, 18).getValue() === "") {
    sheet.getRange(1, 18).setValue("Physical Page")
  }

  var rawData = sheet.getDataRange().getValues().slice(1);
  var formattedData = formatRawData(rawData);

  return formattedData.map((value, index) => {
    return {
      rowId: index + 2, recordId: value[0], title: value[1], businessActivity: value[2],
      creator: value[3], unit: value[4], dateCreated: value[5], dateReceived: value[6], 
      recordType: value[7], confLevel: value[8], storageLoc: value[9], format: value[10], 
      status: value[11], pageCount: value[12] || "", box: value[13] || "", rack: value[14] || "",
      modifiedBy: value[15] || "", modifiedAt: value[16] || "", physicalPage: value[20] || ""
    }
  });
}

function processArsipForm(dataObj) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arsip");
  
  // Failsafe: Ensure 18 columns exist
  var maxCols = sheet.getMaxColumns();
  if (maxCols < 18) {
    sheet.insertColumnsAfter(maxCols, 18 - maxCols);
    sheet.getRange(1, 18).setValue("Nomor Halaman Fisik").setFontWeight("bold");
  }

  var cleanRecordId = dataObj.recordId.toString().trim().toUpperCase();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { 
    if (data[i][0].toString().trim().toUpperCase() === cleanRecordId) return { success: false, message: "Error: Record ID exists!" };
  }

  var klasSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Klasifikasi");
  var klasData = klasSheet.getDataRange().getValues();
  var targetFolderId = null;

  for (var k = 1; k < klasData.length; k++) {
    if (klasData[k][0].toString().trim().toUpperCase() === dataObj.businessActivity.toString().trim().toUpperCase()) {
      var folderUrl = klasData[k][2];
      if (folderUrl) {
        var match = folderUrl.match(/[-\w]{25,}/); 
        if (match) targetFolderId = match[0];
      }
      break;
    }
  }

  if (!targetFolderId) return { success: false, message: "Error: No Folder found for Klasifikasi!" };

  var uploadedFileUrl = "";
  if (dataObj.fileData && dataObj.fileData.base64) {
    try {
      var targetFolder = DriveApp.getFolderById(targetFolderId);
      var ext = "";
      if (dataObj.fileData.fileName.lastIndexOf(".") !== -1) ext = dataObj.fileData.fileName.substring(dataObj.fileData.fileName.lastIndexOf("."));
      var blob = Utilities.newBlob(Utilities.base64Decode(dataObj.fileData.base64), dataObj.fileData.mimeType, cleanRecordId + ext);
      var newFile = targetFolder.createFile(blob);
      
      // LOGIKA KEAMANAN: Cek tingkat kerahasiaan
      if (dataObj.confLevel === "Public") {
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } else {
        newFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      }
      
      uploadedFileUrl = newFile.getUrl();
      // BARIS BARU: Otomatis mengubah akses file Google Drive menjadi "Siapa saja yang memiliki link (Viewer)"
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      uploadedFileUrl = newFile.getUrl(); 
    } catch (error) { return { success: false, message: "Drive Error: " + error.message }; }
  } else if (dataObj.manualLink && dataObj.manualLink.trim() !== "") {
    uploadedFileUrl = dataObj.manualLink.trim();
  }

  var currentUser = Session.getActiveUser().getEmail() || "System User";
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");

  sheet.appendRow([
    cleanRecordId, dataObj.title, dataObj.businessActivity, dataObj.creator, dataObj.unit, 
    dataObj.dateCreated, dataObj.dateReceived, dataObj.recordType, dataObj.confLevel, 
    uploadedFileUrl, dataObj.format, dataObj.status, dataObj.pageCount || "", dataObj.box || "", 
    dataObj.rack || "", currentUser, timestamp, dataObj.physicalPage || ""
  ]);
  
  return { success: true, message: "Success! Record saved." };
}

function editArsipForm(dataObj, row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arsip");
  
  var maxCols = sheet.getMaxColumns();
  if (maxCols < 18) {
    sheet.insertColumnsAfter(maxCols, 18 - maxCols);
    sheet.getRange(1, 18).setValue("Physical Page");
  }

  const cleanRecordId = dataObj.recordId.toString().trim().toUpperCase();
  const newBusinessActivity = dataObj.businessActivity.toString().trim().toUpperCase();
  
  const oldRecordId = sheet.getRange(row, 1).getValue().toString().trim().toUpperCase();
  const oldBusinessActivity = sheet.getRange(row, 3).getValue().toString().trim().toUpperCase();
  const oldFileUrl = sheet.getRange(row, 10).getValue(); 
  let finalFileUrl = oldFileUrl; 

  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if ((i + 1) != row && allData[i][0].toString().trim().toUpperCase() === cleanRecordId) return { success: false, message: "Error: Record ID exists!" };
  }

  var klasSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Klasifikasi");
  var klasData = klasSheet.getDataRange().getValues();
  var targetFolderId = null;

  for (var k = 1; k < klasData.length; k++) {
    if (klasData[k][0].toString().trim().toUpperCase() === newBusinessActivity) {
      var folderUrl = klasData[k][2];
      if (folderUrl) {
        var match = folderUrl.match(/[-\w]{25,}/); 
        if (match) targetFolderId = match[0];
      }
      break;
    }
  }

  if (!targetFolderId) return { success: false, message: "Error: No Folder found!" };

  try {
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    let isOldFileDrive = oldFileUrl && oldFileUrl.toString().includes("drive.google.com");

    if (dataObj.fileData && dataObj.fileData.base64) {
      if (isOldFileDrive) {
        const match = oldFileUrl.match(/[-\w]{25,}/);
        if (match) { try { DriveApp.getFileById(match[0]).setTrashed(true); } catch(e){} }
      }
      var ext = "";
      if (dataObj.fileData.fileName.lastIndexOf(".") !== -1) ext = dataObj.fileData.fileName.substring(dataObj.fileData.fileName.lastIndexOf("."));
      var blob = Utilities.newBlob(Utilities.base64Decode(dataObj.fileData.base64), dataObj.fileData.mimeType, cleanRecordId + ext);
      var newFile = targetFolder.createFile(blob);
      
      // LOGIKA KEAMANAN: Cek tingkat kerahasiaan
      if (dataObj.confLevel === "Public") {
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } else {
        newFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      }
      
      finalFileUrl = newFile.getUrl(); 
      
    } else if (dataObj.manualLink && dataObj.manualLink.trim() !== "") {
      if (isOldFileDrive && dataObj.manualLink !== oldFileUrl) {
        const match = oldFileUrl.match(/[-\w]{25,}/);
        if (match) { try { DriveApp.getFileById(match[0]).setTrashed(true); } catch(e){} }
      }
      finalFileUrl = dataObj.manualLink.trim();
      
    } else {
      if (isOldFileDrive) {
        const match = oldFileUrl.match(/[-\w]{25,}/);
        if (match) {
          const existingFile = DriveApp.getFileById(match[0]);
          if (oldBusinessActivity !== newBusinessActivity) existingFile.moveTo(targetFolder);
          if (oldRecordId !== cleanRecordId) {
            var ext = "";
            if (existingFile.getName().lastIndexOf(".") !== -1) ext = existingFile.getName().substring(existingFile.getName().lastIndexOf("."));
            existingFile.setName(cleanRecordId + ext);
          }
          
          // UPDATE KEAMANAN: Ubah hak akses file lama sesuai dengan opsi yang baru dipilih
          if (dataObj.confLevel === "Public") {
            existingFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } else {
            existingFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
          }
        }
      } else { finalFileUrl = ""; }
    }
  } catch (error) { return { success: false, message: "Drive Error: " + error.message }; }

  var currentUser = Session.getActiveUser().getEmail() || "System User";
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
  
  const range = sheet.getRange(row, 1, 1, 18);
  range.setValues([[
    cleanRecordId, dataObj.title, dataObj.businessActivity, dataObj.creator, dataObj.unit, 
    dataObj.dateCreated, dataObj.dateReceived, dataObj.recordType, dataObj.confLevel, 
    finalFileUrl, dataObj.format, dataObj.status, dataObj.pageCount || "", dataObj.box || "", 
    dataObj.rack || "", currentUser, timestamp, dataObj.physicalPage || "" 
  ]]);
  
  return { success: true, message: "Success! Updated Arsip Record." };
}

function deleteArsipRow(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arsip");
  const fileUrl = sheet.getRange(row, 10).getValue(); 
  if (fileUrl && fileUrl.toString().includes("drive.google.com")) {
    try {
      const match = fileUrl.match(/[-\w]{25,}/);
      if (match) DriveApp.getFileById(match[0]).setTrashed(true); 
    } catch (e) {}
  }
  sheet.deleteRow(row);
  return "Success!";
}

// ==========================================
// AUDIT TRAIL / TRACEABILITY LOG
// ==========================================
function logTrace(actorName, actionDetail) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trace");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Trace");
    sheet.appendRow(["Timestamp", "User", "Action"]);
    sheet.getRange("A1:C1").setFontWeight("bold"); 
  }
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([timestamp, actorName, actionDetail]);
}

// ==========================================
// WAREHOUSE / LOKASI FISIK MANAGEMENT
// ==========================================
function updateArsipLokasi(oldRack, oldBox, newRack, newBox, action) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arsip");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var currentRack = data[i][14] != null ? data[i][14].toString().trim() : "";
    var currentBox = data[i][13] != null ? data[i][13].toString().trim() : "";
    if (action === 'renameRack' && currentRack === oldRack) {
      sheet.getRange(i + 1, 15).setValue(newRack);
    } else if (action === 'deleteRack' && currentRack === oldRack) {
      sheet.getRange(i + 1, 15).setValue("");
      sheet.getRange(i + 1, 14).setValue(""); 
    } else if (action === 'renameBox' && currentRack === oldRack && currentBox === oldBox) {
      sheet.getRange(i + 1, 14).setValue(newBox);
    } else if (action === 'deleteBox' && currentRack === oldRack && currentBox === oldBox) {
      sheet.getRange(i + 1, 14).setValue("");
    }
  }
  return { success: true };
}
