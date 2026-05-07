// ==========================================
// 🛠️ DATABASE GENERATOR (Auto-Setup Sheets)
// ==========================================

function setupDatabaseInit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Sheet Pegawai
  let sheetPegawai = ss.getSheetByName("Pegawai");
  if (!sheetPegawai) {
    sheetPegawai = ss.insertSheet("Pegawai");
    sheetPegawai.appendRow(["Nama Lengkap", "Email", "Kata Sandi", "Tanggal Dibuat"]);
    sheetPegawai.getRange("A1:D1").setFontWeight("bold");
    sheetPegawai.setFrozenRows(1);
    Logger.log("✅ Sheet 'Pegawai' berhasil dibuat.");
  } else {
    Logger.log("⚠️ Sheet 'Pegawai' sudah ada.");
  }
  
  // 2. Setup Sheet Klasifikasi
  let sheetKlasifikasi = ss.getSheetByName("Klasifikasi");
  if (!sheetKlasifikasi) {
    sheetKlasifikasi = ss.insertSheet("Klasifikasi");
    sheetKlasifikasi.appendRow(["Kode Klasifikasi", "Nama Klasifikasi", "Link Folder Drive", "Tanggal Dibuat"]);
    sheetKlasifikasi.getRange("A1:D1").setFontWeight("bold");
    sheetKlasifikasi.setFrozenRows(1);
    Logger.log("✅ Sheet 'Klasifikasi' berhasil dibuat.");
  } else {
    Logger.log("⚠️ Sheet 'Klasifikasi' sudah ada.");
  }
  
  // 3. Setup Sheet Arsip (18 Kolom Sesuai Code.gs terbaru)
  let sheetArsip = ss.getSheetByName("Arsip");
  if (!sheetArsip) {
    sheetArsip = ss.insertSheet("Arsip");
    const arsipHeaders = [
      "Record ID", 
      "Judul Dokumen", 
      "Klasifikasi", 
      "Unit Pencipta", 
      "Unit Pengolah", 
      "Tanggal Dibuat", 
      "Tanggal Diterima", 
      "Jenis Arsip", 
      "Kerahasiaan", 
      "File Digital (Link)", 
      "Format", 
      "Status", 
      "Total Halaman", 
      "Boks/Map", 
      "Rak/Lemari", 
      "Pengunggah", 
      "Waktu Diunggah", 
      "Nomor Halaman Fisik"
    ];
    sheetArsip.appendRow(arsipHeaders);
    sheetArsip.getRange("A1:R1").setFontWeight("bold");
    sheetArsip.setFrozenRows(1); // Kunci baris pertama agar enak saat di-scroll
    Logger.log("✅ Sheet 'Arsip' berhasil dibuat dengan 18 kolom.");
  } else {
    Logger.log("⚠️ Sheet 'Arsip' sudah ada.");
  }
  
  // 4. Setup Sheet Trace (Audit Trail)
  let sheetTrace = ss.getSheetByName("Trace");
  if (!sheetTrace) {
    sheetTrace = ss.insertSheet("Trace");
    sheetTrace.appendRow(["Timestamp", "User", "Action"]);
    sheetTrace.getRange("A1:C1").setFontWeight("bold");
    sheetTrace.setFrozenRows(1);
    Logger.log("✅ Sheet 'Trace' berhasil dibuat.");
  } else {
    Logger.log("⚠️ Sheet 'Trace' sudah ada.");
  }
  
  // 5. Setup Sheet File_Permissions (Queue System)
  let sheetPerms = ss.getSheetByName("File_Permissions");
  if (!sheetPerms) {
    sheetPerms = ss.insertSheet("File_Permissions");
    sheetPerms.appendRow(["Record ID", "File URL", "Email", "Action", "Status", "Timestamp"]);
    sheetPerms.getRange("A1:F1").setFontWeight("bold");
    sheetPerms.setFrozenRows(1);
    Logger.log("✅ Sheet 'File_Permissions' berhasil dibuat.");
  } else {
    Logger.log("⚠️ Sheet 'File_Permissions' sudah ada.");
  }
  
  // Hapus Sheet "Sheet1" bawaan Google jika masih ada dan kosong
  let sheetBawaan = ss.getSheetByName("Sheet1");
  if (sheetBawaan && ss.getSheets().length > 1) {
    ss.deleteSheet(sheetBawaan);
    Logger.log("🧹 Sheet bawaan 'Sheet1' berhasil dihapus.");
  }

  Logger.log("🎉 SETUP DATABASE SELESAI! Semua struktur tabel siap digunakan.");
}

/**
 * Menambahkan kolom "Setup" di header terakhir pada semua sheet.
 */
function addSetupColumnToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const columnName = "Setup"; 

  let addedCount = 0;

  sheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    
    // Jika sheet kosong sama sekali
    if (lastCol === 0) {
      sheet.getRange(1, 1).setValue(columnName);
      addedCount++;
      return;
    }
    
    // Cek header yang sudah ada untuk mencegah duplikasi
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!headers.includes(columnName)) {
      sheet.getRange(1, lastCol + 1).setValue(columnName);
      addedCount++;
    }
  });

  Logger.log(`Selesai. Kolom '${columnName}' berhasil ditambahkan pada ${addedCount} sheet.`);
}

function addSetupColumnToEdge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    // Expand to the edge (last column + 1)
    if (lastCol === 0) {
      sheet.getRange(1, 1).setValue("Setup");
    } else {
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (!headers.includes("Setup")) {
        sheet.getRange(1, lastCol + 1).setValue("Setup");
        // Optional: formatting the new edge column
        sheet.getRange(1, lastCol + 1).setFontWeight("bold").setBackground("#f1f5f9");
      }
    }
  });
  Logger.log("Kolom 'Setup' berhasil ditambahkan di ujung kanan (edge) setiap sheet.");
}

// ==========================================
// 🛠️ DATABASE EXPANSION TO EDGE
// ==========================================
function expandDatabaseToEdge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const edgeColumnName = "Setup"; 

  let modifiedSheets = [];

  sheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    
    // Jika sheet kosong sama sekali
    if (lastCol === 0) {
      sheet.getRange(1, 1).setValue(edgeColumnName);
      sheet.getRange(1, 1).setFontWeight("bold").setBackground("#e2e8f0");
      modifiedSheets.push(sheet.getName());
      return;
    }
    
    // Cek header untuk mencegah duplikasi jika skrip dijalankan 2x
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!headers.includes(edgeColumnName)) {
      // Tambahkan di ujung kanan (Edge + 1)
      const edgeRange = sheet.getRange(1, lastCol + 1);
      edgeRange.setValue(edgeColumnName);
      edgeRange.setFontWeight("bold").setBackground("#e2e8f0");
      modifiedSheets.push(sheet.getName());
    }
  });

  if (modifiedSheets.length > 0) {
    Logger.log("✅ Sukses: Kolom 'Setup' ditambahkan di batas Edge pada sheet: " + modifiedSheets.join(", "));
  } else {
    Logger.log("⚠️ Info: Semua sheet sudah memiliki kolom 'Setup' di batas Edge.");
  }
}
