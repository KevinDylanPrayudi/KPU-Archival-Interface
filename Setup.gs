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
  
  // Hapus Sheet "Sheet1" bawaan Google jika masih ada dan kosong
  let sheetBawaan = ss.getSheetByName("Sheet1");
  if (sheetBawaan && ss.getSheets().length > 1) {
    ss.deleteSheet(sheetBawaan);
    Logger.log("🧹 Sheet bawaan 'Sheet1' berhasil dihapus.");
  }

  Logger.log("🎉 SETUP DATABASE SELESAI! Semua struktur tabel siap digunakan.");
}
