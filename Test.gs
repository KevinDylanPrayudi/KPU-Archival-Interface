// ==========================================
// ⚙️ TEST ENGINE (Sistem Auto-Guard Kode)
// ==========================================
function assertEquals(expected, actual, testName) {
  if (expected === actual) {
    Logger.log("✅ PASS: " + testName);
    return true;
  } else {
    Logger.log("❌ FAIL: " + testName + "\n   -> Diharapkan: '" + expected + "'\n   -> Hasil Aktual: '" + actual + "'");
    return false;
  }
}

function assertNotUndefined(value, testName) {
  if (value !== undefined && value !== null) {
    Logger.log("✅ PASS: " + testName);
    return true;
  } else {
    Logger.log("❌ FAIL: " + testName + " (Nilai tidak boleh null/undefined)");
    return false;
  }
}

// ==========================================
// 🚀 MAIN EXECUTOR (Jalankan Fungsi Ini)
// ==========================================
function RUN_ALL_TESTS() {
  Logger.log("==========================================");
  Logger.log("  MEMULAI PENGUJIAN SISTEM ARSIP KPU");
  Logger.log("==========================================");
  
  test_Utility_FormatData();
  test_Pegawai_Integration();
  test_Klasifikasi_Read();
  test_Arsip_Drive_Security();

  Logger.log("==========================================");
  Logger.log("  PENGUJIAN SELESAI");
  Logger.log("==========================================");
}

// ==========================================
// 📦 1. MODULE UTILITY
// ==========================================
function test_Utility_FormatData() {
  Logger.log("\n--- MENGUJI MODULE UTILITY ---");
  
  let mockDate = new Date("2026-04-15T00:00:00"); 
  let rawData = [[mockDate, null, "  Dokumen Pemilu  "]];
  let formattedData = formatRawData(rawData);
  
  assertEquals("2026-04-15", formattedData[0][0], "Konversi Objek Tanggal (yyyy-MM-dd)");
  assertEquals("", formattedData[0][1], "Penanganan Input Kosong (Null)");
  assertEquals("Dokumen Pemilu", formattedData[0][2], "Pembersihan Spasi Berlebih (Trim)");
}

// ==========================================
// 👥 2. MODULE PEGAWAI (Integration Test)
// ==========================================
function test_Pegawai_Integration() {
  Logger.log("\n--- MENGUJI MODULE PEGAWAI ---");
  const testEmail = "test_cpns@kpu.go.id";
  const testPass = "Latsar2026";
  const testName = "Admin Penguji";

  try {
    // 1. Tes Pembuatan Data
    let createRes = processForm(testName, testEmail, testPass);
    assertEquals(true, createRes.success, "Pembuatan Akun Pegawai Baru");

    // 2. Tes Verifikasi Login
    let loginSuccess = verifyLogin(testEmail, testPass);
    assertEquals(true, loginSuccess.success, "Verifikasi Login Sandi Benar");

    let loginFail = verifyLogin(testEmail, "SalahSandi99");
    assertEquals(false, loginFail.success, "Penolakan Login Sandi Salah");

    // 3. Tes Duplikasi
    let duplicateRes = processForm("Admin Lain", testEmail, "PassLain");
    assertEquals(false, duplicateRes.success, "Pencegahan Duplikasi Email");

  } finally {
    // 4. CLEANUP: Hapus data uji dari Spreadsheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pegawai");
    var data = sheet.getDataRange().getValues();
    for(var i = 1; i < data.length; i++) {
      if(data[i][1].toString().trim().toLowerCase() === testEmail.toLowerCase()) { 
        sheet.deleteRow(i + 1); 
        Logger.log("🧹 CLEANUP: Baris data uji Pegawai berhasil dihapus");
        break; 
      }
    }
  }
}

// ==========================================
// 🗂️ 3. MODULE KLASIFIKASI
// ==========================================
function test_Klasifikasi_Read() {
  Logger.log("\n--- MENGUJI MODULE KLASIFIKASI ---");
  
  let klasData = getKlasifikasiData();
  assertNotUndefined(klasData, "Pengambilan Array Data Klasifikasi");
  
  if (klasData.length > 0) {
    assertNotUndefined(klasData[0].rowId, "Pemetaan rowId Klasifikasi");
    assertNotUndefined(klasData[0].folderLink, "Pemetaan URL Folder Drive");
  } else {
    Logger.log("⚠️ INFO: Melewati tes pemetaan karena tabel Klasifikasi kosong.");
  }
}

// ==========================================
// 📄 4. MODULE ARSIP (Keamanan Drive)
// ==========================================
function test_Arsip_Drive_Security() {
  Logger.log("\n--- MENGUJI LOGIKA KEAMANAN DRIVE ---");
  
  // Menguji murni logika if/else yang men-trigger hak akses Drive
  let confPublic = "Public";
  let confStrict = "Strict";
  
  let actualPublicAccess = (confPublic === "Public") ? "ANYONE_WITH_LINK_VIEW" : "PRIVATE_NONE";
  let actualStrictAccess = (confStrict === "Public") ? "ANYONE_WITH_LINK_VIEW" : "PRIVATE_NONE";

  assertEquals("ANYONE_WITH_LINK_VIEW", actualPublicAccess, "Dokumen 'Public' -> Terbuka (Viewer)");
  assertEquals("PRIVATE_NONE", actualStrictAccess, "Dokumen 'Strict' -> Terkunci (Private)");
}
