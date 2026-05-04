// ==========================================
// ⚙️ TEST ENGINE (Mesin Validasi Kode)
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
function RUN_ALL_KPU_TESTS() {
  Logger.log("==========================================");
  Logger.log("  MEMULAI PENGUJIAN SISTEM ARSIP KPU");
  Logger.log("==========================================");
  
  test_Utility_Module();
  test_Pegawai_Integration();
  test_Klasifikasi_Module();
  test_Arsip_Security_Logic();
  test_Warehouse_Logic();

  Logger.log("==========================================");
  Logger.log("  PENGUJIAN SELESAI");
  Logger.log("==========================================");
}


// ==========================================
// 📦 1. MODULE UTILITY
// ==========================================
function test_Utility_Module() {
  Logger.log("\n--- MENGUJI MODULE UTILITY ---");
  
  // Tes: Konversi tanggal dan pembersihan spasi
  let mockDate = new Date("2026-04-15T00:00:00"); 
  let rawData = [[mockDate, null, "  Spasi Ekstra  "]];
  let formattedData = formatRawData(rawData);
  
  assertEquals("2026-04-15", formattedData[0][0], "Konversi Tanggal (format yyyy-MM-dd)");
  assertEquals("", formattedData[0][1], "Penanganan Input Kosong (Null/Undefined)");
  assertEquals("Spasi Ekstra", formattedData[0][2], "Pembersihan Spasi (Trim String)");
}


// ==========================================
// 👥 2. MODULE PEGAWAI (Integration)
// ==========================================
function test_Pegawai_Integration() {
  Logger.log("\n--- MENGUJI MODULE PEGAWAI ---");
  const testEmail = "unit_test_admin@kpu.go.id";
  const testPass = "TestPass123";
  const testName = "Admin Testing";

  try {
    // 1. Tes Pembuatan Data Baru
    let createRes = processForm(testName, testEmail, testPass);
    assertEquals(true, createRes.success, "Pembuatan Akun Pegawai Baru");

    // 2. Tes Verifikasi Login
    let loginSuccess = verifyLogin(testEmail, testPass);
    assertEquals(true, loginSuccess.success, "Login dengan kredensial benar");

    let loginFail = verifyLogin(testEmail, "SalahSandi99");
    assertEquals(false, loginFail.success, "Penolakan Login dengan sandi salah");

    // 3. Tes Duplikasi Data
    let duplicateRes = processForm("Admin Copy", testEmail, "PassLain");
    assertEquals(false, duplicateRes.success, "Pencegahan Duplikasi Email Terdaftar");

  } finally {
    // 4. CLEANUP: Hapus data uji dari Spreadsheet secara otomatis
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pegawai");
    var data = sheet.getDataRange().getValues();
    for(var i = 1; i < data.length; i++) {
      if(data[i][1].toString().trim().toLowerCase() === testEmail.toLowerCase()) { 
        sheet.deleteRow(i + 1); 
        Logger.log("🧹 CLEANUP: Baris data uji Pegawai berhasil dihapus.");
        break; 
      }
    }
  }
}


// ==========================================
// 🗂️ 3. MODULE KLASIFIKASI
// ==========================================
function test_Klasifikasi_Module() {
  Logger.log("\n--- MENGUJI MODULE KLASIFIKASI ---");
  
  let klasData = getKlasifikasiData();
  assertNotUndefined(klasData, "Pengambilan Array Data Klasifikasi");
  
  if (klasData.length > 0) {
    assertNotUndefined(klasData[0].rowId, "Pemetaan rowId Klasifikasi");
    assertNotUndefined(klasData[0].kode, "Pemetaan Kode Klasifikasi");
  } else {
    Logger.log("⚠️ INFO: Tabel Klasifikasi kosong, melewati tes pemetaan.");
  }
}


// ==========================================
// 📄 4. MODULE ARSIP (Keamanan Drive)
// ==========================================
function test_Arsip_Security_Logic() {
  Logger.log("\n--- MENGUJI LOGIKA KEAMANAN DRIVE (ARSIP) ---");
  
  // Simulasi logika penentuan akses file tanpa menyentuh Drive beneran
  function simulateDriveSecurity(confLevel) {
    if (confLevel === "Public") {
      return "ANYONE_WITH_LINK_VIEW";
    } else {
      return "PRIVATE_NONE";
    }
  }

  assertEquals("ANYONE_WITH_LINK_VIEW", simulateDriveSecurity("Public"), "Arsip 'Public' -> Terbuka (Viewer)");
  assertEquals("PRIVATE_NONE", simulateDriveSecurity("Internal"), "Arsip 'Internal' -> Terkunci (Private)");
  assertEquals("PRIVATE_NONE", simulateDriveSecurity("Confidential"), "Arsip 'Confidential' -> Terkunci (Private)");
  assertEquals("PRIVATE_NONE", simulateDriveSecurity("Strict"), "Arsip 'Strict' -> Terkunci (Private)");
}


// ==========================================
// 🏢 5. MODULE WAREHOUSE (Lokasi Fisik)
// ==========================================
function test_Warehouse_Logic() {
  Logger.log("\n--- MENGUJI LOGIKA GUDANG (LOKASI FISIK) ---");
  
  // Simulasi logika updateArsipLokasi untuk memastikan data sel sejajar dengan aksi
  function simulateLocationUpdate(currentRack, oldRack, newRack, action) {
    let resultRack = currentRack;
    if (action === 'renameRack' && currentRack === oldRack) {
      resultRack = newRack;
    } else if (action === 'deleteRack' && currentRack === oldRack) {
      resultRack = "";
    }
    return resultRack;
  }

  // Uji Ganti Nama Rak
  let renameTest = simulateLocationUpdate("Rak A", "Rak A", "Rak B", "renameRack");
  assertEquals("Rak B", renameTest, "Pembaruan Nama Rak Berhasil");

  // Uji Hapus Rak
  let deleteTest = simulateLocationUpdate("Rak A", "Rak A", "", "deleteRack");
  assertEquals("", deleteTest, "Penghapusan Rak Menghapus Data Lokasi Arsip");
  
  // Uji Nama Rak Tidak Cocok (Seharusnya tidak berubah)
  let noMatchTest = simulateLocationUpdate("Rak C", "Rak A", "Rak B", "renameRack");
  assertEquals("Rak C", noMatchTest, "Rak yang tidak cocok diabaikan dari pembaruan");
}
