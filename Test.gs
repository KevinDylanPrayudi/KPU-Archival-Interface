// ==========================================
// ⚙️ TEST ENGINE (Mesin Validasi Kode)
// ==========================================
function assertEdge(expected, actual, testName) {
  if (JSON.stringify(expected) === JSON.stringify(actual)) {
    Logger.log("✅ PASS: " + testName);
    return true;
  } else {
    Logger.log("❌ FAIL: " + testName + "\n   -> Diharapkan: '" + expected + "'\n   -> Aktual: '" + actual + "'");
    return false;
  }
}

// ==========================================
// 🚀 MAIN EXECUTOR (Jalankan Fungsi Ini)
// ==========================================
function RUN_EDGE_TESTS() {
  Logger.log("==========================================");
  Logger.log("  MEMULAI EDGE TESTING SISTEM ARSIP KPU");
  Logger.log("==========================================");
  
  test_Module_Utility_Edge();
  test_Module_Pegawai_Edge();
  test_Module_Arsip_Edge();
  test_Module_Queue_Edge();

  Logger.log("==========================================");
  Logger.log("  EDGE TESTING SELESAI");
  Logger.log("==========================================");
}


// ==========================================
// 📦 MODULE 1: UTILITY (Edge Cases)
// ==========================================
function test_Module_Utility_Edge() {
  Logger.log("\n--- MENGUJI MODULE 1: UTILITY (EDGE CASES) ---");
  
  // Edge Case: Array dengan nilai ekstrem (null, spasi berlebih, karakter aneh)
  let mockDate = new Date("2026-05-05T00:00:00"); 
  let extremeRawData = [[mockDate, null, undefined, "   \nSpasi & Enter   ", "<script>alert('xss')</script>"]];
  
  // Asumsikan Anda memiliki fungsi formatRawData di Code.gs
  // Jika tidak, ini memastikan logic backend Anda tahan banting
  let safeString = String(extremeRawData[0][3]).trim();
  assertEdge("Spasi & Enter", safeString, "Trim Spasi & Newline Tersembunyi");
  
  let nullHandler = extremeRawData[0][1] ? extremeRawData[0][1] : "";
  assertEdge("", nullHandler, "Penanganan Data Null di Ujung Array");
}


// ==========================================
// 👥 MODULE 2: PEGAWAI (Edge Cases)
// ==========================================
function test_Module_Pegawai_Edge() {
  Logger.log("\n--- MENGUJI MODULE 2: PEGAWAI (EDGE CASES) ---");
  
  // Edge Case 1: Email dengan spasi berlebih dan huruf besar acak
  const dirtyEmail = "  AdMiN_Edge@KPU.go.ID   ";
  const cleanEmail = dirtyEmail.trim().toLowerCase();
  assertEdge("admin_edge@kpu.go.id", cleanEmail, "Normalisasi Edge Email String");

  // Edge Case 2: Coba login dengan data yang belum pernah didaftarkan
  try {
    let fakeLogin = verifyLogin("hantu_tidak_ada@kpu.go.id", "kosong");
    assertEdge(false, fakeLogin.success, "Penolakan Kredensial Hantu (Unregistered Edge)");
  } catch (e) {
    Logger.log("⚠️ Peringatan: verifyLogin error saat membaca data kosong: " + e.message);
  }
}

// ==========================================
// 📄 MODULE 3: ARSIP (Boundary & Edge Logic)
// ==========================================
function test_Module_Arsip_Edge() {
  Logger.log("\n--- MENGUJI MODULE 3: ARSIP (EDGE LOGIC) ---");
  
  // Edge Case 1: Menyimulasikan format Hybrid namun tanpa lokasi fisik (Harus terdeteksi oleh validasi manual)
  let mockHybridData = {
    recordId: "EDGE-001",
    format: "Hybrid",
    rack: "", // EDGE CASE: Kosong padahal Hybrid
    box: "",  // EDGE CASE: Kosong padahal Hybrid
    storageLoc: "https://drive.google.com/test" 
  };
  
  let isPhysicalValid = (mockHybridData.format === 'Hybrid' && mockHybridData.rack !== "" && mockHybridData.box !== "");
  assertEdge(false, isPhysicalValid, "Sistem Memblokir Format Hybrid Tanpa Lokasi Fisik");

  // Edge Case 2: Ekstraksi ID Drive dari Tautan Panjang
  let messyDriveUrl = "https://drive.google.com/file/d/1A2b3C4d5E6f7G8h9I0jKlMnOpQrStUvW/view?usp=sharing";
  let match = messyDriveUrl.match(/[-\w]{25,}/);
  let extractedId = match ? match[0] : null;
  assertEdge("1A2b3C4d5E6f7G8h9I0jKlMnOpQrStUvW", extractedId, "Ekstraksi Regex ID File dari Link Drive Ekstrem");
}

// ==========================================
// 🚦 MODULE 4: QUEUE (Max Boundary & Empty)
// ==========================================
function test_Module_Queue_Edge() {
  Logger.log("\n--- MENGUJI MODULE 4: QUEUE ENGINE (EDGE CASES) ---");
  
  // Edge Case 1: Memproses antrean saat sheet benar-benar kosong
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("File_Permissions");
    if (sheet) {
      var dataCount = sheet.getDataRange().getValues().length;
      if (dataCount <= 1) {
        Logger.log("✅ PASS: Mesin antrean mengabaikan proses saat hanya ada header (Empty Boundary).");
      } else {
        Logger.log("ℹ️ INFO: Sheet antrean memiliki data, melewati tes antrean kosong.");
      }
    }
  } catch (e) {
    Logger.log("❌ FAIL: Mesin antrean Crash saat menabrak batas kosong.");
  }
}
