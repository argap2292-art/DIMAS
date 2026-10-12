# 📄 Panduan & Kode Backend Google Apps Script (Code.gs)
## Aplikasi Pembiasaan Salat Fardu DIMAS - SDN 1 Candi Mlarak Ponorogo

Ikuti langkah-langkah mudah di bawah ini untuk menghubungkan aplikasi **DIMAS** ke Google Spreadsheet milik Bapak Guru.

---

### 1️⃣ Salin Kode Script Berikut ke Google Apps Script

Buka Google Spreadsheet baru (contoh: *"Rekap Pembiasaan Salat DIMAS SDN 1 Candi"*), lalu klik menu:
**Ekstensi (Extensions)** &rarr; **Apps Script**.

Hapus semua kode di dalam berkas `Code.gs`, lalu tempelkan kode di bawah ini:

```javascript
/**
 * =========================================================================
 * BACKEND GOOGLE APPS SCRIPT (Code.gs) - APLIKASI DIMAS
 * SD NEGERI 1 CANDI, MLARAK, PONOROGO
 * =========================================================================
 */

// OPTIONAL: Masukkan ID Spreadsheet Anda jika menggunakan script mandiri (standalone).
// Jika script ini dibuat dari menu Ekstensi > Apps Script di Spreadsheet, biarkan kosong ("").
const SPREADSHEET_ID = ""; 
const SHEET_NAME = "Rekap_Salat_DIMAS";

/**
 * 1. FUNGSI SETUP AWAL (JALANKAN 1 KALI DI AWAL)
 * Fungsi ini membuat sheet baru dan baris header kolom secara otomatis.
 */
function initialSetup() {
  var ss = (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") 
    ? SpreadsheetApp.openById(SPREADSHEET_ID) 
    : SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Definisi Header Kolom
  var headers = [
    "Timestamp",
    "Nama Siswa",
    "Kelas",
    "Jenis Kelamin",
    "Tanggal",
    "Salat Subuh",
    "Salat Zuhur",
    "Salat Asar",
    "Salat Magrib",
    "Salat Isya",
    "Total Poin",
    "Keterangan/Catatan"
  ];

  // Tulis Header di baris pertama
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format Header Cantik Islami (Warna Hijau & Teks Putih Tebal)
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#15803d")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
    
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log("✅ Berhasil! Header tabel telah dibuat pada sheet: " + SHEET_NAME);
}

/**
 * 2. FUNGSI PENERIMA DATA (doPost)
 * Menerima kiriman data JSON dari formulir aplikasi DIMAS peserta didik.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  // Cegah benturan data jika banyak siswa mengirim bersamaan
  lock.tryLock(15000);

  try {
    var ss = (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") 
      ? SpreadsheetApp.openById(SPREADSHEET_ID) 
      : SpreadsheetApp.getActiveSpreadsheet();

    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      initialSetup();
      sheet = ss.getSheetByName(SHEET_NAME);
    }

    // Parsing data JSON yang dikirim frontend
    var contents = e.postData.contents;
    var data = JSON.parse(contents);

    // Format tanggal dan waktu kirim (WIB)
    var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

    // Susun baris baru data siswa (Mendukung namaSiswa & keterangan)
    var newRow = [
      timestamp,
      data.namaSiswa || data.nama || "",
      data.kelas || "",
      data.gender || "",
      data.tanggal || "",
      data.subuh || "Tidak Salat",
      data.zuhur || "Tidak Salat",
      data.asar || "Tidak Salat",
      data.magrib || "Tidak Salat",
      data.isya || "Tidak Salat",
      data.totalPoin !== undefined ? Number(data.totalPoin) : 0,
      data.keterangan || data.catatan || ""
    ];

    sheet.appendRow(newRow);

    // Kirim respons sukses dengan ContentService & Header CORS
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "success",
        message: "Laporan salat peserta didik berhasil dicatat ke Google Sheets.",
        nama: data.namaSiswa || data.nama,
        tanggal: data.tanggal
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("Error doPost: " + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

/**
 * 3. FUNGSI PENGAMBILAN & PENGUJIAN DATA (doGet)
 * Mengembalikan data rekapitulasi salat siswa ke Dashboard Guru DIMAS secara otomatis (JSON),
 * atau status aktif saat script diakses via peramban.
 */
function doGet(e) {
  try {
    var ss = (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") 
      ? SpreadsheetApp.openById(SPREADSHEET_ID) 
      : SpreadsheetApp.getActiveSpreadsheet();

    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: "active",
          app: "DIMAS - Pembiasaan Salat Fardu SDN 1 Candi",
          message: "Web App aktif. Sheet belum dibuat, silakan jalankan initialSetup().",
          count: 0,
          reports: []
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: "active",
          app: "DIMAS - Pembiasaan Salat Fardu SDN 1 Candi",
          message: "Web App aktif dan siap menerima data. Belum ada entri siswa.",
          count: 0,
          reports: []
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var reports = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (!row[1] && !row[4]) continue; // Lewati jika baris kosong
      
      var tglStr = "";
      if (row[4] instanceof Date) {
        tglStr = Utilities.formatDate(row[4], "Asia/Jakarta", "yyyy-MM-dd");
      } else {
        tglStr = String(row[4] || "").trim();
      }

      reports.push({
        timestamp: row[0],
        namaSiswa: String(row[1] || "").trim(),
        nama: String(row[1] || "").trim(),
        kelas: String(row[2] || "").trim(),
        gender: String(row[3] || "").trim(),
        tanggal: tglStr,
        subuh: String(row[5] || "Tidak Salat").trim(),
        zuhur: String(row[6] || "Tidak Salat").trim(),
        asar: String(row[7] || "Tidak Salat").trim(),
        magrib: String(row[8] || "Tidak Salat").trim(),
        isya: String(row[9] || "Tidak Salat").trim(),
        totalPoin: Number(row[10]) || 0,
        keterangan: String(row[11] || "-").trim(),
        catatan: String(row[11] || "-").trim()
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "success",
        app: "DIMAS - Pembiasaan Salat Fardu SDN 1 Candi",
        message: "Berhasil mengambil data laporan salat peserta didik.",
        count: reports.length,
        reports: reports
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        app: "DIMAS - Pembiasaan Salat Fardu SDN 1 Candi",
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

---

### 2️⃣ URL Web App Resmi yang Telah Terpasang di Frontend:

URL Web App resmi Bapak Guru telah terintegrasi di berkas `index.html`:
```javascript
const GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbww4JvG5r72yjRqmKiPei_H8oB5DibI_UKKPvThdC9VVSM0CtLYJ43Gg5KkayytmpEfjA/exec";
```
- **Pengiriman Siswa (POST)**: Setiap kali siswa mengisi laporan salat di form DIMAS, data dikirim secara *real-time* ke URL ini (`mode: 'no-cors'`) dan otomatis masuk ke Google Spreadsheet Bapak Guru.
- **Pengambilan Guru (GET)**: Dashboard Guru secara otomatis mengambil data rekapitulasi dari Google Sheets via `fetch(GOOGLE_APPS_SCRIPT_URL)` sehingga Bapak Guru selalu melihat data teranyar.

