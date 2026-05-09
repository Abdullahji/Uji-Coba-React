/**
 * TOKO WINA - GOOGLE APPS SCRIPT BACKEND (FINAL VERSION)
 * Sistem Inventarisasi & Cek Harga Terintegrasi
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('TOKO WINA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// Helper: Format Tanggal ke String (DD/MM/YYYY)
function formatDate(date) {
  if (!(date instanceof Date)) return date;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

/**
 * AMBIL DATA MASTER CATEGORY
 */
function ambilMasterCategory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master Category Brg');
    if (!sheet) return { sukses: false, pesan: 'Sheet "Master Category Brg" tidak ditemukan!' };
    
    const rows = sheet.getDataRange().getValues();
    const categories = rows.slice(1).map(row => row[0]).filter(c => c && c.toString().trim() !== "");
    return { sukses: true, data: [...new Set(categories)].sort() };
  } catch (e) {
    return { sukses: false, pesan: e.toString() };
  }
}

/**
 * AMBIL DATA SUPPLIER
 */
function ambilDataSupplier() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Supplier');
    
    // Auto-create sheet if missing (helpful for first-time use)
    if (!sheet) {
      sheet = ss.insertSheet('Supplier');
      sheet.appendRow([
        "No.", "Tanggal Pembelian", "Supplier", "Nama Barang", "Kategori Barang", 
        "Barcode Ball", "Barcode Pak", "Barcode Bks", "Harga Ball", "Isi Ball", 
        "Harga Pak", "Isi Pak", "Harga Bks", "Harga Naik", "Expired Date"
      ]);
    }
    
    const dataRange = sheet.getDataRange();
    const rows = dataRange.getValues();
    const displayRows = dataRange.getDisplayValues();
    
    if (rows.length <= 1) return { sukses: true, data: [] }; 

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const displayRow = displayRows[i];
      
      // Syarat data valid: Minimal ada Nama Barang
      if (!row[3]) continue;
      
      data.push({
        no: displayRow[0] || i,
        tglBeli: displayRow[1] || "",
        supplier: displayRow[2] || "-",
        namaBarang: displayRow[3] || "-",
        categoryBarang: displayRow[4] || "-",
        barcodeBall: displayRow[5] || "",
        barcodePak: displayRow[6] || "",
        barcodeBks: displayRow[7] || "",
        hargaBall: parseFloat(row[8]) || 0,
        isiBall: parseInt(row[9]) || 1,
        hargaPak: parseFloat(row[10]) || 0,
        isiPak: parseInt(row[11]) || 1,
        hargaBks: parseFloat(row[12]) || 0,
        hargaNaik: displayRow[13] || "Tidak",
        expiredDate: displayRow[14] || "-",
        id: "S-" + i // Index row data (1-based dari baris ke-2)
      });
    }
    return { sukses: true, data: data.reverse() };
  } catch (e) {
    return { sukses: false, pesan: "Error Ambil Supplier: " + e.toString() };
  }
}

/**
 * AMBIL DATA CEK HARGA
 */
function ambilDataCekHarga() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Cek Harga Barang');
    if (!sheet) return { sukses: false, pesan: 'Sheet tidak ditemukan!' };
    
    const rows = sheet.getDataRange().getValues();
    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[1]) continue; 
      data.push({
        barcode: row[0] ? row[0].toString() : '',
        namaBarang: row[1],
        hargaBarang: Number(row[2]) || 0,
        gambarBarang: row[3] || '',
        id: "C-" + i
      });
    }
    return { sukses: true, data: data };
  } catch (e) {
    return { sukses: false, pesan: e.toString() };
  }
}

/**
 * SIMPAN DATA
 */
function simpanKeSupplier(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSupplier = ss.getSheetByName('Supplier');
    if (!sheetSupplier) throw new Error('Sheet "Supplier" tidak ditemukan!');
    
    let tgl = payload.tglBeli ? new Date(payload.tglBeli) : new Date();
    if (isNaN(tgl.getTime())) tgl = new Date();
    tgl.setHours(0, 0, 0, 0);

    let exp = "-";
    if (payload.expiredDate) {
      const d = new Date(payload.expiredDate);
      if (!isNaN(d.getTime())) {
        d.setHours(0, 0, 0, 0);
        exp = d;
      }
    }
    
    const rowData = [
      payload.no || "",                     // A
      tgl,                                  // B
      payload.supplier || "-",               // C
      payload.namaBarang || "-",             // D
      payload.categoryBarang || "-",         // E
      payload.barcodeBall || "",             // F
      payload.barcodePak || "",              // G
      payload.barcodeBks || "",              // H
      Number(payload.hargaBall) || 0,        // I
      Number(payload.isiBall) || 1,          // J
      Number(payload.hargaPak) || 0,         // K
      Number(payload.isiPak) || 1,           // L
      Number(payload.hargaBks) || 0,         // M
      payload.hargaNaik || "Tidak",           // N
      exp                                   // O
    ];

    if (payload.rowId && payload.rowId.startsWith("S-")) {
      const rowIndex = parseInt(payload.rowId.replace('S-', '')) + 1;
      sheetSupplier.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      const lastRow = sheetSupplier.getLastRow();
      if (!rowData[0]) rowData[0] = lastRow;
      sheetSupplier.appendRow(rowData);
    }

    return { sukses: true, pesan: 'Berhasil disimpan!' };
  } catch (e) {
    return { sukses: false, pesan: e.toString() };
  }
}

function syncToCekHarga(barcode, nama, harga) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Cek Harga Barang');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === barcode.toString()) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[nama, harga]]);
      return;
    }
  }
  sheet.appendRow([barcode, nama, harga, ""]);
}

/**
 * INIALISASI DATA DALAM SATU PANGGILAN (PERFORMANCE OPTIMIZATION)
 */
function ambilInisialisasiData() {
  return {
    sukses: true,
    supplier: ambilDataSupplier().data || [],
    cekHarga: ambilDataCekHarga().data || [],
    category: ambilMasterCategory().data || []
  };
}

/**
 * DISPATCHER TERPUSAT untuk google.script.run dari index.html
 * Diperlukan karena google.script.run tidak mendukung pemanggilan fungsi dinamis.
 */
function dispatchAction(params) {
  const action = params.action;
  const payload = params.payload;
  if (action === "ambilInisialisasiData") return ambilInisialisasiData();
  if (action === "ambilDataSupplier") return ambilDataSupplier();
  else if (action === "ambilDataCekHarga") return ambilDataCekHarga();
  else if (action === "simpanKeSupplier") return simpanKeSupplier(payload);
  else if (action === "hapusDataSupplier") return hapusDataSupplier(payload);
  else if (action === "ambilMasterCategory") return ambilMasterCategory();
  else return { sukses: false, pesan: "Action tidak dikenal: " + action };
}

function doPost(e) {
  try {
    // Membaca payload dari request body
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const payload = params.payload;
    
    let result;
    
    // Gunakan fungsi yang sama dengan dispatcher
    if (action === "ambilInisialisasiData") result = ambilInisialisasiData();
    else if (action === "ambilDataSupplier") result = ambilDataSupplier();
    else if (action === "ambilDataCekHarga") result = ambilDataCekHarga();
    else if (action === "simpanKeSupplier") result = simpanKeSupplier(payload);
    else if (action === "hapusDataSupplier") result = hapusDataSupplier(payload);
    else if (action === "ambilMasterCategory") result = ambilMasterCategory();
    else result = { sukses: false, pesan: "Action tidak dikenal: " + action };
    
    // Mengembalikan response sebagai JSON yang bisa dibaca Fetch dari luar
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      sukses: false, 
      pesan: "Backend Error: " + err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function hapusDataSupplier(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Supplier');
    if (!sheet) throw new Error('Sheet Supplier tidak ada!');
    // id = "S-{i}" dimana i adalah index loop (1-based dari baris data)
    // Baris di sheet = i + 1 (karena baris 1 adalah header)
    const i = parseInt(payload.id.replace('S-', ''));
    const rowIndex = i + 1; // +1 karena header di baris 1
    sheet.deleteRow(rowIndex);
    return { sukses: true, pesan: 'Dihapus!' };
  } catch (e) {
    return { sukses: false, pesan: 'Gagal: ' + e.toString() };
  }
}