// ===============================================================
// KONFIGURASI TERPUSAT
// ===============================================================
const SPREADSHEET_ID = "115M1ltaKNySPWRORMwJ6sQfWLJuGOdw-qAKLx_vfUeY"; // Pastikan ID ini sudah benar
const TIMEZONE = "Asia/Jakarta";


// ===============================================================
// FUNGSI UTAMA (WEB APP)
// ===============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Toto Aluminium Manufacture')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}


// ===============================================================
// FUNGSI HELPER INTERNAL (OPTIMASI PERFORMA)
// ===============================================================
/**
 * In-memory cache for sheet data during a single request execution.
 * Membantu menghindari pembacaan Google Sheet yang sama berulang kali.
 */
const DATA_CACHE = {}; 

/**
 * Mengambil data (headers dan values) dari sheet tertentu.
 * Menggunakan cache in-memory untuk menghindari pembacaan berulang dalam satu eksekusi.
 */
function _getSheetData(sheetName) {
    if (DATA_CACHE[sheetName]) {
        return DATA_CACHE[sheetName];
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet || sheet.getLastRow() < 2) {
        DATA_CACHE[sheetName] = { headers: [], values: [] };
        return DATA_CACHE[sheetName];
    }
    
    // Mengambil semua data
    const dataRange = sheet.getDataRange();
    const allValues = dataRange.getValues();
    const headers = allValues.shift(); // Baris pertama adalah header
    const values = allValues; // Sisanya adalah data

    DATA_CACHE[sheetName] = { headers, values };
    return DATA_CACHE[sheetName];
}

/**
 * Mengubah baris data array menjadi objek JavaScript.
 */
function _rowToObject(row, headers, rowIndex) {
    let obj = { rowNumber: rowIndex };
    headers.forEach((header, i) => {
        let value = row[i];
        if (value instanceof Date) {
            obj[header] = Utilities.formatDate(value, TIMEZONE, "dd/MM/yyyy");
        } else {
            obj[header] = value;
        }
    });
    return obj;
}


// ===============================================================
// API UNTUK SISI KLIEN (FRONT-END)
// ===============================================================

// --- FUNGSI LOGIN ---
function checkLogin(username, password) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Pengguna');

    if (sheet.getLastRow() < 2) return null;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (const row of data) {
      if (row[0] === username && row[1] === password) {
        return { username: row[0], role: row[2] };
      }
    }
    return null;
  } catch (e) {
    throw new Error(`Error pada checkLogin: ${e.message}`);
  }
}

// --- FUNGSI WORK ORDER ---
function getWorkOrders(month, year) {
    try {
        // MODIFIKASI: Gunakan helper _getSheetData (Optimasi Performa)
        const { headers, values } = _getSheetData('WorkOrders');
        
        if (values.length === 0) return [];

        const monthIndex = headers.indexOf('Bulan');
        const yearIndex = headers.indexOf('Tahun');
        if (monthIndex === -1 || yearIndex === -1) {
            throw new Error('Kolom "Bulan" atau "Tahun" tidak ditemukan di WorkOrders.');
        }

        const filterMonth = parseInt(month, 10);
        const filterYear = parseInt(year, 10);
        const data = [];
        values.forEach((row, index) => {
            const rowMonth = parseInt(row[monthIndex], 10);
            const rowYear = parseInt(row[yearIndex], 10);
            
            if (rowMonth === filterMonth && rowYear === filterYear) {
                // index + 2 karena header (baris 1) sudah di-shift, dan index array dimulai dari 0 (baris 2)
                data.push(_rowToObject(row, headers, index + 2)); 
            }
        });
        return data;
    } catch (e) {
        throw new Error(`Error di getWorkOrders: ${e.message}`);
    }
}

function addWorkOrder(orderData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('WorkOrders');
    
    // PERBAIKAN: Menambahkan kolom 'DRAFT' dan memastikan semua 17 kolom terisi.
    const newRow = [
      orderData.Tanggal ? new Date(orderData.Tanggal) : null, // 0
      orderData['Nama Customer'] || '', // 1
      orderData.Deskripsi || '', // 2
      orderData.Ukuran || '', // 3
      orderData.Qty || '', // 4
      orderData.Harga || '', // 5
      orderData['NO INV'] || '', // 6
      false, // Di Produksi (7) - Boolean
      false, // Di Warna (8) - Boolean
      false, // Siap Kirim (9) - Boolean
      false, // Di Kirim (10) - Boolean
      false, // Pembayaran (11) - Boolean
      '', // Ekspedisi (12) - String
      orderData.Bulan || '', // 13
      orderData.Tahun || '', // 14
      'DRAFT', // PO Status (15) - DIUBAH menjadi STRING untuk menghindari Sheet menimpanya ke FALSE
      '', // NoSJWarna (16) - Kolom baru untuk penyelarasan
    ];
    
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    
    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRowValues = sheet.getRange(lastRow, 1, 1, headers.length).getValues()[0];
    return { status: 'success', data: _rowToObject(newRowValues, headers, lastRow) };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function updateWorkOrder(rowNumber, orderData) {
    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName('WorkOrders');
        const range = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn());
        const existingValues = range.getValues()[0];
        // Pastikan array updatedRow memiliki panjang 17 (Kolom A-Q)
        const updatedRow = [
            orderData.Tanggal ? new Date(orderData.Tanggal) : existingValues[0],
            orderData['Nama Customer'] !== undefined ? orderData['Nama Customer'] : existingValues[1],
            orderData.Deskripsi !== undefined ? orderData.Deskripsi : existingValues[2],
            orderData.Ukuran !== undefined ? orderData.Ukuran : existingValues[3],
            orderData.Qty !== undefined ? orderData.Qty : existingValues[4],
            orderData.Harga !== undefined ? orderData.Harga : existingValues[5],
            orderData['NO INV'] !== undefined ? orderData['NO INV'] : existingValues[6],
            existingValues[7], existingValues[8], existingValues[9], existingValues[10], existingValues[11], existingValues[12],
            orderData.Bulan !== undefined ? orderData.Bulan : existingValues[13],
            orderData.Tahun !== undefined ? orderData.Tahun : existingValues[14],
            existingValues[15], // PO Status 
            existingValues[16] // NoSJWarna
        ];
        range.setValues([updatedRow]);
        SpreadsheetApp.flush();

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const updatedValues = range.getValues()[0];
        return { status: 'success', data: _rowToObject(updatedValues, headers, rowNumber) };
    } catch (e) {
        return { status: 'error', message: e.message };
    }
}

function deleteWorkOrder(rowNumber) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('WorkOrders');
    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Data berhasil dihapus.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function updateOrderStatus(rowNumber, columnName, value) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('WorkOrders');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = headers.indexOf(columnName) + 1;
    if (columnIndex > 0) {
      sheet.getRange(rowNumber, columnIndex).setValue(value);
      SpreadsheetApp.flush();
      return { status: 'success', message: 'Status berhasil diperbarui.' };
    } else {
      return { status: 'error', message: 'Kolom tidak ditemukan.' };
    }
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// --- FUNGSI QUOTATION ---
function getQuotations() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Quotations');
    if (sheet.getLastRow() < 2) return [];

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    
    return values.map((row, index) => {
      let obj = _rowToObject(row, headers, index + 2);
      if (obj.Items && typeof obj.Items === 'string') {
        try { obj.Items = JSON.parse(obj.Items); } catch(e) { obj.Items = []; }
      }
      return obj;
    });
  } catch(e) {
    throw new Error('Error pada getQuotations: ' + e.message);
  }
}

function addQuotation(quotationData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Quotations');
    const date = new Date();
    const formattedDate = Utilities.formatDate(date, TIMEZONE, "dd/MM/yyyy");
    const lastRow = sheet.getLastRow();
    const newQuotationNumber = "QUO-" + date.getFullYear() + (date.getMonth() + 1).toString().padStart(2, '0') + "-" + (lastRow > 0 ? lastRow : 1);
    const newRow = [
      newQuotationNumber, formattedDate, quotationData.customerName,
      quotationData.project, JSON.stringify(quotationData.items),
      quotationData.total, "Pending"
    ];
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Quotation berhasil dibuat.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function deleteQuotation(rowNumber) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Quotations');
    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Data berhasil dihapus.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function updateQuotation(rowNumber, quotationData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Quotations');
    const formattedDate = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy");
    const rowValues = [
      quotationData.quotationNumber, formattedDate, quotationData.customerName,
      quotationData.project, JSON.stringify(quotationData.items),
      quotationData.total, quotationData.status
    ];
    sheet.getRange(rowNumber, 1, 1, rowValues.length).setValues([rowValues]);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Data berhasil diperbarui.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getQuotationByRow(rowNumber) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Quotations');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    let obj = _rowToObject(rowValues, headers, rowNumber);
    if (obj.Items && typeof obj.Items === 'string') {
      try { obj.Items = JSON.parse(obj.Items);
    } catch (e) { obj.Items = [] }
    }
    return obj;
  } catch(e) {
    throw new Error('Error pada getQuotationByRow: ' + e.message);
  }
}

// --- FUNGSI INVOICE ---
function getInvoiceData(invoiceNumber) {
    try {
        // MODIFIKASI: Gunakan helper _getSheetData (Optimasi Performa)
        const { headers, values } = _getSheetData('WorkOrders');
        
        if (!invoiceNumber || invoiceNumber.trim() === '' || values.length === 0) return [];

        const invIndex = headers.indexOf('NO INV');

        if (invIndex === -1) throw new Error('Kolom "NO INV" tidak ditemukan.');
        const searchInvNumber = parseInt(invoiceNumber.trim(), 10);
        if (isNaN(searchInvNumber)) return [];
        
        const data = [];
        values.forEach((row, index) => {
            const rowInvNumber = parseInt(row[invIndex], 10);
            if (rowInvNumber === searchInvNumber) {
                data.push(_rowToObject(row, headers, index + 2)); // index + 2 karena header
            }
        });
        return data;
    } catch(e) {
        throw new Error(`Error di getInvoiceData: ${e.message}`);
    }
}

// --- FUNGSI LAPORAN KEUANGAN ---
function getFinancialData(month, year) {
  try {
    // MODIFIKASI: Gunakan helper _getSheetData (Optimasi Performa)
    const { headers, values } = _getSheetData('LaporanKeuangan');
    if (values.length === 0) return [];
    
    return values.map((row, index) => _rowToObject(row, headers, index + 2))
           .filter(record => {
               // Perlu diubah karena _rowToObject mengubah Date menjadi string "dd/MM/yyyy"
               const dateParts = record.Tanggal.split('/'); // Tanggal dalam format DD/MM/YYYY
               return dateParts[1] == month && dateParts[2] == year;
             });
  } catch (e) {
    throw new Error('Gagal mengambil data dari Sheet: ' + e.message);
  }
}

function addFinancialRecord(record) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("LaporanKeuangan");
    const newRow = [
      new Date(record.tanggal), record.tipe, record.deskripsi,
      parseFloat(record.jumlah), record.sumber
    ];
    sheet.appendRow(newRow);
    return { status: "success", message: "Data berhasil ditambahkan." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

// --- FUNGSI STOK BAHAN ---
function getStokBahan() {
  try {
    // MODIFIKASI: Gunakan helper _getSheetData
    const { headers, values } = _getSheetData('StokBahan');
    if (values.length === 0) return [];
    
    return values.map((row, index) => _rowToObject(row, headers, index + 2));
  } catch (e) {
    throw new Error('Gagal mengambil data Stok Bahan: ' + e.message);
  }
}

function addNewBahan(bahanData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("StokBahan");
    const kodeColumn = sheet.getRange("A2:A" + sheet.getLastRow()).getValues().flat(); // Memperbaiki rentang untuk mencari kode
    if (kodeColumn.includes(bahanData.kode)) {
      throw new Error(`Kode Bahan "${bahanData.kode}" sudah ada.`);
    }

    const newRow = [
      bahanData.kode, bahanData.nama, bahanData.satuan,
      bahanData.kategori, parseFloat(bahanData.stok) || 0,
      bahanData.lokasi, new Date()
    ];
    sheet.appendRow(newRow);
    return { status: "success", message: "Bahan baru berhasil ditambahkan." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function updateStok(updateData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const stokSheet = spreadsheet.getSheetByName("StokBahan");
    const riwayatSheet = spreadsheet.getSheetByName("RiwayatStok");

    const kodeColumn = stokSheet.getRange("A2:A" + stokSheet.getLastRow()).getValues().flat();
    const rowIndex = kodeColumn.indexOf(updateData.kode);
    if (rowIndex === -1) throw new Error(`Kode Bahan "${updateData.kode}" tidak ditemukan.`);

    const targetRow = rowIndex + 2;
    const stokCell = stokSheet.getRange(targetRow, 5); // Kolom E = Stok
    const stokSebelum = parseFloat(stokCell.getValue()) || 0;
    const jumlahUpdate = parseFloat(updateData.jumlah);
    
    let stokSesudah;
    if (updateData.tipe === 'MASUK') {
      stokSesudah = stokSebelum + jumlahUpdate;
    } else if (updateData.tipe === 'KELUAR') {
      stokSesudah = stokSebelum - jumlahUpdate;
      if (stokSesudah < 0) throw new Error(`Stok tidak mencukupi untuk ${updateData.nama}. Stok saat ini: ${stokSebelum}`);
    } else {
      throw new Error("Tipe update tidak valid.");
    }

    stokCell.setValue(stokSesudah);
    stokSheet.getRange(targetRow, 7).setValue(new Date()); // Kolom G = LastUpdate

    const logRow = [
      new Date(), updateData.kode, updateData.nama,
      updateData.tipe, jumlahUpdate, stokSebelum,
      stokSesudah, updateData.keterangan || ''
    ];
    riwayatSheet.appendRow(logRow);

    return { status: "success", message: `Stok ${updateData.nama} berhasil diupdate.` };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

// --- FUNGSI SURAT JALAN ---

function logSuratJalan(sjData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("SuratJalanLog");

    const date = new Date();
    const yearMonth = Utilities.formatDate(date, TIMEZONE, "yyyyMM");
    // Membuat Nomor Surat Jalan otomatis, contoh: SJ-202510-001
    const lastRow = sheet.getLastRow();
    const newNumber = (lastRow).toString().padStart(3, '0');
    const newSjNumber = `SJ-${yearMonth}-${newNumber}`;

    const newRow = [
      newSjNumber,
      date,
      sjData.noInvoice,
      sjData.namaCustomer,
      JSON.stringify(sjData.items),
      sjData.namaPenerima,
      sjData.alamatKirim,
      sjData.catatan
    ];
    sheet.appendRow(newRow);

    // Kembalikan nomor SJ yang baru dibuat ke frontend
    return { status: "success", noSuratJalan: newSjNumber };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

/**
 * Mengambil daftar barang dari WorkOrders yang status PO-nya 'PRINTED' 
 * dan belum dikirim untuk pewarnaan ('Di Warna' = FALSE).
 */
function getItemsForColoring() {
  try {
    // MODIFIKASI: Gunakan helper _getSheetData (Optimasi Performa)
    const { headers, values } = _getSheetData("WorkOrders");
    
    if (values.length === 0) return [];

    const poStatusIndex = headers.indexOf('PO Status');
    const diWarnaIndex = headers.indexOf('Di Warna');

    if (poStatusIndex === -1 || diWarnaIndex === -1) {
      throw new Error("Kolom 'PO Status' atau 'Di Warna' tidak ditemukan di sheet WorkOrders.");
    }

    const results = [];
    values.forEach((row, index) => {
      // Perbaikan kriteria: Cek apakah status adalah 'PRINTED' ATAU 'READY'
      const statusValue = String(row[poStatusIndex]).toUpperCase();
      const isPrinted = statusValue === 'PRINTED' || statusValue === 'READY';
      
      const isColored = row[diWarnaIndex] === true || String(row[diWarnaIndex]).toUpperCase() === 'TRUE';

      if (isPrinted && !isColored) {
        // index + 2 karena header
        results.push(_rowToObject(row, headers, index + 2));
      }
    });
    return results;
  } catch (e) {
    throw new Error('Gagal mengambil item untuk pewarnaan: ' + e.message);
  }
}

/**
 * Membuat Surat Jalan untuk Vendor Pewarnaan dan mengupdate status barang.
 */
function createSuratJalanWarna(sjData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = spreadsheet.getSheetByName("SuratJalanLog");
    const woSheet = spreadsheet.getSheetByName("WorkOrders");

    const date = new Date();
    const yearMonth = Utilities.formatDate(date, TIMEZONE, "yyyyMM");
    const lastRow = logSheet.getLastRow();
    const newNumber = (lastRow).toString().padStart(3, '0');
    const newSjNumber = `SJW-${yearMonth}-${newNumber}`; // SJW = Surat Jalan Warna

    // 1. Log Surat Jalan ke 'SuratJalanLog'
    const logRow = [
      "VENDOR", // Tipe
      newSjNumber,
      date,
      null, // NoInvoice (kosong untuk SJ Vendor)
      sjData.vendor, // NamaCustomer diisi dengan nama Vendor
      JSON.stringify(sjData.items),
      sjData.vendor, // NamaPenerima diisi dengan nama Vendor
      sjData.alamatVendor || '',
      sjData.catatan || ''
    ];
    logSheet.appendRow(logRow);
    
    // 2. Update status di 'WorkOrders'
    const headers = woSheet.getRange(1, 1, 1, woSheet.getLastColumn()).getValues()[0];
    const diWarnaIndex = headers.indexOf('Di Warna') + 1;
    const noSjWarnaIndex = headers.indexOf('NoSJWarna') + 1;

    if (diWarnaIndex === 0 || noSjWarnaIndex === 0) {
      throw new Error("Kolom 'Di Warna' atau 'NoSJWarna' tidak ditemukan di WorkOrders.");
    }

    sjData.itemRowNumbers.forEach(rowNum => {
      woSheet.getRange(rowNum, diWarnaIndex).setValue(true); // Centang 'Di Warna'
      woSheet.getRange(rowNum, noSjWarnaIndex).setValue(newSjNumber); // Isi nomor SJ Warna
    });
    SpreadsheetApp.flush();
    return { status: "success", noSuratJalan: newSjNumber, items: sjData.items };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

// =======================================================
// FUNGSI-FUNGSI UNTUK KARYAWAN & PAYROLL
// =======================================================

/**
 * Mengambil daftar karyawan yang berstatus "Aktif" saja.
 * Hanya mengambil ID dan Nama untuk efisiensi.
 */
function getActiveKaryawanList() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("DataKaryawan");
    if (sheet.getLastRow() < 2) return [];

    const data = sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
    const activeKaryawan = [];
    data.forEach((row, index) => {
      // Kolom C (index 2) adalah Status
      if (row[2] === 'Aktif') {
        activeKaryawan.push({
          rowNumber: index + 2,
          id: row[0], // Kolom A = IDKaryawan
          nama: row[1]  // Kolom B = Nama
        });
      }
    });
    return activeKaryawan;
  } catch (e) {
    throw new Error('Gagal mengambil daftar karyawan aktif: ' + e.message);
  }
}

/**
 * Memproses penggajian untuk satu karyawan, mengupdate sisa kasbon, dan mencatatnya di log.
 */
function processPayroll(payrollData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dataSheet = spreadsheet.getSheetByName("DataKaryawan");
    const logSheet = spreadsheet.getSheetByName("PayrollLog");

    // --- 1. Ambil data master karyawan ---
    const karyawanRow = dataSheet.getRange(payrollData.rowNumber, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    const karyawanObj = _rowToObject(karyawanRow, headers, payrollData.rowNumber);

    const gajiHarian = parseFloat(karyawanObj.GajiHarian);
    const totalKasbon = parseFloat(karyawanObj.TotalKasbon);
    const potonganBPJS = parseFloat(karyawanObj.PotonganBPJS);
    const potonganKasbon = parseFloat(payrollData.potonganKasbon);
    if (potonganKasbon > totalKasbon) {
      throw new Error(`Potongan kasbon (${potonganKasbon}) lebih besar dari total sisa kasbon (${totalKasbon}).`);
    }

    // --- 2. Lakukan semua perhitungan ---
    const gajiPokok = gajiHarian * parseFloat(payrollData.hariKerja);
    const uangLembur = gajiHarian * parseFloat(payrollData.lembur);
    const totalPendapatan = gajiPokok + uangLembur;
    const totalPotongan = potonganBPJS + potonganKasbon;
    const gajiBersih = totalPendapatan - totalPotongan;
    const sisaKasbonBaru = totalKasbon - potonganKasbon;
    
    // --- 3. Update Total Kasbon di sheet DataKaryawan ---
    const kasbonColumnIndex = headers.indexOf("TotalKasbon") + 1;
    if (kasbonColumnIndex > 0) {
      dataSheet.getRange(payrollData.rowNumber, kasbonColumnIndex).setValue(sisaKasbonBaru);
    } else {
      throw new Error("Kolom TotalKasbon tidak ditemukan di sheet DataKaryawan.");
    }

    // --- 4. Catat transaksi di PayrollLog ---
    const logRow = [
      `PAY-${Date.now()}`, // IDPayroll Unik
      new Date(), // TanggalCetak
      karyawanObj.IDKaryawan,
      karyawanObj.Nama,
      `${payrollData.periodeBulan}/${payrollData.periodeTahun}`, // Periode
      payrollData.hariKerja,
      payrollData.lembur,
      gajiPokok,
      uangLembur,
      potonganKasbon,
      gajiBersih
    ];
    logSheet.appendRow(logRow);

    SpreadsheetApp.flush(); // Pastikan semua perubahan tersimpan
    
    return { status: "success", message: `Payroll untuk ${karyawanObj.Nama} berhasil diproses.` };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

/**
 * Mengambil semua data dari sheet DataKaryawan.
 */
function getAllKaryawanData() {
  try {
    // MODIFIKASI: Gunakan helper _getSheetData (Optimasi Performa)
    const { headers, values } = _getSheetData("DataKaryawan");
    if (values.length === 0) return [];
    
    // Ubah setiap baris data menjadi objek dan kembalikan sebagai array
    return values.map((row, index) => _rowToObject(row, headers, index + 2));
  } catch (e) {
    // Jika terjadi error, kirim pesan error ke frontend
    throw new Error('Gagal mengambil data Karyawan: ' + e.message);
  }
}

/**
 * Menambahkan karyawan baru ke sheet DataKaryawan.
 */
function addNewKaryawan(karyawanData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("DataKaryawan");
    const lastRow = sheet.getLastRow();
    
    // Membuat ID Karyawan baru secara otomatis, contoh: K-001, K-002, dst.
    const newId = `K-${(lastRow).toString().padStart(3, '0')}`;

    const newRow = [
      newId,
      karyawanData.nama,
      "Aktif", // Status default saat karyawan baru dibuat
      parseFloat(karyawanData.gajiHarian) || 0,
      parseFloat(karyawanData.totalKasbon) || 0,
      parseFloat(karyawanData.potonganBPJS) || 0
    ];
    
    sheet.appendRow(newRow);
    
    return { status: "success", message: "Karyawan baru berhasil ditambahkan." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function toggleKaryawanStatus(rowNum, currentStatus) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("DataKaryawan");
    
    // Kolom Status (Kolom C = 3)
    const statusColumn = 3; 
    const newStatus = currentStatus === 'Aktif' ? 'Non-Aktif' : 'Aktif';
    
    sheet.getRange(rowNum, statusColumn).setValue(newStatus);
    SpreadsheetApp.flush();
    
    return { status: "success", message: `Status karyawan berhasil diubah menjadi ${newStatus}.` };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function updateKaryawanData(karyawanData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("DataKaryawan");
    
    // Ambil data lama untuk dipertahankan (misal: ID dan Status)
    const existingRow = sheet.getRange(karyawanData.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Kolom 1=ID, 2=Nama, 3=Status, 4=GajiHarian, 5=TotalKasbon, 6=PotonganBPJS
    const updatedRow = [
      existingRow[0], // ID
      karyawanData.nama,
      existingRow[2], // Status
      parseFloat(karyawanData.gajiHarian) || existingRow[3],
      parseFloat(karyawanData.totalKasbon) || existingRow[4],
      parseFloat(karyawanData.potonganBPJS) || existingRow[5]
    ];
    
    sheet.getRange(karyawanData.rowNumber, 1, 1, updatedRow.length).setValues([updatedRow]);
    SpreadsheetApp.flush();

    return { status: "success", message: "Data karyawan berhasil diperbarui." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

// =======================================================
// PENAMBAHAN FUNGSI DASHBOARD (Optimized)
// =======================================================

/**
 * Mengambil dan mengagregasi data produksi bulanan dari WorkOrders.
 */
function getDashboardData(month, year) {
  try {
    // MODIFIKASI: Gunakan helper _getSheetData (Optimasi Performa)
    const { headers, values } = _getSheetData('WorkOrders');
    
    if (values.length === 0) {
      return { totalNominal: 0, statusCounts: { 'Di Produksi': 0, 'Di Warna': 0, 'Siap Kirim': 0, 'Di Kirim': 0, 'Lunas': 0, 'TotalWO': 0 } };
    }

    // Temukan indeks kolom yang relevan
    const indices = {
      Ukuran: headers.indexOf('Ukuran'),
      Qty: headers.indexOf('Qty'),
      Harga: headers.indexOf('Harga'),
      DiProduksi: headers.indexOf('Di Produksi'),
      DiWarna: headers.indexOf('Di Warna'),
      SiapKirim: headers.indexOf('Siap Kirim'),
      DiKirim: headers.indexOf('Di Kirim'),
      Pembayaran: headers.indexOf('Pembayaran'),
      Bulan: headers.indexOf('Bulan'),
      Tahun: headers.indexOf('Tahun')
    };

    if (Object.values(indices).some(i => i === -1)) {
      throw new Error('Kolom WorkOrders tidak lengkap untuk dashboard.');
    }

    const filterMonth = parseInt(month, 10);
    const filterYear = parseInt(year, 10);
    
    let totalNominal = 0;
    const statusCounts = {
      'Di Produksi': 0, 'Di Warna': 0, 'Siap Kirim': 0, 'Di Kirim': 0, 'Lunas': 0, 'TotalWO': 0
    };

    values.forEach(row => {
      const rowMonth = parseInt(row[indices.Bulan], 10);
      const rowYear = parseInt(row[indices.Tahun], 10);

      // Filter berdasarkan bulan dan tahun yang dipilih
      if (rowMonth === filterMonth && rowYear === filterYear) {
        statusCounts.TotalWO++;
        
        const ukuran = parseFloat(row[indices.Ukuran]) || 0;
        const qty = parseFloat(row[indices.Qty]) || 0;
        const harga = parseFloat(row[indices.Harga]) || 0;
        const total = ukuran * qty * harga;
        totalNominal += total;
        
        // Hitung status (memeriksa nilai boolean/string 'TRUE')
        if (row[indices.DiProduksi] === true || String(row[indices.DiProduksi]).toUpperCase() === 'TRUE') {
          statusCounts['Di Produksi']++;
        }
        if (row[indices.DiWarna] === true || String(row[indices.DiWarna]).toUpperCase() === 'TRUE') {
          statusCounts['Di Warna']++;
        }
        if (row[indices.SiapKirim] === true || String(row[indices.SiapKirim]).toUpperCase() === 'TRUE') {
          statusCounts['Siap Kirim']++;
        }
        if (row[indices.DiKirim] === true || String(row[indices.DiKirim]).toUpperCase() === 'TRUE') {
          statusCounts['Di Kirim']++;
        }
        // 'Lunas' diwakili oleh kolom Pembayaran
        if (row[indices.Pembayaran] === true || String(row[indices.Pembayaran]).toUpperCase() === 'TRUE') {
          statusCounts['Lunas']++;
        }
      }
    });

    return { totalNominal, statusCounts };
  } catch (e) {
    throw new Error(`Error di getDashboardData: ${e.message}`);
  }
}


/**
 * Mengambil daftar order yang status 'Siap Kirim'-nya TRUE dan TIDAK 'Di Kirim'
 * untuk periode tertentu, agar tidak menampilkan barang yang sudah terkirim.
 */
function getReadyToShipOrders(month, year) {
    try {
        const { headers, values } = _getSheetData('WorkOrders');
        
        if (values.length === 0) return [];

        const monthIndex = headers.indexOf('Bulan');
        const yearIndex = headers.indexOf('Tahun');
        const siapKirimIndex = headers.indexOf('Siap Kirim');
        const diKirimIndex = headers.indexOf('Di Kirim'); // NEW: Index Kolom Di Kirim
        
        if (monthIndex === -1 || yearIndex === -1 || siapKirimIndex === -1 || diKirimIndex === -1) {
            throw new Error('Kolom status/bulan/tahun tidak ditemukan di WorkOrders.');
        }

        const filterMonth = parseInt(month, 10);
        const filterYear = parseInt(year, 10);
        const readyOrders = [];
        
        values.forEach((row, index) => {
            // FIX: Menggunakan index variabel yang sudah didefinisikan (monthIndex, yearIndex)
            const rowMonth = parseInt(row[monthIndex], 10); 
            const rowYear = parseInt(row[yearIndex], 10);
            
            // Cek Status Siap Kirim
            const isReady = row[siapKirimIndex] === true || String(row[siapKirimIndex]).toUpperCase() === 'TRUE';
            
            // Cek Status Di Kirim (Harus FALSE/kosong)
            const isShipped = row[diKirimIndex] === true || String(row[diKirimIndex]).toUpperCase() === 'TRUE';


            // LOGIKA FILTER BARU: Harus Siap Kirim DAN BELUM Di Kirim
            if (rowMonth === filterMonth && rowYear === filterYear && isReady && !isShipped) {
                // index + 2 karena header (baris 1) sudah di-shift
                readyOrders.push(_rowToObject(row, headers, index + 2)); 
            }
        });
        return readyOrders;
    } catch (e) {
        throw new Error(`Error di getReadyToShipOrders: ${e.message}`);
    }
}

// =======================================================
// PENAMBAHAN FUNGSI UNTUK CETAK SLIP GAJI
// =======================================================

/**
 * Mengambil konten HTML dari template slip gaji.
 * Fungsi ini dipanggil dari client-side (JavaScript) untuk mendapatkan template-nya.
 */
function getSlipGajiHtml() {
  return HtmlService.createHtmlOutputFromFile('slipgaji').getContent();
}
