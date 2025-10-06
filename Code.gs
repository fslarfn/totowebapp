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
// FUNGSI HELPER INTERNAL
// ===============================================================
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
    // POLA BARU: Buka spreadsheet dan sheet hanya saat fungsi ini dipanggil.
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
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName('WorkOrders');
        
        if (sheet.getLastRow() < 2) return [];

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

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

    const newRow = [
      orderData.Tanggal ? new Date(orderData.Tanggal) : null, orderData['Nama Customer'] || '',
      orderData.Deskripsi || '', orderData.Ukuran || '', orderData.Qty || '', orderData.Harga || '',
      orderData['NO INV'] || '', false, false, false, false, false, '',
      orderData.Bulan || '', orderData.Tahun || '', ''
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
            existingValues[15]
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
      try { obj.Items = JSON.parse(obj.Items); } catch (e) { obj.Items = [] }
    }
    return obj;
  } catch(e) {
    throw new Error('Error pada getQuotationByRow: ' + e.message);
  }
}

// --- FUNGSI INVOICE ---
function getInvoiceData(invoiceNumber) {
    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName('WorkOrders');
        
        if (!invoiceNumber || invoiceNumber.trim() === '') return [];
        if (sheet.getLastRow() < 2) return [];

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
        const invIndex = headers.indexOf('NO INV');

        if (invIndex === -1) throw new Error('Kolom "NO INV" tidak ditemukan.');

        const searchInvNumber = parseInt(invoiceNumber.trim(), 10);
        if (isNaN(searchInvNumber)) return [];
        
        const data = [];
        values.forEach((row, index) => {
            const rowInvNumber = parseInt(row[invIndex], 10);
            if (rowInvNumber === searchInvNumber) {
                data.push(_rowToObject(row, headers, index + 2));
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
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("LaporanKeuangan");
    if (sheet.getLastRow() < 2) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    return data.map(row => _rowToObject(row, headers, 0)) // rowNumber tidak relevan di sini
           .filter(record => {
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
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("StokBahan");
    if (sheet.getLastRow() < 2) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    return data.map((row, index) => _rowToObject(row, headers, index + 2));
  } catch (e) {
    throw new Error('Gagal mengambil data Stok Bahan: ' + e.message);
  }
}

function addNewBahan(bahanData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("StokBahan");
    
    const kodeColumn = sheet.getRange("A2:A").getValues().flat();
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

    const kodeColumn = stokSheet.getRange("A2:A").getValues().flat();
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

// Tambahkan fungsi ini di Code.gs
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

// Tambahkan dua fungsi ini di file Code.gs Anda

/**
 * Mengambil daftar barang dari WorkOrders yang status PO-nya 'PRINTED' 
 * dan belum dikirim untuk pewarnaan ('Di Warna' = FALSE).
 */
function getItemsForColoring() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("WorkOrders");
    if (sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const poStatusIndex = headers.indexOf('PO Status');
    const diWarnaIndex = headers.indexOf('Di Warna');

    if (poStatusIndex === -1 || diWarnaIndex === -1) {
      throw new Error("Kolom 'PO Status' atau 'Di Warna' tidak ditemukan di sheet WorkOrders.");
    }

    const results = [];
    data.forEach((row, index) => {
      const isPrinted = row[poStatusIndex] === 'PRINTED';
      const isColored = row[diWarnaIndex] === true;

      if (isPrinted && !isColored) {
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
    const diWarnaIndex = woSheet.getRange(1, 1, 1, woSheet.getLastColumn()).getValues()[0].indexOf('Di Warna') + 1;
    const noSjWarnaIndex = woSheet.getRange(1, 1, 1, woSheet.getLastColumn()).getValues()[0].indexOf('NoSJWarna') + 1;

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
// FUNGSI-FUNGSI UNTUK PAYROLL
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
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("DataKaryawan");
    if (sheet.getLastRow() < 2) return []; // Kembalikan array kosong jika tidak ada data
    
    // Ambil semua data termasuk header
    const data = sheet.getDataRange().getValues();
    // Pisahkan baris header
    const headers = data.shift();
    
    // Ubah setiap baris data menjadi objek dan kembalikan sebagai array
    return data.map((row, index) => _rowToObject(row, headers, index + 2));
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
