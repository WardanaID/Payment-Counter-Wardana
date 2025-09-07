// Ganti dengan ID Spreadsheet Anda
const SPREADSHEET_ID = 'ID_SPREADSHEET_ANDA';

// Daftar semua subjek email yang relevan
const TRANSACTION_SUBJECTS = [
  "Transaksi Pembayaran QRIS MPM Kamu Berhasil",
  "Pembayaran Shopee Berhasil",
  "Tarik Tunai Via ATM BSI",
  "Pembelian Saldo E-Wallet Berhasil",
  "Pembayaran Tokopedia Berhasil"
];

function hitungTransaksiQRIS() {
  Logger.log("Memulai eksekusi skrip.");
  
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Spreadsheet berhasil dibuka.");
  } catch (e) {
    Logger.log("ERROR: Spreadsheet tidak ditemukan. Mohon periksa kembali SPREADSHEID_ID Anda.");
    return;
  }
  
  // Sheet untuk ringkasan
  let summarySheet = spreadsheet.getSheetByName("Ringkasan Transaksi");
  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet("Ringkasan Transaksi", 0);
  }
  summarySheet.clear();
  
  // Sheet untuk data detail transaksi
  let detailSheet = spreadsheet.getSheetByName("Data Transaksi");
  if (!detailSheet) {
    detailSheet = spreadsheet.insertSheet("Data Transaksi", 1);
  }
  detailSheet.clear();
  
  detailSheet.getRange('A1').setValue('Nomor');
  detailSheet.getRange('B1').setValue('Tanggal');
  detailSheet.getRange('C1').setValue('Nominal');
  detailSheet.getRange('D1').setValue('Tautan Email');
  detailSheet.getRange('E1').setValue('Jenis Transaksi'); // Kolom baru untuk jenis transaksi
  detailSheet.getRange('F1').setValue('Total Nominal per Jenis'); // Kolom baru untuk total per jenis
  detailSheet.getRange('G1').setValue('Total Keseluruhan'); // Kolom baru untuk total keseluruhan
  
  const nominalData = [];
  const linkData = [];
  const transactionSummary = {};
  
  TRANSACTION_SUBJECTS.forEach(subject => {
    const threads = GmailApp.search('subject:"' + subject + '"');
    Logger.log("Ditemukan " + threads.length + " thread email dengan subjek '" + subject + "'.");
    
    threads.forEach(thread => {
      thread.getMessages().forEach(message => {
        const body = message.getPlainBody();
        const date = message.getDate();
        const messageUrl = "https://mail.google.com/mail/u/0/#inbox/" + thread.getId();
        
        let nominal = 0;
        let nominalStr = "";

        // Menyesuaikan regex untuk setiap jenis transaksi
        if (subject.includes("QRIS")) {
          const regex = /Nominal Transaksi Rp\s*([\d\.,]+)/;
          const match = body.match(regex);
          if (match && match[1]) {
            nominalStr = match[1].replace(/\./g, '').replace(/,/g, '.');
          }
        } else if (subject.includes("Shopee") || subject.includes("Tokopedia")) {
          const regex = /Rp\s*([\d\.,]+)/;
          const match = body.match(regex);
          if (match && match[1]) {
            nominalStr = match[1].replace(/\./g, '').replace(/,/g, '.');
          }
        } else if (subject.includes("Tarik Tunai")) {
          const regex = /Nominal Tarik Tunai\s*Rp\s*([\d\.,]+)/; 
          const match = body.match(regex);
          if (match && match[1]) {
            nominalStr = match[1].replace(/\./g, '').replace(/,/g, '.');
          }
        } else if (subject.includes("E-Wallet")) {
           const regex = /Nominal Top Up\s*Rp\s*([\d\.,]+)/;
          const match = body.match(regex);
          if (match && match[1]) {
            nominalStr = match[1].replace(/\./g, '').replace(/,/g, '.');
          }
        }
        
        if (nominalStr) {
          nominal = parseFloat(nominalStr);
        }
        
        if (!isNaN(nominal)) {
          Logger.log("Ditemukan nominal: Rp" + nominal + " dari email tanggal " + Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss") + " dengan subjek: " + subject);
          
          // Memformat tanggal ke string untuk mencegah Google Sheets salah menginterpretasi
          const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
          nominalData.push([nominalData.length + 1, formattedDate, nominal, messageUrl, subject]);
          
          if (!transactionSummary[subject]) {
            transactionSummary[subject] = 0;
          }
          transactionSummary[subject] += nominal;
        }
      });
    });
  });
  
  // Menulis ringkasan total ke sheet "Ringkasan Transaksi"
  summarySheet.getRange('A1').setValue('Jenis Transaksi');
  summarySheet.getRange('B1').setValue('Total Nominal');
  
  let totalNominalKeseluruhan = 0;
  let summaryRow = 2;
  for (const [type, total] of Object.entries(transactionSummary)) {
    summarySheet.getRange(summaryRow, 1).setValue(type);
    summarySheet.getRange(summaryRow, 2).setValue(total).setNumberFormat("Rp#,##0");
    totalNominalKeseluruhan += total;
    summaryRow++;
  }
  
  // Menambahkan total keseluruhan
  summarySheet.getRange(summaryRow + 1, 1).setValue('TOTAL KESELURUHAN');
  summarySheet.getRange(summaryRow + 1, 2).setValue(totalNominalKeseluruhan).setNumberFormat("Rp#,##0");

  // Autofit kolom di sheet ringkasan
  summarySheet.autoResizeColumns(1, 2);

  // Menulis data detail ke sheet "Data Transaksi"
  if (nominalData.length > 0) {
    // Format nominal sebagai teks dengan simbol Rp
    const formattedNominalData = nominalData.map(row => {
      // Gunakan toLocaleString() untuk format mata uang yang sederhana
      const formattedNominal = "Rp" + row[2].toLocaleString('id-ID');
      return [row[0], row[1], formattedNominal, row[3], row[4]];
    });

    const detailRange = detailSheet.getRange(2, 1, formattedNominalData.length, formattedNominalData[0].length);
    detailRange.setValues(formattedNominalData);
    
    // Menambahkan total per jenis transaksi di kolom F
    let currentRow = 2;
    for (const [type, total] of Object.entries(transactionSummary)) {
      const typeRows = nominalData.filter(row => row[4] === type);
      if (typeRows.length > 0) {
        const startRow = currentRow;
        const endRow = currentRow + typeRows.length - 1;
        detailSheet.getRange(startRow, 6).setValue("Rp" + total.toLocaleString('id-ID'));
        detailSheet.getRange(startRow, 6, typeRows.length).mergeVertically();
        currentRow = endRow + 1;
      }
    }
    
    // Menambahkan total keseluruhan di kolom G
    detailSheet.getRange(2, 7).setValue("Rp" + totalNominalKeseluruhan.toLocaleString('id-ID'));
    detailSheet.getRange(2, 7, nominalData.length).mergeVertically();
    
    // Mengatur perataan teks untuk kolom Tanggal agar rata kanan
    detailSheet.getRange("B2:B" + (nominalData.length + 1)).setHorizontalAlignment("right");
    
    // Autofit kolom di sheet detail
    detailSheet.autoResizeColumns(1, detailSheet.getLastColumn());
    
    // Mengubah data detail menjadi tabel
    const headerRange = detailSheet.getRange(1, 1, 1, detailSheet.getLastColumn());
    headerRange.setBackground("#f3f3f3").setFontWeight("bold");
    const dataRange = detailSheet.getRange(1, 1, detailSheet.getLastRow(), detailSheet.getLastColumn());
    dataRange.setBorder(true, true, true, true, true, true);
    
    Logger.log("Data dan total nominal berhasil ditulis ke spreadsheet.");
  } else {
    detailSheet.getRange(2, 1).setValue('Tidak ada email yang ditemukan.');
    Logger.log("Tidak ada email yang ditemukan.");
  }
  
  // Tambahkan sheet terpisah untuk setiap jenis transaksi
  for (const [type, total] of Object.entries(transactionSummary)) {
    const sheetName = type.replace(/[^a-zA-Z0-9 ]/g, ''); // Membersihkan nama sheet
    let transactionSheet = spreadsheet.getSheetByName(sheetName);
    if (!transactionSheet) {
      transactionSheet = spreadsheet.insertSheet(sheetName);
    }
    transactionSheet.clear();
    transactionSheet.getRange('A1:E1').setValues([['Nomor', 'Tanggal', 'Nominal', 'Tautan Email', 'Jenis Transaksi']]);

    // Filter data untuk jenis transaksi ini
    const typeData = nominalData.filter(row => row[4] === type);

    if (typeData.length > 0) {
      // Memasukkan hanya kolom yang relevan ke sheet baru
      const simplifiedData = typeData.map(row => {
        const formattedNominal = "Rp" + row[2].toLocaleString('id-ID');
        return [row[0], row[1], formattedNominal, row[3], row[4]];
      });
      
      const dataRange = transactionSheet.getRange(2, 1, simplifiedData.length, simplifiedData[0].length);
      dataRange.setValues(simplifiedData);

      // Mengatur perataan teks untuk kolom Tanggal agar rata kanan
      transactionSheet.getRange("B2:B" + (simplifiedData.length + 1)).setHorizontalAlignment("right");

      // Menambahkan total di bawah data
      const lastRow = simplifiedData.length + 2;
      transactionSheet.getRange(lastRow, 2).setValue('Total');
      transactionSheet.getRange(lastRow, 3).setValue("Rp" + total.toLocaleString('id-ID'));
      
      // Mengubah data menjadi tabel dengan format manual
      const transactionHeaderRange = transactionSheet.getRange(1, 1, 1, 5);
      transactionHeaderRange.setBackground("#f3f3f3").setFontWeight("bold");
      const transactionRange = transactionSheet.getRange(1, 1, transactionSheet.getLastRow(), 5);
      transactionRange.setBorder(true, true, true, true, true, true);
      
    } else {
      transactionSheet.getRange(2, 1).setValue('Tidak ada transaksi.');
    }
    
    // Autofit kolom di sheet transaksi
    transactionSheet.autoResizeColumns(1, transactionSheet.getLastColumn());
  }
  
  spreadsheet.setActiveSheet(detailSheet);
  
  Logger.log("Skrip selesai.");
}
