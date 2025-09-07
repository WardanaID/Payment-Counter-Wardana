# Payment-Counter-Wardana

**Payment-Counter-Wardana** adalah skrip Google Apps Script yang digunakan untuk menghitung, merekap, dan memvisualisasikan transaksi keuangan otomatis dari laporan email ke Google Spreadsheet.

## Fitur Utama

- **Ekstraksi Data Otomatis dari Email:**  
  Mengambil data transaksi dari email dengan subjek tertentu seperti pembayaran QRIS, Shopee, Tokopedia, tarik tunai, dan top up e-wallet.
- **Rekapitulasi Transaksi:**  
  Menghitung total nominal per jenis transaksi dan keseluruhan, serta memformat otomatis ke dalam bentuk tabel di Google Spreadsheet.
- **Visualisasi Ringkasan dan Detail:**  
  - Sheet `Ringkasan Transaksi`: Menampilkan total per jenis transaksi dan total keseluruhan.
  - Sheet `Data Transaksi`: Menampilkan detail setiap transaksi, tautan ke email, serta rekap per jenis dan keseluruhan.
  - Sheet khusus per jenis transaksi, misal: `Pembayaran Shopee Berhasil`, berisi data transaksi jenis tersebut.
- **Formatting Otomatis:**  
  Semua data dan nominal diformat secara otomatis dengan simbol Rupiah dan border tabel.

## Cara Kerja

1. **Konfigurasi:**
   - Ganti `SPREADSHEET_ID` di kode dengan ID Google Spreadsheet tujuan Anda.
   - Pastikan daftar `TRANSACTION_SUBJECTS` sesuai dengan notifikasi email bank/payment gateway Anda.

2. **Proses Ekstraksi:**
   - Skrip mencari email dengan subjek sesuai `TRANSACTION_SUBJECTS`.
   - Nominal transaksi diambil dari isi email menggunakan regex yang disesuaikan tiap jenis transaksi.
   - Data yang diambil: Nomor, Tanggal, Nominal, Tautan Email, Jenis Transaksi.

3. **Penulisan ke Spreadsheet:**
   - Data ditulis ke sheet `Data Transaksi` dan direkap ke `Ringkasan Transaksi`.
   - Sheet per jenis transaksi juga dibuat otomatis bila ada data.

4. **Otomatisasi dan Formatting:**
   - Semua nominal diformat Rupiah.
   - Kolom dan baris otomatis diatur agar rapi (autofit & border).
   - Jika tidak ada data, pesan khusus akan muncul di sheet.

## Contoh Penggunaan

1. Deploy skrip Google Apps Script di Google Drive Anda.
2. Jalankan fungsi `hitungTransaksiQRIS`.
3. Lihat hasil rekap dan detail transaksi di Spreadsheet yang sudah dikonfigurasi.

## Catatan

- Kode membutuhkan akses ke Gmail dan Google Spreadsheet.
- Pastikan Anda telah memberikan izin yang diperlukan pada Apps Script.
- Ubah `TRANSACTION_SUBJECTS` jika ada format subjek email yang berbeda.
- Kode ini optimal untuk mengecek semua riwayat transaksi dari Bank BSI melalui Aplikasi BYOND by BSI, kamu bisa mengubah subjek nya berdasarkan pola yang diterima melalui e-mail dari bank yang digunakan

---

**Kontributor:**  
[WardanaID](https://github.com/WardanaID)
