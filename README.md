# SIM Raport Madin (Sistem Informasi Manajemen)

Aplikasi Sistem Informasi Manajemen Raport Madrasah Diniyah (Madin) berbasis **Google Apps Script (GAS)** dan **Google Sheets** sebagai database.

## 🚀 Fitur Utama
*   **Sistem Autentikasi:** Login untuk Administrator dan Guru.
*   **Dashboard Analitik:** Ringkasan statistik data madrasah.
*   **Master Data:** Pengelolaan data Guru, Siswa, Kelas, dan Mata Pelajaran (Mapel).
*   **Transaksi Akademik:** Plotting rombongan belajar (Rombel), Jadwal Mengajar, dan Input Nilai (UTS/UAS).
*   **Pengaturan Profil:** Konfigurasi identitas madrasah dan logo langsung dari aplikasi.
*   **UI/UX Modern:** Tampilan *Glassmorphism* yang responsif untuk Desktop dan Mobile.

## 🛠️ Cara Instalasi (Instruksi Penggunaan)
Karena aplikasi ini dibangun menggunakan Google Apps Script, ikuti langkah berikut untuk menggunakannya:

1. Buat **Google Spreadsheet** baru di Google Drive Anda.
2. Buat sheet/tab sesuai dengan kebutuhan database aplikasi (misal: `mst_users`, `mst_guru`, `mst_siswa`, dll).
3. Klik menu **Ekstensi > Apps Script**.
4. Buat file-file baru di editor Apps Script sesuai dengan nama file di repositori ini.
5. **Penting:** Salin kode dari file `.js` di repositori ini dan tempelkan ke file `.gs` di editor Google Apps Script Anda. Untuk file `.html`, salin dan tempelkan persis sebagai file `.html`.
6. Klik **Terapkan (Deploy) > Deployment Baru**, pilih jenis "Aplikasi Web" (Web App).
7. Atur akses ke "Siapa saja" (Anyone) atau sesuai kebijakan lembaga Anda.
8. Salin URL Web App yang dihasilkan.

## 📝 Lisensi
Hak Cipta Dilindungi &copy; 2026.