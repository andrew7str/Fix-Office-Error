# 🛠️ Microsoft Office Error 0xc0000142 Fixer

[![Platform](https://img.shields.io/badge/Platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Batch](https://img.shields.io/badge/Script-Batch-orange.svg)](https://en.wikipedia.org/wiki/Batch_file)

Solusi praktis dan otomatis untuk mengatasi **Application Error 0xc0000142** pada Microsoft Office 2013, 2016, dan 2019 di sistem operasi Windows.

---

## 📖 Deskripsi Masalah

Sering kali saat membuka Microsoft Office (Word, Excel, PowerPoint), muncul pesan **"Updating Office, please wait a moment"** yang diikuti dengan kode kesalahan **0xc0000142**.

### Penyebab Umum:
* **ClickToRunSvc:** Layanan latar belakang yang menjalankan Office tidak merespons atau berhenti.
* **STATUS_DLL_INIT_FAILED:** Gagalnya inisialisasi file `.dll` yang diperlukan oleh aplikasi Office.

Script ini bekerja dengan merestart layanan **Microsoft Office Click-to-Run** secara otomatis menggunakan hak akses administratif.

## 🚀 Fitur Utama
- **Auto-Elevation Check:** Memastikan script berjalan dengan hak akses Administrator.
- **Service Reset:** Menghentikan dan memulai ulang `ClickToRunSvc` secara bersih.
- **User Friendly:** Memberikan notifikasi status proses secara real-time.

## 🛠️ Cara Penggunaan

### Metode 1: Download & Run (Direkomendasikan)
1. Pergi ke tab [Releases](https://github.com/andrew7str/Fix-Office-Error/releases) atau download file `FixOffice.bat` dari repositori ini.
2. Klik kanan pada file `FixOffice.bat`.
3. Pilih **Run as Administrator**.
4. Tunggu hingga muncul pesan "BERHASIL", lalu coba buka kembali Office Anda.
