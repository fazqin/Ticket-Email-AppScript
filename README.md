# Sistem E-Tiket & Scanner QR Event

Sistem registrasi ulang, verifikasi voucher konsumsi, dan distribusi e-tiket otomatis berbasis **Google Apps Script** dan **Google Sheets**. Dirancang *mobile-first* agar panitia acara (**SPMS The Halal Portofolio Project**) dapat memindai QR Code peserta langsung dari browser ponsel tanpa perlu instalasi aplikasi tambahan.

---

## Fitur Utama

- **Browser-Based QR Scanner:** Pemindaian QR Code secara *real-time* menggunakan kamera belakang/depan ponsel melalui pustaka `html5-qrcode`, dilengkapi *cooldown* 2 detik untuk mencegah pemindaian ganda tidak sengaja.
- **Anti Double Check-In & Klaim Voucher:** Memvalidasi ID tiket secara otomatis terhadap database Google Sheets; menolak tiket yang tidak terdaftar (`NOT_FOUND`) maupun tiket yang sudah digunakan (`ALREADY_CHECKED_IN`) beserta detail waktu *check-in* (`Asia/Jakarta`).
- **Pelacakan Panitia / PJ (Person in Charge):** Daftar nama PJ ditarik secara dinamis dari Google Sheets dan disimpan pada `localStorage` browser, lengkap dengan riwayat peserta yang telah di-scan oleh masing-masing PJ.
- **Umpan Balik Audio & Visual (Web Audio API):** Menghasilkan nada *beep* berbeda untuk status berhasil, peringatan duplikat, dan tiket tidak valid disertai *banner* indikator warna.
- **Automated HTML E-Ticket Blast:** Mengirimkan e-tiket bergaya *stub-ticket* melalui `GmailApp` lengkap dengan QR Code unik tiap peserta, format tanggal pendaftaran berbahasa Indonesia, aset gambar *inline* dari Google Drive (`cid:`), perlindungan warna *dark mode*, serta jeda *rate-limit* pengiriman massal.

---

## Tech Stack

- **Frontend:** HTML5, CSS3 (Responsive Dark UI), Vanilla JavaScript, [html5-qrcode](https://github.com/mebjas/html5-qrcode) v2.3.8, Web Audio API (`AudioContext`), LocalStorage API
- **Backend & Serverless Hosting:** Google Apps Script (`HtmlService`, `SpreadsheetApp`, `GmailApp`, `DriveApp`)
- **Database:** Google Sheets (`Absensi & Voucher`)
- **QR Generator:** `api.qrserver.com`

---

## Struktur Proyek

```text
├── Index.html                # Antarmuka Web App mobile-first untuk scanner QR & riwayat PJ
├── kode.gs                   # Backend Google Apps Script (API check-in, log PJ, & email blast)
└── Absensi & Voucher.xlsx    # Template / struktur database peserta & status tiket
