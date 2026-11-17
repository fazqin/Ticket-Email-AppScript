

const SHEET_TICKET_NAME = 'Ticket';
const SHEET_PJ_NAME = 'PJ registrasi';

// 1. Tampilkan Halaman Web App untuk Panitia
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Scan Registrasi CISP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. Ambil Daftar Nama PJ secara dinamis dari Baris 1 tab 'PJ registrasi'
function getPjList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pjSheet = ss.getSheetByName(SHEET_PJ_NAME);
  if (!pjSheet) return ['PJ 1', 'PJ 2', 'PJ 3', 'PJ 4'];
  
  const lastCol = pjSheet.getLastColumn();
  if (lastCol < 1) return ['PJ 1', 'PJ 2', 'PJ 3', 'PJ 4'];
  
  const headers = pjSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const pjList = [];
  
  for (let c = 0; c < headers.length; c++) {
    const val = headers[c] ? headers[c].toString().trim() : '';
    if (val !== '' && !pjList.includes(val)) {
      pjList.push(val);
    }
  }
  return pjList.length > 0 ? pjList : ['PJ 1', 'PJ 2', 'PJ 3', 'PJ 4'];
}

// 3. Ambil riwayat peserta yang sudah di-scan oleh PJ yang bersangkutan
function getPjScannedList(pjName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pjSheet = ss.getSheetByName(SHEET_PJ_NAME);
  const result = [];
  if (!pjSheet) return result;

  const colIndex = findPjColumnIndex(pjSheet, pjName);
  if (colIndex === -1) return result;

  const lastRow = pjSheet.getLastRow();
  if (lastRow < 3) return result;

  const data = pjSheet.getRange(3, colIndex, lastRow - 2, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    const id = data[i][0] ? data[i][0].toString().trim() : '';
    const name = data[i][1] ? data[i][1].toString().trim() : '';
    if (id !== '') {
      result.push({ idTicket: id, name: name });
    }
  }
  return result;
}

// 4. Proses Check-In Peserta dari Scanner QR
function processCheckIn(rawCode, pjName) {
  try {
    const idTicket = rawCode.trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ticketSheet = ss.getSheetByName(SHEET_TICKET_NAME);
    const pjSheet = ss.getSheetByName(SHEET_PJ_NAME);

    if (!ticketSheet) throw new Error("Tab '" + SHEET_TICKET_NAME + "' tidak ditemukan!");

    const data = ticketSheet.getDataRange().getValues();
    let rowIndex = -1;

    // Cari tiket berdasarkan ID Tiket (Kolom A)
    for (let r = 1; r < data.length; r++) {
      if (data[r][0] && data[r][0].toString().trim().toUpperCase() === idTicket.toUpperCase()) {
        rowIndex = r + 1; // 1-based index
        break;
      }
    }

    if (rowIndex === -1) {
      return {
        success: false,
        status: 'NOT_FOUND',
        message: 'Tiket ' + idTicket + ' tidak terdaftar dalam database!'
      };
    }

    const currentAttendance = ticketSheet.getRange(rowIndex, 5).getValue(); // Kolom E: Status Kehadiran
    const currentVoucher = ticketSheet.getRange(rowIndex, 6).getValue();    // Kolom F: Status Voucher
    const currentCheckIn = ticketSheet.getRange(rowIndex, 7).getValue();    // Kolom G: Waktu Check-in
    const attendeeName = ticketSheet.getRange(rowIndex, 2).getValue();      // Kolom B: Nama

    // Cek tiket sudah pernah di-scan
    if (currentAttendance === 'Hadir' || (currentCheckIn && currentCheckIn.toString().trim() !== '')) {
      const formattedTime = Utilities.formatDate(new Date(currentCheckIn), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
      return {
        success: false,
        status: 'ALREADY_CHECKED_IN',
        idTicket: idTicket,
        name: attendeeName,
        checkInTime: formattedTime,
        message: 'Peserta ini SUDAH CHECK-IN sebelumnya pada ' + formattedTime + '!'
      };
    }

    // Update Tab 'Ticket'
    const now = new Date();
    const nowFormatted = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    
    ticketSheet.getRange(rowIndex, 5).setValue('Hadir');
    ticketSheet.getRange(rowIndex, 6).setValue('Sudah Diambil');
    ticketSheet.getRange(rowIndex, 7).setValue(nowFormatted);

    // Rekam ke Tab 'PJ registrasi' di bawah kolom PJ terkait
    if (pjSheet && pjName) {
      appendPjRegistration(pjSheet, pjName, idTicket, attendeeName);
    }

    return {
      success: true,
      status: 'SUCCESS',
      idTicket: idTicket,
      name: attendeeName,
      checkInTime: nowFormatted,
      message: 'Check-in & Voucher konsumsi berhasil dicatat!'
    };
  } catch (err) {
    return {
      success: false,
      status: 'ERROR',
      message: 'Terjadi kesalahan sistem: ' + err.message
    };
  }
}

// Helper: Cari indeks kolom berdasarkan nama PJ
function findPjColumnIndex(sheet, pjName) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    if (headers[c] && headers[c].toString().trim().toUpperCase() === pjName.trim().toUpperCase()) {
      return c + 1; // 1-based index
    }
  }
  return -1;
}

// Helper: Tambahkan peserta ke kolom PJ yang bersangkutan
function appendPjRegistration(sheet, pjName, idTicket, name) {
  let colIndex = findPjColumnIndex(sheet, pjName);
  if (colIndex === -1) {
    colIndex = 1; // Fallback jika tidak cocok
  }

  const lastRow = Math.max(sheet.getLastRow(), 2);
  let targetRow = 3;

  if (lastRow >= 3) {
    const colValues = sheet.getRange(3, colIndex, lastRow - 2, 1).getValues();
    let foundEmpty = false;
    for (let i = 0; i < colValues.length; i++) {
      if (!colValues[i][0] || colValues[i][0].toString().trim() === '') {
        targetRow = 3 + i;
        foundEmpty = true;
        break;
      }
    }
    if (!foundEmpty) {
      targetRow = lastRow + 1;
    }
  }

  sheet.getRange(targetRow, colIndex, 1, 2).setValues([[idTicket, name]]);
}

// =========================================================================
// PENGIRIMAN E-TIKET KE EMAIL PESERTA (DESAIN STUB TICKET CISP)
// =========================================================================

// Nama acara untuk kalimat pembuka email
const EVENT_TITLE = 'SPMS The Halal Portofolio Project';

// Gambar dari Google Drive (ID = bagian di antara /d/ dan /view pada link Drive).
// Akun yang menjalankan script harus punya akses ke kedua file.
// Kalau dikosongkan, email tetap terkirim (tanpa gambar tsb).
const HEADER_IMAGE_FILE_ID = '1jcU-Fxy81S7Tk1wVf5lTr9gcPeQGlI_e'; // header di atas kartu
const BANNER_IMAGE_FILE_ID = '1kXvWa53JiOAMCIxWHai4wzuvkHa0otaW'; // banner di dalam tiket

const BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
                  'Agustus', 'September', 'Oktober', 'November', 'Desember'];

// Ubah nilai kolom H (Tanggal Pendaftaran) menjadi "14 Oktober 2026, 21.12 WIB".
// Mendukung sel bertipe tanggal (Date) maupun teks "dd/MM/yyyy HH:mm:ss".
function formatTanggalDaftar_(value) {
  if (value === null || value === undefined || value === '') return '-';

  let d, m, y, hh, mm;
  if (value instanceof Date && !isNaN(value.getTime())) {
    let tz = 'Asia/Jakarta';
    try { tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || tz; } catch (e) {}
    const parts = Utilities.formatDate(value, tz, 'd|M|yyyy|HH|mm').split('|');
    d = +parts[0]; m = +parts[1]; y = +parts[2]; hh = parts[3]; mm = parts[4];
  } else {
    const match = String(value).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})[ ,T]+(\d{1,2}):(\d{2})/);
    if (!match) return String(value).trim();
    d = +match[1]; m = +match[2]; y = +match[3];
    hh = ('0' + match[4]).slice(-2); mm = match[5];
  }
  return d + ' ' + BULAN_ID[m - 1] + ' ' + y + ', ' + hh + '.' + mm + ' WIB';
}

// Ambil header & banner dari Drive sebagai gambar inline (cid) agar tampil di Gmail tanpa hosting luar.
function getEmailInlineImages_() {
  const images = {};
  try {
    if (HEADER_IMAGE_FILE_ID) {
      images.headerImage = DriveApp.getFileById(HEADER_IMAGE_FILE_ID).getBlob().setName('header.png');
    }
  } catch (err) {
    Logger.log('Header tidak bisa dimuat dari Drive: ' + err.message);
  }
  try {
    if (BANNER_IMAGE_FILE_ID) {
      images.bannerImage = DriveApp.getFileById(BANNER_IMAGE_FILE_ID).getBlob().setName('banner.png');
    }
  } catch (err) {
    Logger.log('Banner tidak bisa dimuat dari Drive: ' + err.message);
  }
  return images;
}

// Kirim satu email tiket
function kirimEmailTiket_(idTicket, name, email, qrUrl, tglDaftar, images) {
  const htmlBody = createTicketEmailHtml(name, idTicket, qrUrl, {
    orderDate: tglDaftar,
    hasHeader: !!images.headerImage,
    hasBanner: !!images.bannerImage
  });
  const options = { htmlBody: htmlBody, name: "Panitia SPMS ISP" };
  if (Object.keys(images).length > 0) options.inlineImages = images;
  GmailApp.sendEmail(email, `[E-TIKET] Akses Masuk & Voucher Konsumsi - ${idTicket}`, "", options);
}

// Template HTML Email Tiket
//
// Seluruh layout memakai <table> + inline style agar tampil di Gmail/Outlook.
// Lubang di tepi dan lekukan di sisi dibuat dari sel berwarna CONTAINER (sama dengan
// kartu biru di belakang tiket) yang diberi border-radius. Kalau CONTAINER diganti,
// lubangnya ikut berubah warna.
//
// Tips dark mode: Gmail membalik warna latar yang terang. Latar yang memakai
// background-image gradient (lihat solid()) tidak ikut dibalik, jadi tiket dan QR
// tetap putih dan bisa discan.
//
// opts (semua opsional):
//   orderDate : isi kolom H (Date atau teks dd/MM/yyyy HH:mm:ss)
//   hasHeader, hasBanner : true jika gambar inline (cid) disertakan saat kirim
//   eventDate, venue     : info acara tambahan di bawah banner
function createTicketEmailHtml(name, idTicket, qrUrl, opts) {
  opts = opts || {};

  const esc = (s) => String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");

  // Tanggal pendaftaran diambil dari kolom H
  const orderDate = esc(formatTanggalDaftar_(opts.orderDate));

  // Palet mengikuti logo CISP: biru, abu, putih
  const BG = "#E6EEF4";         // latar email (biru sangat muda)
  const CONTAINER = "#04526F";  // kartu biru di belakang tiket
  const PAPER = "#FFFFFF";      // badan tiket
  const BLUE = "#04526F";       // teks & aksen biru
  const GRAY = "#6B7378";       // abu untuk label
  const LINE = "#BFCBD3";       // garis sobekan
  const FRAME = "#EEF3F7";      // bingkai QR

  const font = "'Helvetica Neue', Helvetica, Arial, sans-serif";

  // Warna latar yang tahan dark mode (bgcolor + gradient satu warna)
  const solid = (c) => `bgcolor="${c}" style="background-color:${c};background-image:linear-gradient(${c},${c});`;

  const eventDate = opts.eventDate ? esc(opts.eventDate) : "";
  const venue = opts.venue ? esc(opts.venue) : "";

  // Tepi berlubang: deretan setengah lingkaran berwarna CONTAINER di sepanjang tepi.
  // edge = "top" (lubang menggantung ke bawah) atau "bottom" (lubang naik ke atas)
  const perforation = (edge) => {
    const radius = edge === "top" ? "0 0 9px 9px" : "9px 9px 0 0";
    const hole = `<td width="18" height="9" ${solid(CONTAINER)}border-radius:${radius};width:18px;height:9px;font-size:0;line-height:0;">&nbsp;</td>`;
    const gap = `<td height="9" style="height:9px;font-size:0;line-height:0;">&nbsp;</td>`;
    const end = `<td width="16" height="9" style="width:16px;height:9px;font-size:0;line-height:0;">&nbsp;</td>`;
    let cells = end;
    for (let i = 0; i < 10; i++) cells += hole + (i < 9 ? gap : "");
    cells += end;
    return `<table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" ${solid(PAPER)}border-collapse:separate;"><tr>${cells}</tr></table>`;
  };


  const detail = (label, value, size) => `
    <div style="font-family:${font};font-size:12px;line-height:16px;color:${GRAY};">${label}</div>
    <div style="font-family:${font};font-size:${size}px;line-height:1.3;font-weight:700;color:${BLUE};word-break:break-word;">${value}</div>`;

  const eventLines = [
    eventDate ? `<div style="font-family:${font};font-size:14px;line-height:21px;color:#1B2B34;">${eventDate}</div>` : "",
    venue ? `<div style="font-family:${font};font-size:14px;line-height:21px;color:#1B2B34;">${venue}</div>` : ""
  ].join("");

  // Header di luar kartu (seperti logo NextDev)
  const headerImg = opts.hasHeader
    ? `<img src="cid:headerImage" alt="Investor Saham Pemula" width="220" style="display:block;margin:0 auto;width:220px;max-width:100%;height:auto;border:0;" />`
    : "";

  // Banner di bagian atas tiket; kalau tidak ada gambar, pakai tulisan
  const bannerBlock = opts.hasBanner
    ? `<img src="cid:bannerImage" alt="${esc(EVENT_TITLE)}" width="356" style="display:block;width:100%;max-width:356px;height:auto;border:0;border-radius:10px;" />`
    : `<div style="font-family:${font};font-size:36px;line-height:40px;font-weight:900;letter-spacing:3px;color:${BLUE};text-align:center;padding:10px 0;">E-TICKET</div>`;

  return `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="color-scheme" content="light only">
  <meta name="supported-color-schemes" content="light only">
  <style>:root { color-scheme: light only; supported-color-schemes: light only; }</style>
  <title>E-Ticket ${esc(idTicket)}</title>
</head>
<body style="margin:0;padding:0;background-color:${BG};">
<table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" ${solid(BG)}">
  <tr>
    <td align="center" style="padding:30px 12px 36px;">

      <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="max-width:420px;">

        <!-- Header di luar kartu -->
        ${headerImg ? `<tr><td align="center" style="padding:0 0 22px;">${headerImg}</td></tr>` : ""}

        <tr>
          <td>

            <!-- KARTU BIRU (wadah tiket) -->
            <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" ${solid(CONTAINER)}border-collapse:separate;border-radius:22px;">
              <tr>
                <td style="padding:26px 22px 24px;">

                  <!-- Kalimat pembuka -->
                  <div style="font-family:${font};font-size:20px;line-height:27px;font-weight:700;color:#FFFFFF;">
                    Congratulations Investors, you&rsquo;ve got the ticket!
                  </div>
                  <div style="font-family:${font};font-size:15px;line-height:24px;color:#DCE9F0;padding:16px 0 22px;">
                    Selamat, kamu <b style="color:#FFFFFF;">mendapatkan E-Tiket ${esc(EVENT_TITLE)}!</b><br>
                    Cek detail e-tiket kamu di bawah ini ya.
                  </div>

                  <!-- TIKET -->
                  <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" ${solid(PAPER)}border-collapse:separate;border-radius:16px;overflow:hidden;">

                    <!-- Tepi atas berlubang -->
                    <tr><td style="font-size:0;line-height:0;">${perforation("top")}</td></tr>

                    <!-- Banner -->
                    <tr>
                      <td align="center" style="padding:10px 14px 4px;">
                        ${bannerBlock}
                        ${eventLines ? `<div style="padding-top:10px;text-align:center;">${eventLines}</div>` : ""}
                      </td>
                    </tr>

                    <!-- Data peserta -->
                    <tr>
                      <td style="padding:20px 24px 20px;">
                        <div style="padding-bottom:16px;">${detail("Nama pemilik tiket", esc(name), 20)}</div>
                        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="38%" valign="top">${detail("Nomor tiket", esc(idTicket), 17)}</td>
                            <td width="62%" valign="top">${detail("Tanggal pendaftaran", orderDate, 15)}</td>
                          </tr>
                        </table>
                      </td>
                    </tr>

                    <!-- Garis sobekan + lekukan kiri/kanan -->
                    <tr>
                      <td style="font-size:0;line-height:0;">
                        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" ${solid(PAPER)}border-collapse:separate;">
                          <tr>
                            <td width="16" height="32" ${solid(CONTAINER)}width:16px;height:32px;border-radius:0 16px 16px 0;font-size:0;line-height:0;">&nbsp;</td>
                            <td height="32" valign="middle" style="height:32px;padding:0 10px;font-size:0;line-height:0;">
                              <div style="border-top:3px dotted ${LINE};font-size:0;line-height:0;height:0;">&nbsp;</div>
                            </td>
                            <td width="16" height="32" ${solid(CONTAINER)}width:16px;height:32px;border-radius:16px 0 0 16px;font-size:0;line-height:0;">&nbsp;</td>
                          </tr>
                        </table>
                      </td>
                    </tr>

                    <!-- Stub: QR -->
                    <tr>
                      <td align="center" style="padding:16px 24px 22px;">
                        <table role="presentation" align="center" border="0" cellpadding="0" cellspacing="0" ${solid(FRAME)}border-radius:18px;">
                          <tr>
                            <td align="center" style="padding:14px;">
                              <table role="presentation" border="0" cellpadding="0" cellspacing="0" ${solid("#FFFFFF")}border-radius:12px;">
                                <tr>
                                  <td style="padding:10px;">
                                    <img src="${esc(qrUrl)}" alt="QR tiket ${esc(idTicket)}" width="190" height="190"
                                         style="display:block;width:190px;height:190px;border:0;outline:none;" />
                                  </td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>

                    
                    <tr>
                      <td align="center" style="padding:0 24px 22px;">
                        <div style="font-family:${font};font-size:12px;line-height:18px;color:${GRAY};padding-top:8px;">
                          Tunjukkan QR ini ke panitia di meja registrasi untuk absensi masuk dan pengambilan voucher konsumsi.
                        </div>
                      </td>
                    </tr>

                    <!-- Tepi bawah berlubang -->
                    <tr><td style="font-size:0;line-height:0;">${perforation("bottom")}</td></tr>

                  </table>

                </td>
              </tr>
            </table>

          </td>
        </tr>
      </table>

    </td>
  </tr>
</table>
</body>
</html>`;
}

// 5. TEST FUNCTION: Uji coba pengiriman HANYA untuk 2 peserta mock (TKT-112 dan TKT-113)
function testKirimEmailMock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ticketSheet = ss.getSheetByName(SHEET_TICKET_NAME);
  const data = ticketSheet.getDataRange().getValues();
  const images = getEmailInlineImages_();

  const mockTickets = ['TKT-112', 'TKT-113'];
  let sentCount = 0;

  for (let r = 1; r < data.length; r++) {
    const idTicket = data[r][0] ? data[r][0].toString().trim() : '';
    const name = data[r][1] ? data[r][1].toString().trim() : '';
    const email = data[r][2] ? data[r][2].toString().trim() : '';
    const qrUrl = data[r][3] ? data[r][3].toString().trim() : '';
    const tglDaftar = data[r][7]; // Kolom H: Tanggal Pendaftaran

    if (mockTickets.includes(idTicket)) {
      if (email && email.includes('@')) {
        kirimEmailTiket_(idTicket, name, email, qrUrl, tglDaftar, images);
        Logger.log(`Terkirim ke mock: ${idTicket} - ${name} (${email})`);
        sentCount++;
      }
    }
  }
  Logger.log(`Uji coba selesai. Total email terkirim: ${sentCount}`);
}

// 6. BROADCAST MASSAL: Kirim ke semua peserta 
function kirimSemuaTiketEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ticketSheet = ss.getSheetByName(SHEET_TICKET_NAME);
  const data = ticketSheet.getDataRange().getValues();
  const images = getEmailInlineImages_();
  let count = 0;

  for (let r = 1; r < data.length; r++) {
    const idTicket = data[r][0] ? data[r][0].toString().trim() : '';
    const name = data[r][1] ? data[r][1].toString().trim() : '';
    const email = data[r][2] ? data[r][2].toString().trim() : '';
    const qrUrl = data[r][3] ? data[r][3].toString().trim() : '';
    const tglDaftar = data[r][7]; // Kolom H: Tanggal Pendaftaran

    if (idTicket && email && email.includes('@')) {
      kirimEmailTiket_(idTicket, name, email, qrUrl, tglDaftar, images);
      count++;
      Utilities.sleep(400); //rate limit
    }
  }
  Logger.log(`Kirim serentak selesai. Total ${count} email diproses.`);
}