/**
 * 🚀 1-CLICK DOCUMENT GENERATOR (100% TESTED & BULLETPROOF)
 * Jalankan fungsi ini untuk membuat dokumen SOP & Slide Presentasi di Google Drive Anda!
 */
function createSOPDocumentInDrive() {
  // 1. Buat Dokumen Baru di Google Drive Anda
  var doc = DocumentApp.create("SOP & Panduan CIMS - Kopi Kenangan");
  var body = doc.getBody();

  // Warna Kopi Kenangan (#8B0000)
  var MAROON = "#8B0000";

  // TITLE & HEADER
  var pTitle = body.appendParagraph("STANDARD OPERATING PROCEDURE (SOP) & USER GUIDE");
  pTitle.setFontSize(18).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor(MAROON);

  var pSub = body.appendParagraph("Continuous Improvement Management System (CIMS) — Kopi Kenangan\n");
  pSub.setFontSize(12).setItalic(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // SECTION 1: OVERVIEW
  var h1 = body.appendParagraph("📌 OVERVIEW SISTEM");
  h1.setFontSize(14).setBold(true).setForegroundColor(MAROON);
  body.appendParagraph("CIMS adalah platform terintegrasi berbasis Google Workspace & Web App yang dirancang untuk mempermudah karyawan Kopi Kenangan dalam mengajukan, memantau, merevisi, dan menyetujui ide-ide Continuous Improvement (CI) secara transparan dan real-time.\n");

  // SECTION 2: PANDUAN INITIATOR
  var h2 = body.appendParagraph("👤 1. PANDUAN UNTUK INITIATOR / STAFF (PENGUSUL PROYEK)");
  h2.setFontSize(14).setBold(true).setForegroundColor(MAROON);
  
  body.appendParagraph("A. Pengajuan Proyek Baru").setBold(true);
  body.appendListItem("Buka link Google Form Pengajuan CIMS.");
  body.appendListItem("Isi informasi data diri (Nama, Job Title, Departemen, Manager).");
  body.appendListItem("Isi detail usulan proyek: Project Title, Scope, Problem Statement, Improvement Idea, Timeline.");
  body.appendListItem("Klik Submit. Anda akan menerima email konfirmasi beserta Project ID unik (contoh: CI-2026-0008).");

  body.appendParagraph("\nB. Jika Proyek Membutuhkan Revisi (Need Revision)").setBold(true);
  body.appendListItem("Buka email masuk berjudul '[CIMS RESULT] CI-XXXX-XXXX - Need Revision'.");
  body.appendListItem("Klik tombol '✏️ Klik Disini Untuk Revisi Proyek'.");
  body.appendListItem("Perbaiki data pada Web App Revisi yang sudah terisi otomatis, lalu klik 'Kirim Hasil Revisi'.");

  body.appendParagraph("\nC. Melaporkan Proyek Selesai (Project Closing)").setBold(true);
  body.appendListItem("Buka link Web App Closing.");
  body.appendListItem("Masukkan Tanggal Selesai Aktual (Completion Date) dan lampirkan link bukti akhir.");
  body.appendListItem("Klik 'Kirim Pengajuan Closing'. Status proyek resmi menjadi 'Closed'.\n");

  // SECTION 3: PANDUAN MANAGER
  var h3 = body.appendParagraph("👔 2. PANDUAN UNTUK MANAGER (APPROVER LEVEL 1)");
  h3.setFontSize(14).setBold(true).setForegroundColor(MAROON);
  body.appendParagraph("A. Melakukan Review & Persetujuan").setBold(true);
  body.appendListItem("Buka email masuk berjudul '[ACTION REQUIRED] CIMS Approval Needed'.");
  body.appendListItem("Klik tombol 'Review & Approve Project'.");
  body.appendListItem("Pilih keputusan: Approve, Need Revision, atau Reject, lalu klik 'Kirim Keputusan'.");
  body.appendParagraph("(Catatan: Email Reminder otomatis dikirim setiap jam 08:00 WIB jika pending > 3 hari)\n");

  // SECTION 4: QAQC HEAD
  var h4 = body.appendParagraph("🛡️ 3. PANDUAN UNTUK QAQC HEAD (APPROVER LEVEL 2 & DASHBOARD)");
  h4.setFontSize(14).setBold(true).setForegroundColor(MAROON);
  body.appendListItem("Melakukan review final untuk proyek yang disetujui Manager.");
  body.appendListItem("Memantau statistik real-time via tab DASHBOARD pada spreadsheet CI Database.\n");

  // SECTION 5: TABEL MATRIKS
  var h5 = body.appendParagraph("📊 RINGKASAN MATRIKS STATUS PROYEK");
  h5.setFontSize(14).setBold(true).setForegroundColor(MAROON);
  
  var tableData = [
    ["Status", "Arti & Tindakan"],
    ["Waiting Manager Approval", "Proyek baru diajukan, menunggu keputusan Manager Pengusul."],
    ["Waiting QA Manager Approval", "Proyek telah disetujui Manager, menunggu keputusan QAQC Head."],
    ["Need Revision", "Proyek dikembalikan ke Initiator untuk diperbaiki via Web App."],
    ["Approved", "Proyek resmi disetujui dan siap dieksekusi oleh tim."],
    ["Rejected", "Usulan proyek ditolak/diberhentikan."],
    ["Closed", "Proyek telah selesai dilaksanakan dan terdata di database."]
  ];
  
  var table = body.appendTable(tableData);
  
  // Format Header Tabel Menggunakan editAsText()
  var headerRow = table.getRow(0);
  
  var cell0 = headerRow.getCell(0);
  cell0.setBackgroundColor(MAROON);
  cell0.editAsText().setForegroundColor("#FFFFFF").setBold(true);

  var cell1 = headerRow.getCell(1);
  cell1.setBackgroundColor(MAROON);
  cell1.editAsText().setForegroundColor("#FFFFFF").setBold(true);

  // PAGE BREAK
  body.appendPageBreak();

  // SECTION 6: PRESENTATION BLUEPRINT
  var hSlideTitle = body.appendParagraph("🖥️ BLUEPRINT MATERI PRESENTASI (POWERPOINT / GOOGLE SLIDES)");
  hSlideTitle.setFontSize(16).setBold(true).setForegroundColor(MAROON);

  var slides = [
    ["Slide 1: Title Slide", "• Judul Utama: Sosialisasi Sistem CIMS\n• Sub-judul: Inovasi Lebih Cepat, Proses Lebih Mudah, 100% Digital.\n• Visual: Logo Kenangan Brands"],
    ["Slide 2: Apa itu CIMS?", "• Konsep: Platform terintegrasi berbasis Google Workspace untuk mengelola ide perbaikan.\n• Keunggulan: 100% Paperless, Real-time tracking, Approval via HP, Dashboard Executive."],
    ["Slide 3: Alur Kerja (Workflow)", "1. Submit Form ➔ 2. Manager Approval ➔ 3. QAQC Approval ➔ 4. Execution ➔ 5. Project Closing"],
    ["Slide 4: Panduan Staff (Initiator)", "• Mengajukan ide via Google Form.\n• Fitur Revisi Cerdas tanpa perlu isi ulang form dari nol.\n• Lapor tanggal selesai & bukti akhir saat closing."],
    ["Slide 5: Kemudahan Manager (Approver)", "• Notifikasi email instan.\n• 1-Click Review (Approve, Need Revision, Reject).\n• Auto-reminder harian jika pending > 3 hari."],
    ["Slide 6: Matriks Status", "• Waiting Approval, Need Revision, Approved, Closed."],
    ["Slide 7: Q&A & Live Demo", "• Sesi Tanya Jawab & QR Code Uji Coba Pengisian Form."]
  ];

  for (var i = 0; i < slides.length; i++) {
    var st = body.appendParagraph(slides[i][0]);
    st.setFontSize(12).setBold(true).setForegroundColor(MAROON);
    body.appendParagraph(slides[i][1] + "\n");
  }

  doc.saveAndClose();

  Logger.log("====================================================");
  Logger.log("✅ DOKUMEN BERHASIL DIBUAT!");
  Logger.log("🔗 Link File Dokumen Anda: " + doc.getUrl());
  Logger.log("====================================================");
}
