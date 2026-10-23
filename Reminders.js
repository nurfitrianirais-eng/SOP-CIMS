/**
 * =================================================================
 * CIMS REMINDER SYSTEM (WITH AUTOMATIC MEMBER CC SUPPORT)
 * File: reminder.gs
 * =================================================================
 */

/**
 * 1. PENGINGAT BULANAN OTOMATIS UNTUK INITIATOR & MEMBER TIM (PROJECT ON PROGRESS)
 * Mengirimkan email status & target timeline kepada Initiator (To) dan Member Tim (CC) setiap bulan
 */
function sendMonthlyInitiatorReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("MASTER_PROJECT");

  if (!masterSheet) {
    Logger.log("Error: Sheet MASTER_PROJECT tidak ditemukan!");
    return;
  }

  var data = masterSheet.getDataRange().getValues();
  var sentCount = 0;

  for (var i = 1; i < data.length; i++) {
    var projectTitle = String(data[i][2] || "").trim();    // Kolom C: Project Title
    var initiatorName = String(data[i][3] || "").trim();   // Kolom D: Initiator Name
    var initiatorEmail = String(data[i][5] || "").trim().toLowerCase(); // Kolom F: Email
    var membersStr = String(data[i][7] || "").trim();     // Kolom H: Members Tim
    var status = String(data[i][17] || "").trim();        // Kolom R: Status
    var startDate = data[i][1];                            // Kolom B: Submit Date
    var dueDate = data[i][18];                             // Kolom S: Due Date

    // Hanya kirim jika status "Approved" dan email mengandung karakter "@"
    if (status === "Approved" && initiatorEmail.indexOf("@") !== -1) {
      var formattedStart = startDate ? Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd MMM yyyy") : "-";
      var formattedDue = dueDate ? Utilities.formatDate(new Date(dueDate), Session.getScriptTimeZone(), "dd MMM yyyy") : "TBD";

      // 👥 Tarik daftar email member tim untuk di-CC
      var memberEmails = (typeof getMemberEmails_ === "function") ? getMemberEmails_(membersStr) : "";

      var subject = "☕ [CIMS Kopi Kenangan] Monthly Reminder: Ongoing Project Update - " + projectTitle;
      
      var htmlBody = `
        <div style="font-family: Arial, sans-serif; color: #333333; max-width: 600px; border: 1px solid #E0E0E0; padding: 20px; border-radius: 8px;">
          <h2 style="color: #8B0000; margin-top: 0;">☕ CIMS Monthly Project Reminder</h2>
          <p>Halo <strong>${initiatorName} & Team</strong>,</p>
          <p>Ini adalah pengingat bulanan rutin dari sistem <strong>Continuous Improvement Management System (CIMS) Kopi Kenangan</strong> untuk proyek perbaikan yang sedang kamu jalankan.</p>
          
          <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #F9F9F9;">
            <tr>
              <td style="padding: 10px; border: 1px solid #DDDDDD; font-weight: bold; width: 35%;">Nama Project</td>
              <td style="padding: 10px; border: 1px solid #DDDDDD; color: #8B0000; font-weight: bold;">${projectTitle}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #DDDDDD; font-weight: bold;">Status Saat Ini</td>
              <td style="padding: 10px; border: 1px solid #DDDDDD;"><span style="background-color: #28A745; color: white; padding: 3px 8px; border-radius: 4px; font-size: 12px;">${status}</span></td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #DDDDDD; font-weight: bold;">Planned Timeline</td>
              <td style="padding: 10px; border: 1px solid #DDDDDD;">${formattedStart} s/d ${formattedDue}</td>
            </tr>
          </table>

          <p>Mohon pastikan proyek berjalan sesuai target lini masa yang diajukan. Apabila proyek telah rampung 100%, jangan lupa untuk melakukan pengajuan <strong>Project Closing</strong> melalui portal CIMS.</p>
          
          <br>
          <p style="font-size: 12px; color: #777777;">Email ini dikirimkan secara otomatis oleh CIMS Internal QA/QC Kopi Kenangan System kepada Inisiator dan Anggota Tim.</p>
        </div>
      `;

      var mailOptions = {
        to: initiatorEmail,
        subject: subject,
        htmlBody: htmlBody,
        name: "CIMS Kopi Kenangan System"
      };

      // Tambahkan CC ke member jika ada
      if (memberEmails) {
        mailOptions.cc = memberEmails;
      }

      MailApp.sendEmail(mailOptions);
      sentCount++;
    }
  }

  Logger.log("✅ Berhasil mengirim " + sentCount + " email pengingat bulanan ke Initiator & CC Member.");
}

/**
 * 2. PENGINGAT HARIAN OTOMATIS (DEADLINE H-7, HARI H, OVERDUE)
 * Mengecek setiap hari, HANYA mengirim email jika sisa waktu <= 7 hari (Di-CC ke Member Tim).
 */
function sendAutoClosingReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MASTER_PROJECT");
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var webAppUrl = (typeof WEB_APP_URL !== "undefined") ? WEB_APP_URL : "https://script.google.com/a/macros/kopikenangan.com/s/AKfycbzzVDRlzp97MeKa5RjpUxFKm3P_nGn84_ZJjDaPv36NZtJSyUZyEJsE_VTnlcHk7omi/exec"; 
  var sentCount = 0;

  for (var i = 1; i < data.length; i++) {
    var projectId = String(data[i][0] || "").trim();    // Kolom A: Project ID
    var projectTitle = String(data[i][2] || "").trim(); // Kolom C: Project Title
    var initiatorName = String(data[i][3] || "").trim();// Kolom D: Initiator Name
    var initiatorEmail = String(data[i][5] || "").trim().toLowerCase(); // Kolom F: Email
    var membersStr = String(data[i][7] || "").trim();   // Kolom H: Members Tim
    var status = String(data[i][17] || "").trim();      // Kolom R: Status
    var targetDateRaw = data[i][18];                    // Kolom S: Due Date

    // Hanya cek yang "Approved", punya email valid, dan Due Date-nya format Tanggal
    if (status === "Approved" && initiatorEmail.indexOf("@") !== -1 && targetDateRaw instanceof Date) {
      
      var targetDate = new Date(targetDateRaw);
      targetDate.setHours(0, 0, 0, 0);
      
      var diffTime = targetDate.getTime() - today.getTime();
      var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      // Jika target penyelesaian tinggal <= 7 Hari (Atau sudah lewat / minus)
      if (diffDays <= 7) {
        
        // 👥 Tarik daftar email member tim untuk di-CC
        var memberEmails = (typeof getMemberEmails_ === "function") ? getMemberEmails_(membersStr) : "";

        var closingLink = webAppUrl + "?action=close&id=" + projectId;
        var subject = "";
        var statusBadge = "";
        var color = "";

        if (diffDays > 0) {
          subject = "⏳ [CIMS] Action Required: Deadline Proyek H-" + diffDays + " - " + projectTitle;
          statusBadge = "Mendekati Tenggat Waktu (H-" + diffDays + ")";
          color = "#D97706"; // Orange (Warning)
        } else if (diffDays === 0) {
          subject = "🚨 [CIMS] DUE TODAY: Tenggat Waktu Proyek CIMS Hari Ini!";
          statusBadge = "TENGGAT WAKTU HARI INI";
          color = "#DC3545"; // Merah Terang (Urgent)
        } else {
          subject = "⚠️ [CIMS] OVERDUE: Proyek CIMS Terlambat " + Math.abs(diffDays) + " Hari!";
          statusBadge = "TERLAMBAT " + Math.abs(diffDays) + " HARI";
          color = "#8B0000"; // Merah Gelap Kopi Kenangan (Critical)
        }

        var formattedDue = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "dd MMM yyyy");

        var htmlBody = `
        <div style="font-family: Arial, sans-serif; color: #333333; max-width: 600px; border: 1px solid #E0E0E0; padding: 20px; border-radius: 8px;">
          <h2 style="color: #8B0000; margin-top: 0;">☕ CIMS Deadline Reminder</h2>
          <p>Halo <strong>${initiatorName} & Team</strong>,</p>
          <p>Sistem mendeteksi bahwa proyek perbaikan (CIMS) yang kamu pimpin saat ini berstatus: <strong>${statusBadge.toLowerCase()}</strong>.</p>
          
          <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #F9F9F9;">
            <tr>
              <td style="padding: 10px; border: 1px solid #DDDDDD; font-weight: bold; width: 35%;">Project ID</td>
              <td style="padding: 10px; border: 1px solid #DDDDDD;">${projectId}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #DDDDDD; font-weight: bold;">Nama Project</td>
              <td style="padding: 10px; border: 1px solid #DDDDDD; color: #8B0000; font-weight: bold;">${projectTitle}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #DDDDDD; font-weight: bold;">Target Selesai</td>
              <td style="padding: 10px; border: 1px solid #DDDDDD;"><span style="color: ${color}; font-weight: bold;">${formattedDue}</span></td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #DDDDDD; font-weight: bold;">Status Waktu</td>
              <td style="padding: 10px; border: 1px solid #DDDDDD;"><span style="background-color: ${color}; color: white; padding: 3px 8px; border-radius: 4px; font-size: 12px; font-weight: bold;">${statusBadge}</span></td>
            </tr>
          </table>

          <p>Jika proyek ini <strong>SUDAH SELESAI</strong> diimplementasikan, mohon segera laporkan hasil penghematan/benefitnya (beserta bukti laporan) melalui tombol di bawah ini, agar proyekmu bisa segera ditutup dan diverifikasi oleh Manager:</p>
          
          <div style="text-align: center; margin: 25px 0;">
            <a href="${closingLink}" style="background-color: #8B0000; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 14px; display: inline-block;">AJUKAN PROJECT CLOSING</a>
          </div>
          
          <br>
          <p style="font-size: 12px; color: #777777; border-top: 1px solid #EEEEEE; padding-top: 10px;">
            Email ini dikirimkan secara otomatis oleh CIMS Internal QA/QC Kopi Kenangan System kepada Inisiator dan Anggota Tim.<br>
            Jika butuh perpanjangan waktu penyelesaian, mohon laporkan/hubungi QA/QC Manager Anda.
          </p>
        </div>
        `;

        var mailOptions = {
          to: initiatorEmail,
          subject: subject,
          htmlBody: htmlBody,
          name: "CIMS Kopi Kenangan System"
        };

        // Tambahkan CC ke member jika ada
        if (memberEmails) {
          mailOptions.cc = memberEmails;
        }

        MailApp.sendEmail(mailOptions);
        sentCount++;
      }
    }
  }
  Logger.log("✅ Berhasil mengirim " + sentCount + " email reminder DEADLINE ke Initiator & CC Member.");
}