/*************************************************
 * MAIL.GS
 * CIMS v1.3 (Fixed Project Getter & 2-Tier Approval)
 *************************************************/

/*************************************************
 * Render Email HTML (For Proposal Approval)
 *************************************************/
function buildApprovalEmail_(project, approver, approvalUrl, currentStatus, progressHtml) {
  const template = HtmlService.createTemplateFromFile("EmailTemplate");
  template.project = project;
  template.approver = approver;
  template.logo = CONFIG.LOGO_URL;
  template.company = CONFIG.COMPANY_NAME;
  template.systemName = CONFIG.SYSTEM_NAME;
  template.systemDesc = CONFIG.SYSTEM_DESC;
  template.currentStatus = currentStatus;
  template.progress = progressHtml;
  template.approvalUrl = approvalUrl;

  return template.evaluate().getContent();
}

/*************************************************
 * Approval Progress (For Proposal)
 *************************************************/
function buildProgress_(order) {
  if (order == 1) {
    return `🟢 QAQC Head<br>⚪ QA Manager<br>⚪ QC Manager`;
  }
  if (order == 2) {
    return `✅ QAQC Head<br>🟢 QA Manager<br>⚪ QC Manager`;
  }
  return `✅ QAQC Head<br>✅ QA Manager<br>🟢 QC Manager`;
}

/*************************************************
 * Send Approval Email (For Proposal)
 *************************************************/
function sendApprovalEmail_(projectId, nextOrder) {
  const project = getProjectDetailsByID_(projectId); // FIXED GETTER
  if (!project) {
    Logger.log("Project tidak ditemukan : " + projectId);
    return;
  }

  const approver = getApproverByOrder_(nextOrder);
  if (!approver) {
    Logger.log("Approver tidak ditemukan.");
    return;
  }

  let approverCode = "";
  switch (nextOrder) {
    case 1: approverCode = "manager"; break;
    case 2: approverCode = "qa"; break;
    case 3: approverCode = "qc"; break;
    default: approverCode = "manager";
  }

  // Handle case where project.projectId might be project.id based on the getter returned object
  const safeId = project.projectId || project.id || projectId;
  const approvalUrl = CONFIG.WEB_APP_URL + "?id=" + safeId + "&approver=" + approverCode;

  let currentStatus = "";
  switch (nextOrder) {
    case 1: currentStatus = CONFIG.STATUS_WAITING_MANAGER; break;
    case 2: currentStatus = CONFIG.STATUS_WAITING_QA; break;
    case 3: currentStatus = CONFIG.STATUS_WAITING_QC; break;
  }

  const progressHtml = buildProgress_(nextOrder);
  const htmlBody = buildApprovalEmail_(project, approver, approvalUrl, currentStatus, progressHtml);
  const subject = "[CIMS] Approval Required - " + safeId + " - " + project.title;

  GmailApp.sendEmail(approver.email, subject, "Your email client does not support HTML email.", {
    htmlBody: htmlBody,
    name: CONFIG.SYSTEM_NAME
  });

  Logger.log("Approval email sent to : " + approver.email);
}

/*************************************************
 * Build Result Email
 *************************************************/
function buildResultEmail_(project, finalStatus) {
  const color = (finalStatus == CONFIG.STATUS_APPROVED) ? "#2E7D32" : "#C62828";
  const safeId = project.projectId || project.id;
  const template = `
  <div style="font-family:Arial,sans-serif;max-width:800px;margin:auto;border:1px solid #ddd;border-radius:10px;overflow:hidden;">
    <div style="background:${color};color:white;padding:25px;text-align:center;">
      <img src="${CONFIG.LOGO_URL}" style="width:180px;"><br><br>
      <h2 style="margin:0;">${CONFIG.SYSTEM_NAME}</h2>
      <div>${CONFIG.SYSTEM_DESC}</div>
    </div>
    <div style="padding:30px;">
      <p>Dear <b>${project.initiator}</b>,</p>
      <p>Your Continuous Improvement Project has been processed.</p>
      <table style="width:100%;border-collapse:collapse;">
        <tr><td style="padding:8px;font-weight:bold;width:180px;">Project ID</td><td>${safeId}</td></tr>
        <tr><td style="padding:8px;font-weight:bold;">Project Title</td><td>${project.title}</td></tr>
        <tr>
          <td style="padding:8px;font-weight:bold;">Final Status</td>
          <td>
            <span style="background:${color};color:white;padding:6px 12px;border-radius:5px;">
              ${finalStatus}
            </span>
          </td>
        </tr>
      </table>
      <br>
      <div style="background:#F5F5F5;padding:18px;border-left:5px solid ${color};">
          Thank you for contributing to Continuous Improvement.
      </div>
    </div>
    <div style="background:#f5f5f5;padding:20px;text-align:center;font-size:12px;color:#777;">
      Powered by <b>${CONFIG.SYSTEM_NAME}</b><br>QAQC Department<br>${CONFIG.COMPANY_NAME}
    </div>
  </div>
  `;
  return template;
}

/*************************************************
 * Send Result Email
 *************************************************/
function sendResultEmail_(projectId, finalStatus) {
  const project = getProjectDetailsByID_(projectId); // FIXED GETTER
  if (!project) return;
  const email = getInitiatorEmail_(projectId);
  if (!email) return;

  const html = buildResultEmail_(project, finalStatus);
  const safeId = project.projectId || project.id || projectId;
  
  GmailApp.sendEmail(email, "[CIMS] " + finalStatus + " - " + safeId, "HTML Email", {
    htmlBody: html,
    name: CONFIG.SYSTEM_NAME
  });

  Logger.log("Result email sent to : " + email);
}

/*************************************************
 * HELPER: Dapatkan email Manager dari MASTER_PROJECT / MASTER_EMPLOYEE
 *************************************************/
function getInitiatorManagerEmail_(projectIdOrManagerName) {
  if (!projectIdOrManagerName) return "";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var projSheet = ss.getSheetByName("MASTER_PROJECT");
  var empSheet = ss.getSheetByName("MASTER_EMPLOYEE");
  var managerName = String(projectIdOrManagerName).trim();

  if (managerName.toUpperCase().indexOf("CI-") === 0 && projSheet) {
    var projData = projSheet.getDataRange().getValues();
    var searchId = managerName.toLowerCase();
    for (var i = 1; i < projData.length; i++) {
      if (String(projData[i][0] || "").trim().toLowerCase() === searchId) {
        managerName = String(projData[i][10] || "").trim(); 
        break;
      }
    }
  }

  if (!managerName || !empSheet) return "";
  var empData = empSheet.getDataRange().getValues();
  var searchManager = managerName.toLowerCase();

  for (var j = 1; j < empData.length; j++) {
    var empName = String(empData[j][0] || "").trim().toLowerCase(); 
    var empEmail = String(empData[j][2] || "").trim();             
    if (empName && (empName === searchManager || searchManager.indexOf(empName) !== -1 || empName.indexOf(searchManager) !== -1)) {
      if (empEmail && empEmail.indexOf("@") !== -1) return empEmail;
    }
  }
  return "";
}

/*************************************************
 * Send Closing Approval Email (TAHAP 1)
 *************************************************/
function sendClosingApprovalEmail_(projectId, formData) {
  try {
    const projectObj = getProjectDetailsByID_(projectId); // FIXED GETTER
    if (!projectObj) return;

    const projectTitle = projectObj.title || "";
    const initiatorName = projectObj.initiator || "";
    const managerName = projectObj.managerName || projectObj.manager || "";
    const remarksText = (formData && formData.finalEvidence) ? formData.finalEvidence : (projectObj.remarks || "-");

    const managerEmail = getInitiatorManagerEmail_(projectId) || (typeof getManagerEmailByName_ === "function" ? getManagerEmailByName_(managerName) : "");
    if (!managerEmail) {
      Logger.log("⚠️ Email Manager tidak ditemukan untuk Project: " + projectId);
      return;
    }

    const subject = "[CIMS Approval Required] Pengajuan Project Closing - " + projectId + " (" + projectTitle + ")";
    const webAppUrl = ScriptApp.getService().getUrl();
    const approveUrl = webAppUrl + "?action=closing_approval&id=" + encodeURIComponent(projectId) + "&decision=approve";
    const rejectUrl = webAppUrl + "?action=closing_approval&id=" + encodeURIComponent(projectId) + "&decision=reject";

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; max-width: 600px;">
        <h2 style="color: #8B0000; margin-top: 0;">☕ CIMS Project Closing Approval</h2>
        <p>Halo <b>${managerName || 'Manager'}</b>,</p>
        <p>Inisiator <b>${initiatorName}</b> telah mengajukan permohonan <b>Project Closing</b> untuk proyek berikut:</p>
        <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 14px;">
          <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px; font-weight: bold; width: 35%;">Project ID</td><td style="padding: 8px;">${projectId}</td></tr>
          <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px; font-weight: bold;">Nama Project</td><td style="padding: 8px;">${projectTitle}</td></tr>
          <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px; font-weight: bold;">Evidence & Summary</td><td style="padding: 8px;">${remarksText}</td></tr>
        </table>
        <p>Silakan klik salah satu tombol di bawah untuk memberikan keputusan persetujuan:</p>
        <div style="margin-top: 25px; text-align: center;">
          <a href="${approveUrl}" style="background-color: #28a745; color: white; padding: 12px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; margin-right: 10px; display: inline-block;">APPROVE CLOSING</a>
          <a href="${rejectUrl}" style="background-color: #dc3545; color: white; padding: 12px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">REJECT / REVISE</a>
        </div>
      </div>
    `;

    const systemName = (typeof CONFIG !== "undefined" && CONFIG.SYSTEM_NAME) ? CONFIG.SYSTEM_NAME : "CIMS Kopi Kenangan";
    GmailApp.sendEmail(managerEmail, subject, "Silakan buka email ini menggunakan HTML viewer.", {
      htmlBody: htmlBody,
      name: systemName + " - QAQC"
    });

    Logger.log("✅ Closing Approval email sent to : " + managerEmail);
  } catch (err) {
    Logger.log("❌ Error sendClosingApprovalEmail_: " + err.toString());
  }
}

/*************************************************
 * Kirim Email Final Approval ke Pak Bayun (Tahap 2)
 *************************************************/
function sendFinalClosingApprovalEmail_(projectId) {
  try {
    const projectObj = getProjectDetailsByID_(projectId); // FIXED GETTER
    if (!projectObj) return;

    const projectTitle = projectObj.title || "";
    const initiatorName = projectObj.initiator || "";
    const managerName = projectObj.managerName || projectObj.manager || "";
    const remarksText = projectObj.remarks || "-";

    const bayunEmail = "bayun.binantoro@kopikenangan.com"; 

    const subject = "[FINAL APPROVAL REQUIRED] Project Closing - " + projectId + " (" + projectTitle + ")";
    const webAppUrl = ScriptApp.getService().getUrl();
    const approveUrl = webAppUrl + "?action=closing_approval&id=" + encodeURIComponent(projectId) + "&decision=approve";
    const rejectUrl = webAppUrl + "?action=closing_approval&id=" + encodeURIComponent(projectId) + "&decision=reject";

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; max-width: 600px;">
        <h2 style="color: #0056b3; margin-top: 0;">☕ CIMS Final Closing Approval</h2>
        <p>Halo <b>Pak Bayun</b>,</p>
        <p>Proyek berikut telah disetujui oleh Manager (<b>${managerName}</b>) dan saat ini membutuhkan <b>Final Approval</b> dari Anda untuk ditutup secara permanen (Closed):</p>
        <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 14px;">
          <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px; font-weight: bold; width: 35%;">Project ID</td><td style="padding: 8px;">${projectId}</td></tr>
          <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px; font-weight: bold;">Nama Project</td><td style="padding: 8px;">${projectTitle}</td></tr>
          <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px; font-weight: bold;">Inisiator</td><td style="padding: 8px;">${initiatorName}</td></tr>
          <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px; font-weight: bold;">Evidence & Summary</td><td style="padding: 8px;">${remarksText}</td></tr>
        </table>
        <p>Silakan klik salah satu tombol di bawah untuk memberikan keputusan final:</p>
        <div style="margin-top: 25px; text-align: center;">
          <a href="${approveUrl}" style="background-color: #28a745; color: white; padding: 12px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; margin-right: 10px; display: inline-block;">FINAL APPROVE (CLOSE PROJECT)</a>
          <a href="${rejectUrl}" style="background-color: #dc3545; color: white; padding: 12px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">REJECT / REVISE</a>
        </div>
      </div>
    `;

    const systemName = (typeof CONFIG !== "undefined" && CONFIG.SYSTEM_NAME) ? CONFIG.SYSTEM_NAME : "CIMS Kopi Kenangan";
    GmailApp.sendEmail(bayunEmail, subject, "Silakan buka email ini menggunakan HTML viewer.", {
      htmlBody: htmlBody,
      name: systemName + " - Final Approval"
    });

    Logger.log("✅ Final Approval email sent to Pak Bayun : " + bayunEmail);
  } catch (err) {
    Logger.log("❌ Error sendFinalClosingApprovalEmail_: " + err.toString());
  }
}

/*************************************************
 * Send Closing Result Email (To Initiator)
 *************************************************/
function sendClosingResultEmail_(projectId, isApproved, note) {
  const project = getProjectDetailsByID_(projectId); // FIXED GETTER
  if (!project) return;
  
  const email = getInitiatorEmail_(projectId);
  if (!email) return;

  const statusStr = isApproved ? "CLOSED (Telah Disetujui)" : "NEED REVISION (Ditolak/Revisi)";
  const color = isApproved ? "#28A745" : "#DC3545";
  const htmlBody = `
  <div style="font-family:Arial,sans-serif;max-width:800px;margin:auto;border:1px solid #ddd;border-radius:10px;overflow:hidden;">
    <div style="background:${color};color:white;padding:20px;text-align:center;">
      <h2 style="margin:0;">Informasi Penutupan Proyek</h2>
    </div>
    <div style="padding:30px;">
      <p>Dear <b>${project.initiator}</b>,</p>
      <p>Pengajuan Closing untuk proyek Anda telah direview oleh Manager / Final Approver.</p>
      <p>Status Penutupan: <strong style="color:${color};">${statusStr}</strong></p>
      <div style="background:#F5F5F5;padding:15px;border-left:5px solid ${color};">
        <strong>Catatan Approver:</strong><br>
        ${note || 'Tidak ada catatan.'}
      </div>
    </div>
  </div>
  `;

  const subject = "[CIMS CLOSING] " + statusStr + " - " + projectId;
  GmailApp.sendEmail(email, subject, "Silakan buka email ini menggunakan HTML viewer.", {
    htmlBody: htmlBody,
    name: (typeof CONFIG !== "undefined" && CONFIG.SYSTEM_NAME) ? CONFIG.SYSTEM_NAME : "CIMS Kopi Kenangan"
  });

  Logger.log("Closing Result email sent to : " + email);
}

/**
 * UTILITY: Tembak Ulang Email Closing Approval yang Sempat Tertunda (Tendy / Ellyta)
 */
function triggerPendingClosingEmails() {
  // Ganti array di bawah dengan ID proyek yang mau ditembak ulang (misal Tendy: CI-2026-0031)
  var pendingIds = ["CI-2026-0003"]; 
  
  for (var i = 0; i < pendingIds.length; i++) {
    var pId = pendingIds[i];
    var pData = getProjectDetailsByID_(pId);
    
    if (pData && typeof sendClosingApprovalEmail_ === "function") {
      var formData = {
        projectId: pId,
        finalEvidence: pData.remarks || "-",
        closingSummary: pData.remarks || "-",
        benefitCategory: pData.impact || "Uncategorized",
        claimedSaving: pData.impact || "0",
        completionDate: pData.dueDate || new Date().toLocaleDateString()
      };
      sendClosingApprovalEmail_(pId, formData);
      Logger.log("✅ Email tembak ulang sukses untuk Project: " + pId);
    } else {
      Logger.log("❌ Gagal menembak ulang email untuk " + pId + ". Pastikan ID benar.");
    }
  }
}
/**
 * =========================================================
 * ⏰ TRIGGER OTOMATIS: PENGINGAT CLOSING APPROVAL YANG TERTUNDA
 * =========================================================
 */
function sendPendingClosingApprovalReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var projectSheet = ss.getSheetByName("MASTER_PROJECT");
  if (!projectSheet) return;

  var data = projectSheet.getDataRange().getValues();
  var webAppUrl = ScriptApp.getService().getUrl();
  var reminderCount = 0;

  for (var i = 1; i < data.length; i++) {
    var projectId = String(data[i][0] || "").trim();      // Kolom A
    var projectTitle = String(data[i][2] || "").trim();   // Kolom C
    var initiatorName = String(data[i][3] || "").trim();  // Kolom D
    var managerName = String(data[i][10] || "").trim();   // Kolom K
    var status = String(data[i][17] || "").trim();        // Kolom R
    var remarksText = String(data[i][21] || "-").trim();  // Kolom V

    if (!projectId) continue;

    var recipientEmail = "";
    var recipientName = "";
    var stageTitle = "";

    // 1. CEK SIAPA YANG BELUM APPROVE
    if (status === "Waiting Closing Approval") {
      // Pending di Manager
      recipientEmail = (typeof getInitiatorManagerEmail_ === "function") ? getInitiatorManagerEmail_(projectId) : "";
      if (!recipientEmail && typeof getManagerEmailByName_ === "function") {
        recipientEmail = getManagerEmailByName_(managerName);
      }
      recipientName = managerName || "Manager";
      stageTitle = "Manager Approval (Tahap 1)";

    } else if (status === "Waiting Final Closing Approval") {
      // Pending di Pak Bayun
      recipientEmail = "bayun.binantoro@kopikenangan.com"; // Email Asli Pak Bayun
      recipientName = "Pak Bayun";
      stageTitle = "Final Approval (Tahap 2)";
    }

    // 2. KIRIM EMAIL REMINDER
    if (recipientEmail && stageTitle) {
      var approveUrl = webAppUrl + "?action=closing_approval&id=" + encodeURIComponent(projectId) + "&decision=approve";
      var rejectUrl = webAppUrl + "?action=closing_approval&id=" + encodeURIComponent(projectId) + "&decision=reject";

      var subject = "[REMINDER APPROVAL] Pengajuan Project Closing - " + projectId + " (" + projectTitle + ")";

      var htmlBody = `
        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; max-width: 600px; margin: auto;">
          <h2 style="color: #d9534f; margin-top: 0;">⏰ Pengingat: Persetujuan Project Closing Tertunda</h2>
          <p>Halo <b>${recipientName}</b>,</p>
          <p>Sistem mendeteksi terdapat pengajuan <b>Project Closing</b> yang <b>belum diproses / membutuhkan persetujuan Anda</b> (${stageTitle}):</p>
          
          <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 14px; border: 1px solid #ddd;">
            <tr style="border-bottom: 1px solid #ddd; background-color: #f9f9f9;">
              <td style="padding: 8px; font-weight: bold; width: 35%;">Project ID</td>
              <td style="padding: 8px;">${projectId}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
              <td style="padding: 8px; font-weight: bold;">Nama Project</td>
              <td style="padding: 8px; color: #8B0000; font-weight: bold;">${projectTitle}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
              <td style="padding: 8px; font-weight: bold;">Inisiator</td>
              <td style="padding: 8px;">${initiatorName}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
              <td style="padding: 8px; font-weight: bold;">Evidence & Summary</td>
              <td style="padding: 8px;">${remarksText}</td>
            </tr>
          </table>

          <p style="color: #555;">Mohon untuk meninjau dan memberikan keputusan persetujuan agar proyek ini dapat segera berstatus Closed. Klik salah satu tombol di bawah ini:</p>
          
          <div style="margin-top: 25px; text-align: center;">
            <a href="${approveUrl}" style="background-color: #28a745; color: white; padding: 12px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; margin-right: 10px; display: inline-block;">APPROVE CLOSING</a>
            <a href="${rejectUrl}" style="background-color: #dc3545; color: white; padding: 12px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">REJECT / REVISE</a>
          </div>

          <hr style="border: none; border-top: 1px solid #eee; margin-top: 25px;">
          <p style="font-size: 11px; color: #888; text-align: center;">Email pengingat otomatis dikirimkan oleh CIMS Kopi Kenangan.</p>
        </div>
      `;

      var systemName = (typeof CONFIG !== "undefined" && CONFIG.SYSTEM_NAME) ? CONFIG.SYSTEM_NAME : "CIMS Kopi Kenangan";

      GmailApp.sendEmail(recipientEmail, subject, "Silakan buka email ini menggunakan HTML viewer.", {
        htmlBody: htmlBody,
        name: systemName + " - Reminder"
      });

      Logger.log("✅ Reminder dikirim ke " + recipientEmail + " untuk Project: " + projectId);
      reminderCount++;
    }
  }

  Logger.log("Selesai! Total email reminder terkirim: " + reminderCount);
}