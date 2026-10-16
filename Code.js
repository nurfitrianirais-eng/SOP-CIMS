/*
 * =================================================
 * CIMS v1.3 - Complete Backend Logic (FIXED & UPDATED)
 * QAQC Continuous Improvement Management System
 * Dynamic 2-Level Proposal & Closing Approval Routing
 * =================================================
 */

// ==========================================
// KONSTANTA & KONFIGURASI UTAMA
// ==========================================
const LOGO_URL = "https://lh3.googleusercontent.com/d/1fdq6E6_Z11XVNGzIE0-k9x4XzMQMG9ab";
const WEB_APP_URL = "https://script.google.com/a/macros/kopikenangan.com/s/AKfycbzzVDRlzp97MeKa5RjpUxFKm3P_nGn84_ZJjDaPv36NZtJSyUZyEJsE_VTnlcHk7omi/exec";

/**
 * MAIN TRIGGER
 * Trigger: From Spreadsheet -> On Form Submit
 */
function onFormSubmit(e) {
  try {
    const project = readFormData_(e);

    // 🟢 Mengutamakan project.employeeName agar nama pemohon terdeteksi presisi
    const initiatorName = project.employeeName || project.initiator || project.initiatorName || project.nama || project.name || "";
    const jobTitle = project.jobTitle || project.jabatan || "";
    const email = project.email || project.emailAddress || project.initiatorEmail || "";

    // 🛑 1. VALIDASI DARWIN BOX (MASTER_EMPLOYEE)
    const checkDarwin = validateDarwinBoxData_(initiatorName, jobTitle);

    if (!checkDarwin.isValid) {
      if (email) {
        sendDarwinBoxCorrectionEmail_(email, initiatorName || "Pemohon", jobTitle, checkDarwin);
      }
      Logger.log("⛔ SUBMISSION REJECTED (Tidak sesuai Darwin Box): " + (initiatorName || "Nama Tidak Terdeteksi"));
      return; // 🛑 STOP! Jangan simpan ke MASTER_PROJECT
    }

    // 🟢 2. JIKA VALID, JALANKAN PROSES NORMAL
    validateProject_(project);
    saveProject_(project);
    createApproval_(project);
    notifyManagers_(project);
    notifyInitiator_(project);

    // AUTO-REFRESH DASHBOARD
    if (typeof setupCIMSExecutiveDashboard === "function") {
      setupCIMSExecutiveDashboard();
    }

    Logger.log("SUCCESS SUBMISSION: " + (project.projectID || "OK"));
  } catch (err) {
    Logger.log(err);
    MailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      "CIMS ERROR - Form Submission",
      err.stack || err.toString()
    );
  }
}

/*******************
 READ FORM DATA & DYNAMIC ROUTING
*******************/

function readFormData_(e) {
  let row = {};

  // 🟢 SUNTIKAN 1: Membaca langsung dari trigger form agar 100% kebal baris kosong
  if (e && e.namedValues) {
    for (let key in e.namedValues) {
      row[String(key).trim()] = String(e.namedValues[key][0] || "");
    }
  } else {
    // Fallback jika dijalankan manual (Scan baris terakhir yang benar-benar ada isinya)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formSheet = ss.getSheetByName("Form Responses 2");
    const aVals = formSheet.getRange("A1:A").getValues();
    let lastRow = 1;
    for (let i = aVals.length - 1; i >= 0; i--) {
      if (String(aVals[i][0]).trim() !== "") {
        lastRow = i + 1;
        break;
      }
    }
    const headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
    const values = formSheet.getRange(lastRow, 1, 1, formSheet.getLastColumn()).getValues()[0];
    headers.forEach(function (header, index) {
      row[String(header).trim()] = values[index];
    });
  }

  const jobTitleInput = row["Job Title"] || "";
  const deptInput = row["Department"] || "QAQC";
  
  // LOGIKA ROUTING MANAGER
  const routing = getRoutingByJobTitle_(jobTitleInput, deptInput);

  const membersInput = 
    row["Project Members"] || 
    row["Project Member"] || 
    row["Anggota Tim"] || 
    row["Nama Anggota"] || 
    row["Project Members (if team)"] || 
    "";

  return {
    projectID: generateProjectID_(),
    submitDate: row["Timestamp"] || new Date(),
    email: row["Email Address"] || "",
    employeeName: row["Employee Name"] || row["Nama"] || "",
    employeeID: row["Employee ID"] || row["NIK"] || "",
    jobTitle: jobTitleInput,
    department: "QAQC",
    outlet: row["Outlet"] || "",
    manager: routing.managerName,
    managerEmail: routing.managerEmail,
    projectTitle: row["Project Title"] || "",
    problem: row["Problem Statement"] || "",
    evidence: (row["Supporting Evidence"] && String(row["Supporting Evidence"]).indexOf("http") === 0) ? '=HYPERLINK("' + row["Supporting Evidence"] + '", "🔗 Buka Link Drive")' : (row["Supporting Evidence"] || ""),
    idea: row["Improvement Idea"] || "",
    scope: row["Project Scope"] || "Individual",
    members: membersInput,
    impact: row["Possible Impact"] || "",
    timeline: row["Estimated Timeline"] || "",
    remarks: row["Additional Notes"] || "",
    priority: "Medium",
    status: routing.initialStatus
  };
}

/**
 * DETERMINASI ROUTING MANAGER & STATUS BERDASARKAN JOB TITLE
 * (100% Membaca dari Job Title Saja)
 */
function getRoutingByJobTitle_(jobTitle) {
  var title = String(jobTitle || "").trim().toLowerCase();

  // 1. Identifikasi Asisten (agar Asisten Manager tidak langsung lari ke Head)
  var isAsstManager = title.includes("assistant") || 
                      title.includes("asisten") || 
                      title.includes("assisten") || 
                      title.includes("asst");

  // 2. Direct ke Pak Bayun (QAQC Head) jika:
  //    - Level Global / Group / Head / Director / VP
  //    - Manager Murni (bukan asisten)
  var isDirectToHead = title.includes("global") || 
                       title.includes("group") || 
                       title.includes("head") || 
                       title.includes("director") || 
                       title.includes("vp") || 
                       (title.includes("manager") && !isAsstManager);

  if (isDirectToHead) {
    return {
      managerName: "Bayun Binantoro",
      managerEmail: "bayun.binantoro@kopikenangan.com",
      initialStatus: "Waiting QAQC Head Approval"
    };
  }

  // 3. TIM QA -> Bu Ellyta Widia Putri (contoh: QA Specialist, QA Staff, QA Officer)
  if (title.includes("qa") || title.includes("quality assurance")) {
    return {
      managerName: "Ellyta Widia Putri",
      managerEmail: "ellyta.putri@kopikenangan.com",
      initialStatus: "Waiting QA Manager Approval"
    };
  }

  // 4. TIM QC -> Pak Yanuar Alif Rohmansyah (contoh: QC Specialist, QC Staff, QC Officer)
  if (title.includes("qc") || title.includes("quality control")) {
    return {
      managerName: "Yanuar Alif Rohmansyah",
      managerEmail: "yanuar.arlief@kopikenangan.com",
      initialStatus: "Waiting QC Manager Approval"
    };
  }

  // 5. Default Fallback (jika tidak terdeteksi spesifik) -> Pak Bayun
  return {
    managerName: "Bayun Binantoro",
    managerEmail: "bayun.binantoro@kopikenangan.com",
    initialStatus: "Waiting QAQC Head Approval"
  };
}

function getManagerByRole_(roleName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("MASTER_MANAGER");
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const managerName = data[i][0]; 
    const managerEmail = data[i][1]; 
    const managerRole = String(data[i][2]).trim(); 
    const active = String(data[i][5]).toUpperCase().trim();

    if (managerRole.toUpperCase() === roleName.toUpperCase() && active === "YES") {
      return {
        name: managerName,
        email: managerEmail,
        role: managerRole
      };
    }
  }

  throw new Error("Approver dengan Role '" + roleName + "' tidak ditemukan di MASTER_MANAGER.");
}

function getManagerEmailByName_(managerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("MASTER_MANAGER");
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(managerName).trim().toLowerCase()) {
      return data[i][1];
    }
  }
  return "";
}

/*******************
 DATABASE & HELPER
*******************/

function generateProjectID_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MASTER_PROJECT");
  const year = new Date().getFullYear();
  
  // 🟢 SUNTIKAN 2: Scan seluruh ID di kolom A, abaikan baris kosong berapapun jumlahnya
  const ids = sheet.getRange("A2:A").getValues().flat();
  let maxRunning = 0;

  for (let i = 0; i < ids.length; i++) {
    let val = String(ids[i]).trim();
    if (val.indexOf("CI-") === 0) {
      let split = val.split("-");
      if (split.length === 3) {
        let num = Number(split[2]);
        if (num > maxRunning) {
          maxRunning = num;
        }
      }
    }
  }

  let running = maxRunning + 1;
  return "CI-" + year + "-" + Utilities.formatString("%04d", running);
}

function validateProject_(project) {
  const requiredFields = [
    ["Employee Name", project.employeeName],
    ["Employee ID", project.employeeID],
    ["Project Title", project.projectTitle],
    ["Problem Statement", project.problem],
    ["Improvement Idea", project.idea]
  ];

  requiredFields.forEach(function (item) {
    if (!item[1]) {
      throw new Error(item[0] + " is required.");
    }
  });
}

function checkDuplicateProject_(projectID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MASTER_PROJECT");
  if (sheet.getLastRow() <= 1) return;

  const ids = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getValues()
    .flat();

  if (ids.includes(projectID)) {
    throw new Error("Duplicate Project ID : " + projectID);
  }
}

function saveProject_(project) {
  checkDuplicateProject_(project.projectID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MASTER_PROJECT");
  const row = sheet.getLastRow() + 1;

  // 🟢 HITUNG DUE DATE OTOMATIS BERDASARKAN TIMELINE FORM
  const calculatedDueDate = calculateSmartDueDate_(project.submitDate, project.timeline);

  sheet.getRange(row, 1, 1, 22).setValues([[
    project.projectID,   // A
    project.submitDate,  // B
    project.projectTitle,// C
    project.employeeName,// D
    project.employeeID,  // E
    project.email,       // F
    project.scope,       // G
    project.members,     // H
    project.department,  // I
    project.jobTitle,    // J
    project.manager,     // K (Terkunci sebagai Atasan Direct)
    project.problem,     // L
    project.evidence,    // M
    project.idea,        // N
    project.impact,      // O
    project.timeline,    // P
    project.priority,    // Q
    project.status,      // R
    calculatedDueDate,   // S (Due Date Otomatis)
    "",                  // T
    "",                  // U
    project.remarks      // V
  ]]);
}

function createApproval_(project) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APPROVAL");
  const nextRow = sheet.getLastRow() + 1;

  sheet.getRange(nextRow, 1, 1, 6).setValues([[
    project.projectID,
    "Waiting",
    project.priority,
    "No",
    "",
    ""
  ]]);
}

/*************************************************
 * PROCESS CLOSING FROM WEB APP
 *************************************************/
function processClosingFromWeb(formData) {
  try {
    // 🛑 VALIDASI WAJIB: Cek jika Final Evidence atau Closing Summary Kosong
    if (!formData.finalEvidence || String(formData.finalEvidence).trim() === "") {
      return { success: false, message: "Gagal: Link Bukti Akhir (Final Evidence) wajib diisi!" };
    }
    if (!formData.closingSummary || String(formData.closingSummary).trim() === "") {
      return { success: false, message: "Gagal: Ringkasan Hasil (Closing Summary) wajib diisi!" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MASTER_PROJECT");
    const data = sheet.getDataRange().getValues();

    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(formData.projectId).trim()) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      return { success: false, message: "Project ID " + formData.projectId + " tidak ditemukan!" };
    }

    // Update Data Penutupan di MASTER_PROJECT
    sheet.getRange(foundRow, 18).setValue("Waiting Closing Approval");
    sheet.getRange(foundRow, 20).setValue(formData.completionDate);
    sheet.getRange(foundRow, 21).setValue((formData.benefitCategory || "Tidak Diklaim") + " - " + (formData.claimedSaving || "-"));
    
    const currentRemarks = sheet.getRange(foundRow, 22).getValue() || "";
    const closingNotes = "Closing Summary: " + formData.closingSummary + " | Evidence: " + formData.finalEvidence;
    sheet.getRange(foundRow, 22).setValue(currentRemarks ? currentRemarks + " || " + closingNotes : closingNotes);

    if (typeof sendClosingApprovalEmail_ === "function") {
      sendClosingApprovalEmail_(formData.projectId, formData);
    }

    return { 
      success: true, 
      message: "Pengajuan Closing untuk proyek " + formData.projectId + " berhasil dikirim dan menunggu verifikasi Manager." 
    };

  } catch (err) {
    return { success: false, message: "Gagal memproses closing: " + err.toString() };
  }
}

function processClosingApprovalAction(projectId, isApproved, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MASTER_PROJECT");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(projectId).trim()) {
      const rowIdx = i + 1;
      
      if (isApproved) {
        sheet.getRange(rowIdx, 18).setValue("Closed");
        
        if (typeof sendClosingNotification_ === "function") {
          sendClosingNotification_({
            projectId: projectId,
            completionDate: data[i][19],
            benefitCategory: "Proyek Tuntas",
            claimedSaving: data[i][20],
            finalEvidence: "Lihat di Master Project (Kolom V)"
          });
        }
        if (typeof sendClosingResultEmail_ === "function") {
          sendClosingResultEmail_(projectId, true, note);
        }
      } else {
        sheet.getRange(rowIdx, 18).setValue("Need Revision"); 
        if (typeof sendClosingResultEmail_ === "function") {
          sendClosingResultEmail_(projectId, false, note);
        }
      }
      return "Status penutupan proyek " + projectId + " berhasil diperbarui.";
    }
  }
  return "Project tidak ditemukan.";
}

/*******************
 APPROVAL AUTOMATION LEVEL 1 & 2
*******************/
function updateApproval_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const approvalSheet = ss.getSheetByName("APPROVAL");
  const projectSheet = ss.getSheetByName("MASTER_PROJECT");

  const approval = approvalSheet.getRange(row, 1, 1, 6).getValues()[0];
  const projectID = approval[0];
  const decision = approval[1]; 
  const priority = approval[2];
  const escalate = approval[3];
  const comments = approval[4];

  if (!projectID || decision === "Waiting") return;

  if (!approval[5]) {
    approvalSheet.getRange(row, 6).setValue(new Date());
  }

  const data = projectSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectID) {
      const currentStatus = data[i][17]; 
      const initiatorEmail = data[i][5];  
      const projectTitle = data[i][2];    
      const jobTitle = data[i][9];        
      const currentManagerName = String(data[i][10]).trim();

      projectSheet.getRange(i + 1, 17).setValue(priority);
      projectSheet.getRange(i + 1, 22).setValue(comments);

      if (decision === "Approved") {
        
        if (currentStatus === "Waiting QA Manager Approval" || currentStatus === "Waiting QC Manager Approval") {
          
          const qaqcHead = getManagerByRole_("QAQC Head");
          
          // Jika Atasan Direct dari awal memang sudah Pak Bayun (misal: Manager Murni/Global)
          if (qaqcHead.name.trim().toLowerCase() === currentManagerName.toLowerCase()) {
            
            projectSheet.getRange(i + 1, 18).setValue("Approved"); 

            const due = new Date();
            due.setDate(due.getDate() + 30);
            projectSheet.getRange(i + 1, 19).setValue(due);

            sendApprovalResult_(data[i], "Approved", comments, escalate);

          } else {
            // KOLOM K (MANAGER DIRECT) DIBIARKAN TERKUNCI PERMANEN
            // HANYA STATUS (KOLOM R) YANG DINAIKKAN KE QAQC HEAD APPROVAL
            projectSheet.getRange(i + 1, 18).setValue("Waiting QAQC Head Approval"); 

            approvalSheet.getRange(row, 2).setValue("Waiting");
            approvalSheet.getRange(row, 6).setValue(""); 

            notifyManagers_({
              projectID: projectID,
              projectTitle: projectTitle,
              employeeName: data[i][3], 
              jobTitle: jobTitle,
              scope: data[i][6],       
              members: data[i][7],     
              problem: data[i][11],    
              idea: data[i][13],       
              manager: qaqcHead.name,
              managerEmail: qaqcHead.email
            });

            sendProgressEmail_(initiatorEmail, projectID, projectTitle, "Lolos Review Manager (Menunggu Approval QAQC Head)", data[i][7]);
          }

        } else if (currentStatus === "Waiting QAQC Head Approval") {
          projectSheet.getRange(i + 1, 18).setValue("Approved"); 

          const due = new Date();
          due.setDate(due.getDate() + 30);
          projectSheet.getRange(i + 1, 19).setValue(due);

          sendApprovalResult_(data[i], "Approved", comments, escalate);
        }

      } else {
        const newStatus = (decision === "Rejected") ? "Rejected" : "Need Revision";
        projectSheet.getRange(i + 1, 18).setValue(newStatus); 
        sendApprovalResult_(data[i], newStatus, comments, escalate);
      }

      if (typeof setupCIMSExecutiveDashboard === "function") {
        setupCIMSExecutiveDashboard();
      }

      break;
    }
  }
}

/*******************
 EMAIL NOTIFICATION SYSTEM
*******************/

function notifyManagers_(project) {
  if (!project.managerEmail) return;

  const approvalUrl = WEB_APP_URL + "?id=" + project.projectID;
  const logoHtml = "<div style='text-align: center; margin-bottom: 20px;'><img src='" + LOGO_URL + "' alt='Kenangan Brands' style='max-width: 180px; height: auto;'></div>";

  MailApp.sendEmail({
    to: project.managerEmail,
    subject: "[CIMS APPROVAL NEEDED] " + project.projectID + " - " + project.projectTitle,
    htmlBody:
      "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 24px; border-radius: 8px; background-color: #ffffff;'>" +
      logoHtml +
      "<h2 style='color: #8B0000; text-align: center; margin-top: 0;'>Continuous Improvement Management System</h2>" +
      "<p>Halo <b>" + project.manager + "</b>,</p>" +
      "<p>Ada pengajuan proyek Continuous Improvement baru yang membutuhkan review dan approval Anda.</p>" +
      "<table border='1' cellpadding='8' style='border-collapse: collapse; width: 100%; border-color: #eee; font-size: 13px; margin-bottom: 20px;'>" +
      "<tr style='background-color: #f9f9f9;'><td width='35%'><b>Project ID</b></td><td><b style='color: #8B0000;'>" + project.projectID + "</b></td></tr>" +
      "<tr><td><b>Pengusul</b></td><td>" + project.employeeName + " (" + project.jobTitle + ")</td></tr>" +
      "<tr style='background-color: #f9f9f9;'><td><b>Judul Proyek</b></td><td>" + project.projectTitle + "</td></tr>" +
      "<tr><td><b>Project Scope</b></td><td>" + (project.scope || "Individual") + "</td></tr>" +
      "<tr style='background-color: #f9f9f9;'><td><b>Project Members</b></td><td>" + (project.members || "-") + "</td></tr>" +
      "<tr><td><b>Problem Statement</b></td><td>" + project.problem + "</td></tr>" +
      "<tr style='background-color: #f9f9f9;'><td><b>Ide Perbaikan</b></td><td>" + project.idea + "</td></tr>" +
      "</table>" +
      "<div style='text-align: center; margin: 25px 0;'>" +
      "<a href='" + approvalUrl + "' style='background-color: #8B0000; color: #ffffff; padding: 12px 26px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px; display: inline-block;'>Review & Approve Project</a>" +
      "</div>" +
      "</div>"
  });
}

/**
 * Mengirim email notifikasi ke Inisiator & CC ke Member Tim saat Form diajukan
 */
function notifyInitiator_(project) {
  const initiatorEmail = project.email || project.emailAddress || project.initiatorEmail || "";
  if (!initiatorEmail) return;

  const projectID = project.projectID || project.id || "-";
  const projectTitle = project.projectTitle || project.title || project.namaProyek || "-";
  const membersStr = project.members || project.member || project.anggota || "";

  // 1. Tarik email member tim dari tab MASTER_EMPLOYEE
  const memberEmails = (typeof getMemberEmails_ === "function") ? getMemberEmails_(membersStr) : "";

  const logoHtml = "<div style='text-align: center; margin-bottom: 20px;'><img src='" + LOGO_URL + "' alt='Kenangan Brands' style='max-width: 180px; height: auto;'></div>";

  const mailOptions = {
    to: initiatorEmail,
    subject: "[CIMS SUBMISSION] " + projectID + " - Pengajuan Proyek Berhasil",
    htmlBody:
      "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; background-color: #ffffff;'>" +
      logoHtml +
      "<h2 style='color: #8B0000; text-align: center; margin-top: 0;'>Continuous Improvement Management System</h2>" +
      "<p>Halo Team,</p>" +
      "<p>Proyek Continuous Improvement baru telah berhasil diajukan dan saat ini sedang menunggu proses peninjauan (<i>review</i>) Manager.</p>" +
      "<table border='1' cellpadding='8' style='border-collapse:collapse; width: 100%; border-color: #ddd; font-size: 13px; margin-bottom: 15px;'>" +
      "<tr style='background-color: #f9f9f9;'><td width='35%'><b>Project ID</b></td><td>" + projectID + "</td></tr>" +
      "<tr><td><b>Judul Proyek</b></td><td>" + projectTitle + "</td></tr>" +
      "<tr style='background-color: #f9f9f9;'><td><b>Status Saat Ini</b></td><td><b style='color: #D97706;'>Waiting Approval Manager</b></td></tr>" +
      "</table>" +
      "<p style='font-size: 12px; color: #777;'>Email ini dikirimkan otomatis kepada Inisiator dan ditembuskan (CC) ke seluruh Anggota Tim yang terdaftar.</p>" +
      "</div>"
  };

  // 2. Tambahkan CC jika terdapat email member tim
  if (memberEmails) {
    mailOptions.cc = memberEmails;
  }

  MailApp.sendEmail(mailOptions);
}

function sendProgressEmail_(email, projectID, projectTitle, statusText, membersStr) {
  if (!email) return;

  // Dapatkan daftar email member tim (jika ada)
  const memberEmails = (typeof getMemberEmails_ === "function") ? getMemberEmails_(membersStr) : "";

  const logoHtml = "<div style='text-align: center; margin-bottom: 20px;'><img src='" + LOGO_URL + "' alt='Kenangan Brands' style='max-width: 180px; height: auto;'></div>";

  const mailOptions = {
    to: email,
    subject: "[CIMS UPDATE] " + projectID + " - Approved by Manager",
    htmlBody:
      "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; background-color: #ffffff;'>" +
      logoHtml +
      "<h2 style='color: #8B0000; text-align: center; margin-top: 0;'>Continuous Improvement Management System</h2>" +
      "<p>Selamat! Project Anda telah disetujui di tingkat Manager dan dilanjutkan ke <b>QAQC Head</b>.</p>" +
      "<table border='1' cellpadding='8' style='border-collapse:collapse; width: 100%; border-color: #ddd; font-size: 13px;'>" +
      "<tr style='background-color: #f9f9f9;'><td width='35%'><b>Project ID</b></td><td>" + projectID + "</td></tr>" +
      "<tr><td><b>Judul Proyek</b></td><td>" + projectTitle + "</td></tr>" +
      "<tr style='background-color: #f9f9f9;'><td><b>Status Saat Ini</b></td><td><b style='color: #008000;'>" + statusText + "</b></td></tr>" +
      "</table>" +
      "</div>"
  };

  // Jika ada email member tim, masukkan ke CC
  if (memberEmails) {
    mailOptions.cc = memberEmails;
  }

  MailApp.sendEmail(mailOptions);
}

function sendApprovalResult_(projectRow, status, comments, escalate) {
  const email = projectRow[5]; // Email Inisiator (Kolom F)
  if (!email) return;

  const projectID = projectRow[0];
  const membersStr = projectRow[7]; // Members Tim (Kolom H)

  // 🟢 Tarik daftar email member tim untuk di-CC
  const memberEmails = (typeof getMemberEmails_ === "function") ? getMemberEmails_(membersStr) : "";

  const revisionUrl = WEB_APP_URL + "?action=revise&id=" + projectID;
  const closingUrl = WEB_APP_URL + "?action=close&id=" + projectID; 
  const logoHtml = "<div style='text-align: center; margin-bottom: 20px;'><img src='" + LOGO_URL + "' alt='Kenangan Brands' style='max-width: 180px; height: auto;'></div>";

  let bodyHtml = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; background-color: #ffffff;'>" +
    logoHtml +
    "<h2 style='color: #8B0000; text-align: center; margin-top: 0;'>Continuous Improvement Management System</h2>" +
    "<p>Proses review proyek Continuous Improvement Anda telah diperbarui.</p>" +
    "<table border='1' cellpadding='8' style='border-collapse:collapse; width: 100%; border-color: #ddd; font-size: 13px;'>" +
    "<tr style='background-color: #f9f9f9;'><td width='35%'><b>Project ID</b></td><td>" + projectID + "</td></tr>" +
    "<tr><td><b>Project Title</b></td><td>" + projectRow[2] + "</td></tr>" +
    "<tr style='background-color: #f9f9f9;'><td><b>Status Keputusan</b></td><td><b style='color: #8B0000;'>" + status + "</b></td></tr>" +
    "<tr><td><b>Catatan Approver</b></td><td>" + (comments || "-") + "</td></tr>" +
    "</table>";

  // 🔴 PENANGANAN REVISI
  if (status === "Need Revision") {
    bodyHtml += "<div style='text-align: center; margin: 25px 0;'>" +
      "<a href='" + revisionUrl + "' style='background-color: #d9534f; color: #ffffff; padding: 12px 26px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px; display: inline-block;'>✏️ Klik Disini Untuk Revisi Proyek</a>" +
      "</div>";
  } 
  // 🟢 PENANGANAN APPROVED (TOMBOL CLOSING DINI)
  else if (status === "Approved") {
    bodyHtml += "<div style='background-color: #E8F8F5; border-left: 4px solid #27AE60; padding: 12px; margin: 20px 0; border-radius: 4px;'>" +
      "<p style='margin: 0; color: #1E8449; font-weight: bold;'>🎉 Selamat! Proyek Anda Resmi Disetujui.</p>" +
      "<p style='margin: 5px 0 0 0; font-size: 12px; color: #27AE60;'>Simpan email ini. Jika proyek Anda telah selesai dilaksanakan (meskipun lebih cepat dari target), Anda dapat mengajukan penutupan proyek (*Closing*) kapan saja melalui tombol di bawah.</p>" +
      "</div>" +
      "<div style='text-align: center; margin: 25px 0;'>" +
      "<a href='" + closingUrl + "' style='background-color: #27AE60; color: #ffffff; padding: 12px 26px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px; display: inline-block;'>✅ Klik Disini Untuk Ajukan Closing Proyek</a>" +
      "</div>";
  }

  bodyHtml += "<br><p style='font-size: 12px; color: #777;'>Email ini dikirimkan otomatis kepada Inisiator dan ditembuskan (CC) ke seluruh Anggota Tim.</p>" +
    "<p>Terima kasih atas kontribusi Anda dalam budaya Continuous Improvement!</p></div>";

  const mailOptions = {
    to: email,
    subject: "[CIMS RESULT] " + projectID + " - " + status,
    htmlBody: bodyHtml
  };

  // 👥 Tambahkan CC ke seluruh Member Tim jika ada
  if (memberEmails) {
    mailOptions.cc = memberEmails;
  }

  MailApp.sendEmail(mailOptions);
}

/**
 * PENGINGAT APPROVAL (Membaca STATUS di Kolom R agar Kolom K Terkunci)
 */
function sendApprovalReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MASTER_PROJECT");
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var sentCount = 0;

  for (var i = 1; i < data.length; i++) {
    var projectId = String(data[i][0] || "").trim();
    var projectTitle = data[i][2];
    var initiatorName = data[i][3];
    var directManager = String(data[i][10] || "").trim(); // Kolom K (Atasan Direct)
    var status = String(data[i][17] || "").trim();       // Kolom R (Status Proyek)

    var targetEmail = "";
    var targetName = "";

    // 🟢 CEK PENERIMA BERDASARKAN STATUS APPROVAL (KOLOM R)
    if (status === "Waiting QA Manager Approval") {
      targetEmail = "ellyta.putri@kopikenangan.com";
      targetName = "Ellyta Widia Putri";
    } else if (status === "Waiting QC Manager Approval") {
      targetEmail = "yanuar.arlief@kopikenangan.com";
      targetName = "Yanuar Alif Rohmansyah";
    } else if (status === "Waiting QAQC Head Approval") {
      targetEmail = "bayun.binantoro@kopikenangan.com";
      targetName = "Bayun Binantoro";
    }

    if (targetEmail !== "") {
      var subject = "⚠️ [ACTION REQUIRED] Pengingat Approval CIMS - " + projectId;
      var approvalUrl = WEB_APP_URL + "?action=approve&id=" + projectId;

      var htmlBody = `
        <div style="font-family: Arial, sans-serif; max-width: 600px; border: 1px solid #E0E0E0; padding: 20px; border-radius: 8px;">
          <h2 style="color: #8B0000; margin-top: 0;">⚠️ PENGINGAT APPROVAL TERTUNDA</h2>
          <p>Halo <strong>${targetName}</strong>,</p>
          <p>Sistem mencatat ada pengajuan proyek Continuous Improvement yang <strong>belum Anda proses (Pending)</strong>.</p>
          
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
              <td style="padding: 10px; border: 1px solid #DDDDDD; font-weight: bold;">Initiator</td>
              <td style="padding: 10px; border: 1px solid #DDDDDD;">${initiatorName}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #DDDDDD; font-weight: bold;">Atasan Direct</td>
              <td style="padding: 10px; border: 1px solid #DDDDDD;">${directManager}</td>
            </tr>
          </table>

          <div style="text-align: center; margin: 25px 0;">
            <a href="${approvalUrl}" style="background-color: #8B0000; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">PROSES APPROVAL SEKARANG</a>
          </div>
        </div>
      `;

      GmailApp.sendEmail(targetEmail, subject, "Buka dengan HTML Viewer", {
        htmlBody: htmlBody,
        name: "CIMS Kopi Kenangan System"
      });
      
      sentCount++;
      Logger.log("✅ Reminder sent to: " + targetEmail + " for Project: " + projectId);
    }
  }
}

function onEdit(e) {
  // Safe Guard OnEdit
}



/**
 * Helper 1: Memeriksa Nama & Job Title ke MASTER_EMPLOYEE (SMART LOOKUP NAMA & NIK)
 */
function validateDarwinBoxData_(inputName, inputTitle) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName("MASTER_EMPLOYEE");
  
  if (!empSheet || empSheet.getLastRow() < 2) {
    return { isValid: true };
  }

  var cleanInputName = String(inputName || "").toLowerCase().trim();
  var cleanInputTitle = String(inputTitle || "").toLowerCase().trim();

  if (!cleanInputName) {
    return {
      isValid: false,
      type: "NAME_NOT_FOUND",
      message: "Nama pemohon tidak terdeteksi dari formulir."
    };
  }

  // Ambil Kolom A (Nama), Kolom B (Job Title), dan Kolom C (NIK jika ada)
  var empData = empSheet.getRange("A2:C" + empSheet.getLastRow()).getValues();
  var matchedName = "";
  var officialTitle = "";
  var isNameFound = false;

  for (var i = 0; i < empData.length; i++) {
    var empName = String(empData[i][0] || "").trim();
    var empTitle = String(empData[i][1] || "").trim();
    var empNik = String(empData[i][2] || "").trim(); // NIK di Kolom C (jika ada)
    var empNameLower = empName.toLowerCase();

    var isMatch = false;

    // 1. Cek Pencocokan Nama
    if (cleanInputName === empNameLower || 
       (cleanInputName.length >= 3 && (empNameLower.indexOf(cleanInputName) !== -1 || cleanInputName.indexOf(empNameLower) !== -1))) {
      isMatch = true;
    } 
    // 2. Cek Pencocokan jika user mengetik NIK di kolom Nama
    else if (empNik && cleanInputName === empNik.toLowerCase()) {
      isMatch = true;
    }

    if (isMatch) {
      isNameFound = true;
      matchedName = empName;
      officialTitle = empTitle || "Job Title Tidak Terdaftar";

      // Match Job Title (Toleran spasi & huruf besar/kecil)
      var cleanEmpTitle = empTitle.toLowerCase().replace(/\s+/g, '');
      var cleanUserTitle = cleanInputTitle.replace(/\s+/g, '');

      if (cleanUserTitle === cleanEmpTitle) {
        return { isValid: true };
      }
      break; // Karyawan ketemu, lanjut ke penanganan kesalahan Job Title
    }
  }

  if (!isNameFound) {
    return {
      isValid: false,
      type: "NAME_NOT_FOUND",
      message: "Nama / NIK <b>'" + inputName + "'</b> tidak terdaftar di basis data MASTER_EMPLOYEE (Darwin Box)."
    };
  }

  return {
    isValid: false,
    type: "TITLE_MISMATCH",
    inputTitle: inputTitle,
    officialName: matchedName,
    officialTitle: officialTitle,
    message: "Job Title <b>'" + inputTitle + "'</b> tidak sesuai dengan data resmi Darwin Box."
  };
}

/**
 * Helper 2: Mengirimkan Email Notifikasi Penolakan (SELALU MENAMPILKAN RINCIAN DATA RESMI)
 */
function sendDarwinBoxCorrectionEmail_(recipientEmail, inputName, inputTitle, errorInfo) {
  var subject = "⚠️ [Perhatian Required] Form CIMS Gagal Diproses - Ketidaksesuaian Data Darwin Box";
  var displayName = errorInfo.officialName || inputName || "Pemohon";

  var htmlBody = '<div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #E0E0E0; padding: 20px; border-radius: 8px;">' +
    '<h2 style="color: #D9534F; margin-top: 0;">⚠️ Pengajuan Proyek CIMS Gagal Diproses</h2>' +
    '<p>Halo <b>' + displayName + '</b>,</p>' +
    '<p>Pengajuan proyek Continuous Improvement (CIMS) Anda <b>belum dapat diproses/masuk ke dalam sistem</b> karena terdapat ketidaksesuaian data identitas dengan basis data resmi <b>Darwin Box (MASTER_EMPLOYEE)</b>.</p>' +
    
    '<div style="background-color: #FADBD8; border-left: 4px solid #D9534F; padding: 12px; margin: 15px 0;">' +
      '<p style="margin: 0; color: #78281F; font-weight: bold;">Detail Masalah:</p>' +
      '<p style="margin: 5px 0 0 0; color: #78281F;">' + errorInfo.message + '</p>' +
    '</div>';

  // 🟢 DIPAKSA SELALU MUNCUL JIKA DATA KARYAWAN DITEMUKAN
  if (errorInfo.officialName && errorInfo.officialTitle) {
    htmlBody += '<p><b>Rincian Data Resmi Anda di Darwin Box:</b></p>' +
      '<ul style="line-height: 1.6; margin-bottom: 20px;">' +
        '<li><b>Nama Resmi:</b> ' + errorInfo.officialName + '</li>' +
        '<li><b>Job Title Resmi:</b> <span style="color: #27AE60; font-weight: bold;">' + errorInfo.officialTitle + '</span> (Anda mengetik: <i>' + (inputTitle || "-") + '</i>)</li>' +
      '</ul>';
  }

  htmlBody += '<p><b>Langkah Selanjutnya:</b></p>' +
    '<ol style="line-height: 1.6;">' +
      '<li>Silakan buka kembali <b>Google Form CIMS</b>.</li>' +
      '<li>Isi kembali formulir dengan menginputkan <b>Nama Lengkap dan Job Title resmi yang sesuai dengan data Darwin Box di atas</b>.</li>' +
      '<li>Lakukan submit ulang agar alur persetujuan (*approval*) dapat berjalan secara otomatis ke Manager yang tepat.</li>' +
    '</ol>' +
    '<hr style="border: none; border-top: 1px solid #EEE; margin: 20px 0;">' +
    '<p style="font-size: 11px; color: #777;">Pesan ini dikirimkan secara otomatis oleh CIMS Automation System.</p>' +
    '</div>';

  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    htmlBody: htmlBody
  });
}
/**
 * FITUR ADMIN: Kirim Massal (Blast) Link Closing ke SEMUA Proyek "Approved"
 */
function blastClosingLinkToApprovedProjects() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MASTER_PROJECT");
  var ui = SpreadsheetApp.getUi();

  // Konfirmasi sebelum mengirim massal
  var confirm = ui.alert("🚀 Konfirmasi Blast Email",
    "Sistem akan mencari SEMUA proyek berstatus 'Approved' dan mengirimkan email Link Closing ke inisiator masing-masing.\n\nApakah Anda yakin ingin melanjutkan?",
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  var data = sheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    var projectId = String(data[i][0]).trim();
    var projectTitle = data[i][2];
    var email = data[i][5];
    var status = String(data[i][17]).trim(); // Status di Kolom R

    // Jika status Approved, kirim email
    if (status === "Approved" && email && email !== "-" && projectId && projectId !== "-") {
      var closingUrl = WEB_APP_URL + "?action=close&id=" + projectId;
      var logoHtml = "<div style='text-align: center; margin-bottom: 20px;'><img src='" + LOGO_URL + "' alt='Kenangan Brands' style='max-width: 180px; height: auto;'></div>";

      var subject = "[CIMS UPDATE] Link Pengajuan Closing Proyek Anda - " + projectId;
      
      var bodyHtml = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; background-color: #ffffff;'>" +
        logoHtml +
        "<h2 style='color: #27AE60; text-align: center; margin-top: 0;'>Pembaruan Sistem CIMS</h2>" +
        "<p>Halo Inisiator,</p>" +
        "<p>Sistem CIMS baru saja diperbarui. Untuk memudahkan Anda, kami melampirkan <b>Tautan Pengajuan Closing</b> khusus untuk proyek Anda yang saat ini sedang aktif berjalan:</p>" +
        "<p style='font-size: 15px;'><b>" + projectId + " - " + projectTitle + "</b></p>" +
        "<div style='background-color: #E8F8F5; border-left: 4px solid #27AE60; padding: 12px; margin: 20px 0; border-radius: 4px;'>" +
          "<p style='margin: 0; color: #1E8449; font-weight: bold;'>Kapan tombol di bawah ini digunakan?</p>" +
          "<p style='margin: 5px 0 0 0; font-size: 13px; color: #27AE60;'>Simpan email ini. Jika proyek Anda telah selesai dieksekusi 100% <b>meskipun selesai sebelum timeline yang diajukan (selesai lebih awal)</b>, Anda tidak perlu menunggu email pengingat dari sistem. Silakan langsung gunakan tombol di bawah ini untuk mengajukan Closing dan melaporkan hasilnya (*Benefit/Saving*).</p>" +
        "</div>" +
        "<div style='text-align: center; margin: 25px 0;'>" +
        "<a href='" + closingUrl + "' style='background-color: #27AE60; color: #ffffff; padding: 12px 26px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px; display: inline-block;'>✅ Ajukan Closing Proyek Ini</a>" +
        "</div>" +
        "</div>";

      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: bodyHtml
        });
        count++; // Hitung jumlah email sukses
      } catch (e) {
        Logger.log("Gagal kirim ke " + email + ": " + e.message);
      }
    }
  }

  // Notifikasi selesai
  ui.alert("✅ Blast Selesai!", "Sistem telah berhasil mengirimkan Tautan Closing ke " + count + " proyek yang berstatus Approved.", ui.ButtonSet.OK);
}
/**
 * FITUR ADMIN: Mengirimkan Link Closing secara manual untuk 1 Proyek yang dipilih
 * (Format Tampilan Email Persis Seperti Template Blast Email)
 */
function sendManualClosingLink() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== "MASTER_PROJECT") {
    ui.alert("⚠️ PERHATIAN", "Silakan buka tab MASTER_PROJECT dan klik salah satu sel pada baris proyek yang ingin dikirimkan link.", ui.ButtonSet.OK);
    return;
  }

  var row = sheet.getActiveCell().getRow();
  if (row < 2) {
    ui.alert("⚠️ PERHATIAN", "Silakan klik baris data proyek yang valid (baris 2 ke bawah).", ui.ButtonSet.OK);
    return;
  }

  var data = sheet.getRange(row, 1, 1, 18).getValues()[0];
  var projectId = String(data[0] || "").trim();
  var projectTitle = data[2] || "-";
  var email = data[5] || "";

  if (!projectId || projectId === "-") {
    ui.alert("⚠️ ERROR", "Baris yang Anda pilih tidak memiliki Project ID yang valid.", ui.ButtonSet.OK);
    return;
  }

  if (!email || email === "-") {
    ui.alert("⚠️ ERROR", "Email inisiator pada baris ini tidak ditemukan.", ui.ButtonSet.OK);
    return;
  }

  // Konfirmasi sebelum mengirim
  var confirm = ui.alert("Kirim Link Closing?", 
    "Kirim link closing ke: " + email + "\nUntuk Proyek: " + projectId + " - " + projectTitle, 
    ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;

  // Konstruksi Email (Persis seperti Template Blast)
  var closingUrl = WEB_APP_URL + "?action=close&id=" + projectId;
  var logoHtml = "<div style='text-align: center; margin-bottom: 20px;'><img src='" + LOGO_URL + "' alt='Kenangan Brands' style='max-width: 180px; height: auto;'></div>";
  var subject = "[CIMS UPDATE] Link Pengajuan Closing Proyek Anda - " + projectId;
  
  var bodyHtml = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; background-color: #ffffff;'>" +
    logoHtml +
    "<h2 style='color: #27AE60; text-align: center; margin-top: 0;'>Pembaruan Sistem CIMS</h2>" +
    "<p>Halo Inisiator,</p>" +
    "<p>Sistem CIMS baru saja diperbarui. Untuk memudahkan Anda, kami melampirkan <b>Tautan Pengajuan Closing</b> khusus untuk proyek Anda yang saat ini sedang aktif berjalan:</p>" +
    "<p style='font-size: 15px;'><b>" + projectId + " - " + projectTitle + "</b></p>" +
    "<div style='background-color: #E8F8F5; border-left: 4px solid #27AE60; padding: 12px; margin: 20px 0; border-radius: 4px;'>" +
      "<p style='margin: 0; color: #1E8449; font-weight: bold;'>Kapan tombol di bawah ini digunakan?</p>" +
      "<p style='margin: 5px 0 0 0; font-size: 13px; color: #27AE60;'>Simpan email ini. Jika proyek Anda telah selesai dieksekusi 100% <b>meskipun selesai sebelum timeline yang diajukan (selesai lebih awal)</b>, Anda tidak perlu menunggu email pengingat dari sistem. Silakan langsung gunakan tombol di bawah ini untuk mengajukan Closing dan melaporkan hasilnya (*Benefit/Saving*).</p>" +
    "</div>" +
    "<div style='text-align: center; margin: 25px 0;'>" +
    "<a href='" + closingUrl + "' style='background-color: #27AE60; color: #ffffff; padding: 12px 26px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px; display: inline-block;'>✅ Ajukan Closing Proyek Ini</a>" +
    "</div>" +
    "</div>";

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: bodyHtml
    });
    ui.alert("✅ BERHASIL", "Link closing telah sukses dikirim ke: " + email, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("❌ GAGAL", "Gagal mengirim email: " + e.message, ui.ButtonSet.OK);
  }
}
/**
 * HELPER: Mengambil email member tim dari tab MASTER_EMPLOYEE untuk di-CC
 */
function getMemberEmails_(membersStr) {
  if (!membersStr || membersStr === "-") return "";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName("MASTER_EMPLOYEE");
  if (!empSheet || empSheet.getLastRow() < 2) return "";

  const memberNames = (typeof parseIndividualNames_ === "function") ? parseIndividualNames_(membersStr) : membersStr.split(",");
  if (memberNames.length === 0) return "";

  // Kolom A = Nama, Kolom C = Email
  const empData = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 3).getValues();
  const emailList = [];

  for (let i = 0; i < memberNames.length; i++) {
    const rawName = String(memberNames[i]).trim().toLowerCase();
    if (!rawName || rawName === "-") continue;

    for (let j = 0; j < empData.length; j++) {
      const empName = String(empData[j][0]).trim().toLowerCase();
      const empEmail = String(empData[j][2]).trim(); 

      if (empName && (rawName === empName || rawName.indexOf(empName) !== -1 || empName.indexOf(rawName) !== -1)) {
        if (empEmail && empEmail.includes("@") && !emailList.includes(empEmail)) {
          emailList.push(empEmail);
        }
      }
    }
  }

  return emailList.join(",");
}
/**
 * FITUR ADMIN: Mengirim Email Informasi Proyek Manual Khusus Member Tim (1 Project)
 */
function sendManualInfoToMembers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== "MASTER_PROJECT") {
    ui.alert("⚠️ PERHATIAN", "Silakan buka tab MASTER_PROJECT dan klik salah satu sel pada baris proyek yang ingin dikirimkan email.", ui.ButtonSet.OK);
    return;
  }

  var row = sheet.getActiveCell().getRow();
  if (row < 2) {
    ui.alert("⚠️ PERHATIAN", "Pilih baris data proyek yang valid (baris 2 ke atas).", ui.ButtonSet.OK);
    return;
  }

  var data = sheet.getRange(row, 1, 1, 18).getValues()[0];
  var projectId = String(data[0] || "").trim();
  var projectTitle = data[2] || "-";
  var membersStr = data[7] || "";
  var statusText = data[17] || "Approved";

  // Tarik email member dari tab MASTER_EMPLOYEE
  var memberEmails = (typeof getMemberEmails_ === "function") ? getMemberEmails_(membersStr) : "";

  if (!memberEmails) {
    ui.alert("⚠️ INFO", "Proyek ini tidak memiliki Member, ATAU email member tidak ditemukan di tab MASTER_EMPLOYEE.\n\nNama di kolom Member: " + membersStr, ui.ButtonSet.OK);
    return;
  }

  // Konfirmasi pengiriman
  var confirm = ui.alert("Kirim Informasi Proyek ke Member?", 
    "Email akan dikirimkan langsung kepada Anggota Tim berikut:\n" + memberEmails, 
    ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;

  // Template Email Resmi (Produksi)
  var logoHtml = "<div style='text-align: center; margin-bottom: 20px;'><img src='" + LOGO_URL + "' alt='Kenangan Brands' style='max-width: 180px; height: auto;'></div>";
  var subject = "[CIMS INFO] Informasi Proyek Tim - " + projectId;
  
  var bodyHtml = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; background-color: #ffffff;'>" +
    logoHtml +
    "<h2 style='color: #8B0000; text-align: center; margin-top: 0;'>Continuous Improvement Management System</h2>" +
    "<p>Halo Team,</p>" +
    "<p>Berikut adalah pembaruan informasi proyek di mana Anda terdaftar sebagai <b>Member Tim</b>:</p>" +
    "<table border='1' cellpadding='8' style='border-collapse:collapse; width: 100%; border-color: #ddd; font-size: 13px; margin-bottom: 15px;'>" +
      "<tr style='background-color: #f9f9f9;'><td width='35%'><b>Project ID</b></td><td>" + projectId + "</td></tr>" +
      "<tr><td><b>Judul Proyek</b></td><td>" + projectTitle + "</td></tr>" +
      "<tr style='background-color: #f9f9f9;'><td><b>Status Saat Ini</b></td><td><b style='color: #27AE60;'>" + statusText + "</b></td></tr>" +
    "</table>" +
    "<p style='font-size: 12px; color: #777;'>Email ini dikirimkan otomatis sebagai pembaruan status proyek kepada Anggota Tim terkait.</p>" +
    "<p>Terima kasih atas kontribusi Anda dalam budaya Continuous Improvement!</p>" +
    "</div>";

  try {
    MailApp.sendEmail({
      to: memberEmails, 
      subject: subject,
      htmlBody: bodyHtml
    });
    ui.alert("✅ BERHASIL", "Email informasi proyek berhasil terkirim ke member tim!", ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("❌ GAGAL", "Gagal mengirim email: " + e.message, ui.ButtonSet.OK);
  }
}
/**
 * HELPER: Mencari data proyek berdasarkan Project ID di tab MASTER_PROJECT
 */
function getProjectById_(projectId) {
  if (!projectId) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MASTER_PROJECT");
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var searchId = String(projectId).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var currentId = String(data[i][0] || "").trim().toLowerCase(); // Kolom A: Project ID
    if (currentId === searchId) {
      return {
        rowIndex: i + 1, // Baris asli di spreadsheet (1-based index)
        row: data[i]     // Array data dalam 1 baris proyek
      };
    }
  }

  return null;
}
/**
 * HELPER: Menghitung Due Date (Batas Waktu) otomatis dari opsi Timeline form
 */
function calculateSmartDueDate_(submitDate, timelineStr) {
  if (!submitDate) return "";
  
  var date = new Date(submitDate);
  if (isNaN(date.getTime())) return "";

  var text = String(timelineStr || "").toLowerCase();
  var monthsToAdd = 1; // Default jika tidak terdeteksi (1 Bulan)

  // Logika pembacaan opsi Timeline
  if (text.indexOf("1 - 6") !== -1 || text.indexOf("6 bul") !== -1) {
    monthsToAdd = 6;
  } else if (text.indexOf("6 - 12") !== -1 || text.indexOf("12 bul") !== -1) {
    monthsToAdd = 12;
  } else if (text.indexOf("tahun") !== -1) {
    monthsToAdd = 12;
  } else if (text.indexOf("3 bul") !== -1) {
    monthsToAdd = 3;
  } else if (text.indexOf("< 1") !== -1) {
    monthsToAdd = 1;
  }

  // Tambahkan bulan maksimal ke tanggal submit
  date.setMonth(date.getMonth() + monthsToAdd);
  return date;
}
/**
 * 🛠️ UTILITY UTAMA: Sinkronisasi Massal Due Date Data Lama
 * Menghitung ulang Kolom S (Due Date) berdasarkan Kolom P (Timeline)
 */
function syncAllOldDueDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MASTER_PROJECT");
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var updateCount = 0;

  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][17] || "").trim(); // Kolom R: Status
    
    // Abaikan proyek yang sudah selesai atau ditolak
    if (status === "Closed" || status === "Rejected" || status === "") continue;

    var submitDate = data[i][1];   // Kolom B: Submit Date
    var timelineStr = data[i][15]; // Kolom P: Timeline

    if (submitDate && timelineStr) {
      // Panggil fungsi logika pintar yang sudah kita buat sebelumnya
      var newDueDate = calculateSmartDueDate_(submitDate, timelineStr);
      
      if (newDueDate) {
        // Timpa data lama di Kolom S (Kolom ke-19)
        sheet.getRange(i + 1, 19).setValue(newDueDate); 
        updateCount++;
      }
    }
  }
  
  Logger.log("✅ Berhasil sinkronisasi " + updateCount + " proyek.");
  try {
    SpreadsheetApp.getUi().alert("✅ Sinkronisasi Selesai! Berhasil memperbarui Due Date untuk " + updateCount + " proyek aktif.");
  } catch (err) {
    // Mengabaikan error UI jika dijalankan langsung dari Apps Script Editor
  }
}
/**
 * UTILITY: Tembak Ulang Email Closing Approval yang Sempat Tertunda
 */
function triggerPendingClosingEmails() {
  // 👇 INI YANG DIUBAH: Hanya masukkan ID proyek milik Ellyta
  var pendingIds = ["CI-2026-0017"]; 
  
  for (var i = 0; i < pendingIds.length; i++) {
    var pId = pendingIds[i];
    var pData = getProjectDetailsByID_(pId);
    
    if (pData && typeof sendClosingApprovalEmail_ === "function") {
      var formData = {
        projectId: pData.id,
        finalEvidence: pData.remarks,
        closingSummary: pData.remarks,
        benefitCategory: pData.impact,
        claimedSaving: pData.impact,
        completionDate: pData.submitDate
      };
      sendClosingApprovalEmail_(pId, formData);
      Logger.log("✅ Email closing berhasil dikirim ke Manager untuk Project: " + pId);
    }
  }
}
