/**
 * =================================================================
 * CIMS WEB APP & APPROVAL HANDLER
 * Render Halaman Web App (Approval, Revisi, atau Closing)
 * =================================================================
 */
function doGet(e) {
  try {
    var projectId = (e && e.parameter && e.parameter.id) ? e.parameter.id.trim() : "";
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action.trim() : "approval";

    // Pilihan Tampilan: approval, revise, close, atau closing_approval (Untuk Manager)
    var fileName = "Index";
    if (action === "revise") fileName = "Revision";
    if (action === "close") fileName = "Closing";
    
    // Auto-fallback nama file jika Index dinamai EmailTemplate
    if (fileName === "Index") {
      try { HtmlService.createTemplateFromFile("Index"); } catch (err) { fileName = "EmailTemplate"; }
    }

    // Jika Manager sedang melakukan approve form closing, tampilkan form simpel via pesan
    if (action === "closing_approval") {
        return handleClosingApprovalRequest_(projectId, e.parameter.decision, e.parameter.note);
    }

    var template = HtmlService.createTemplateFromFile(fileName);
    var projectData = getProjectDetailsByID_(projectId);

    if (projectData) {
      template.project = projectData;
      var htmlOutput = template.evaluate();
      htmlOutput.setTitle(action === "close" ? "CIMS - Project Closing" : (action === "revise" ? "CIMS - Revisi Project" : "CIMS Approval - Kopi Kenangan"));
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return htmlOutput;
    } 

    return HtmlService.createHtmlOutput("<div style='padding:20px; text-align:center;'><h2 style='color:#DC3545;'>Project Not Found</h2><p>Project ID " + projectId + " tidak ditemukan di sistem.</p></div>");

  } catch (error) {
    return HtmlService.createHtmlOutput("<div style='padding:20px; color:red;'>Error: " + error.toString() + "</div>");
  }
}

function cleanStr_(str) {
  return String(str || "").replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
}

/**
 * Mengambil Detail Proyek (Termasuk Remarks/Catatan Revisi)
 */
function getProjectDetailsByID_(projectId) {
  if (!projectId) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MASTER_PROJECT");
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var searchClean = cleanStr_(projectId);

  for (var i = 1; i < data.length; i++) {
    if (cleanStr_(data[i][0]) === searchClean && searchClean !== "") {
      return {
        id: data[i][0],          // A: Project ID
        submitDate: data[i][1],  // B: Submit Date
        title: data[i][2],       // C: Project Title
        initiator: data[i][3],   // D: Initiator
        scope: data[i][6],       // G: Scope
        members: data[i][7],     // H: Members
        department: data[i][8],  // I: Department
        jobTitle: data[i][9],    // J: Job Title
        manager: data[i][10],    // K: Manager
        problem: data[i][11],    // L: Problem
        evidence: data[i][12],   // M: Evidence
        idea: data[i][13],       // N: Idea
        impact: data[i][14],     // O: Impact
        timeline: data[i][15],   // P: Timeline
        priority: data[i][16],   // Q: Priority
        status: data[i][17],     // R: Status
        remarks: data[i][21]     // V: Remarks / Revision Note
      };
    }
  }
  return null;
}

/**
 * Helper: Mencari Email Manager Berdasarkan Nama dari tab MASTER_EMPLOYEE
 */
function getManagerEmailByName_(managerName) {
  if (!managerName) return "";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName("MASTER_EMPLOYEE");
  if (!empSheet) return "";

  var data = empSheet.getDataRange().getValues();
  var searchName = String(managerName).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var empName = String(data[i][0] || "").trim().toLowerCase(); // Kolom A: Nama Karyawan
    var empEmail = String(data[i][4] || "").trim();             // Kolom E: Email

    if (empName && (empName === searchName || searchName.indexOf(empName) !== -1 || empName.indexOf(searchName) !== -1)) {
      if (empEmail && empEmail.indexOf("@") !== -1) {
        return empEmail;
      }
    }
  }
  return "";
}

/**
 * Memproses Keputusan Approval (Proposal Awal)
 */
function processApprovalFromWeb(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var approvalSheet = ss.getSheetByName("APPROVAL");
    if (!approvalSheet) return { success: false, message: "Sheet APPROVAL tidak ditemukan." };

    var data = approvalSheet.getDataRange().getValues();
    var rowIndex = -1;
    var searchClean = cleanStr_(formData.projectId);

    for (var i = 1; i < data.length; i++) {
      if (cleanStr_(data[i][0]) === searchClean) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "ID tidak ditemukan di sheet APPROVAL." };

    approvalSheet.getRange(rowIndex, 2).setValue(formData.decision);
    approvalSheet.getRange(rowIndex, 5).setValue(formData.comment);
    approvalSheet.getRange(rowIndex, 6).setValue(new Date());

    if (typeof updateApproval_ === "function") {
      updateApproval_(rowIndex);
    }

    return { success: true, message: "Keputusan approval berhasil diproses!" };

  } catch (err) {
    return { success: false, message: "Error: " + err.toString() };
  }
}

/**
 * Memproses Pengiriman Revisi dari Initiator (Versi Aman)
 */
function processRevisionFromWeb(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var projectSheet = ss.getSheetByName("MASTER_PROJECT");
    var approvalSheet = ss.getSheetByName("APPROVAL");

    var data = projectSheet.getDataRange().getValues();
    var rowIndex = -1;
    var searchClean = cleanStr_(formData.projectId);

    for (var i = 1; i < data.length; i++) {
      if (cleanStr_(data[i][0]) === searchClean) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "Project ID tidak ditemukan di MASTER_PROJECT." };

    // 1. Update Data di Sheet MASTER_PROJECT
    projectSheet.getRange(rowIndex, 3).setValue(formData.title);     // C: Title
    projectSheet.getRange(rowIndex, 7).setValue(formData.scope);     // G: Scope
    projectSheet.getRange(rowIndex, 8).setValue(formData.members);   // H: Members
    projectSheet.getRange(rowIndex, 12).setValue(formData.problem);  // L: Problem
    projectSheet.getRange(rowIndex, 13).setValue(formData.evidence); // M: Evidence
    projectSheet.getRange(rowIndex, 14).setValue(formData.idea);     // N: Idea
    projectSheet.getRange(rowIndex, 15).setValue(formData.impact);   // O: Impact
    projectSheet.getRange(rowIndex, 16).setValue(formData.timeline); // P: Timeline

    // 2. Tentukan Status Approval Kembali
    var jobTitle = String(projectSheet.getRange(rowIndex, 10).getValue() || "").toUpperCase();
    var newStatus = (jobTitle.indexOf("QC") !== -1) ? "Waiting QC Manager Approval" : "Waiting QA Manager Approval";
    projectSheet.getRange(rowIndex, 18).setValue(newStatus); // R: Status

    // 3. Reset Status di Sheet APPROVAL
    if (approvalSheet) {
      var appData = approvalSheet.getDataRange().getValues();
      for (var j = 1; j < appData.length; j++) {
        if (cleanStr_(appData[j][0]) === searchClean) {
          approvalSheet.getRange(j + 1, 2).setValue("Waiting");
          approvalSheet.getRange(j + 1, 6).setValue("");
          break;
        }
      }
    }

    // 4. Kirim Email Notifikasi Kembali ke Manager (Secara Aman)
    var managerName = projectSheet.getRange(rowIndex, 11).getValue();
    var managerEmail = getManagerEmailByName_(managerName);

    if (typeof notifyManagers_ === "function") {
      notifyManagers_({
        projectID: formData.projectId,
        projectTitle: formData.title,
        employeeName: projectSheet.getRange(rowIndex, 4).getValue(),
        jobTitle: projectSheet.getRange(rowIndex, 10).getValue(),
        scope: formData.scope,
        members: formData.members,
        problem: formData.problem,
        idea: formData.idea,
        manager: managerName,
        managerEmail: managerEmail
      });
    }

    return { success: true, message: "Revisi berhasil disimpan! Proyek telah dikirim kembali ke Manager untuk di-review." };

  } catch (err) {
    return { success: false, message: "Error: " + err.toString() };
  }
}

/**
 * =========================================================
 * NEW ALUR CLOSING: PENGIRIMAN DARI INITIATOR
 * Mengubah Status Menjadi "Waiting Closing Approval" (Bukan Langsung Closed)
 * =========================================================
 */
function processClosingFromWeb(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var projectSheet = ss.getSheetByName("MASTER_PROJECT");
    if (!projectSheet) return { success: false, message: "Sheet MASTER_PROJECT tidak ditemukan." };

    var data = projectSheet.getDataRange().getValues();
    var rowIndex = -1;
    var searchClean = cleanStr_(formData.projectId);

    for (var i = 1; i < data.length; i++) {
      if (cleanStr_(data[i][0]) === searchClean) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "Project ID tidak ditemukan." };

    // Format Data untuk Kolom Remarks (V)
    var remarksText = "Evidence: " + formData.finalEvidence + " | Summary: " + formData.closingSummary;
    var benefitText = (formData.benefitCategory || "Uncategorized") + " - " + (formData.claimedSaving || "0");

    // Update Data Sementara ke Sheet MASTER_PROJECT (Status = Waiting Closing Approval)
    projectSheet.getRange(rowIndex, 18).setValue("Waiting Closing Approval"); // Kolom R: Status
    projectSheet.getRange(rowIndex, 20).setValue(formData.completionDate);    // Kolom T: Completion Date
    projectSheet.getRange(rowIndex, 21).setValue(benefitText);                // Kolom U: Benefit
    projectSheet.getRange(rowIndex, 22).setValue(remarksText);                // Kolom V: Remarks

    // Kirim Notifikasi Verifikasi ke Manager (berada di file Mail.gs)
    if (typeof sendClosingApprovalEmail_ === "function") {
      sendClosingApprovalEmail_(formData.projectId, formData);
    }

    return { success: true, message: "Pengajuan Closing untuk proyek " + formData.projectId + " berhasil dikirim! Menunggu persetujuan Manager." };

  } catch (err) {
    return { success: false, message: "Error: " + err.toString() };
  }
}

/**
 * =========================================================
 * NEW ALUR CLOSING: SMART 1 TAHAP & 2 TAHAP APPROVAL
 * =========================================================
 */
function handleClosingApprovalRequest_(projectId, decision, note) {
    try {
    Logger.log("=== handleClosingApprovalRequest_ called ===");
    Logger.log("Project ID from URL: " + projectId);
    Logger.log("Decision: " + decision);
    Logger.log("Note: " + note);

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        Logger.log("Spreadsheet name: " + ss.getName());
        var projectSheet = ss.getSheetByName("MASTER_PROJECT");
        Logger.log("MASTER_PROJECT sheet found: " + (projectSheet ? "YES" : "NO"));
        var data = projectSheet.getDataRange().getValues();
        Logger.log("Total rows in MASTER_PROJECT: " + data.length);
        var rowIndex = -1;
        var searchClean = cleanStr_(projectId);
        Logger.log("Searching for (cleaned): " + searchClean);

        // 1. CARI BARIS DATA
        for (var i = 1; i < data.length; i++) {
            Logger.log("Row " + i + " project ID: " + data[i][0] + " | Cleaned: " + cleanStr_(data[i][0]));
            if (cleanStr_(data[i][0]) === searchClean) {
                rowIndex = i + 1;
                break;
            }
        }

        if (rowIndex === -1) {
            return HtmlService.createHtmlOutput("<div style='text-align:center; padding:50px;'><h2 style='color:#DC3545;'>Error!</h2><p>Project ID " + projectId + " tidak ditemukan!</p></div>");
        }

        var currentStatus = String(projectSheet.getRange(rowIndex, 18).getValue()).trim();
var managerName = String(data[rowIndex - 1][10] || "").toLowerCase();
        
        // Cek apakah Manager proyek ini adalah Pak Bayun
        var managerEmail = (typeof getInitiatorManagerEmail_ === "function") ? getInitiatorManagerEmail_(projectId) : "";
        var isPakBayunManager = (managerName.indexOf("bayun") !== -1 || managerEmail.toLowerCase().indexOf("bayun") !== -1);

        // 2. EKSEKUSI KEPUTUSAN APPROVE
        if (decision === "approve") {
            
            if (currentStatus === "Waiting Closing Approval") {
                
                // JIKA MANAGER ADALAH PAK BAYUN -> CUKUP 1 TAHAP (LANGSUNG CLOSED)
                if (isPakBayunManager) {
                    projectSheet.getRange(rowIndex, 18).setValue("Closed");
                    
                    if (typeof setupCIMSExecutiveDashboard === "function") {
                        setupCIMSExecutiveDashboard();
                    }

                    if (typeof sendClosingResultEmail_ === "function") {
                        sendClosingResultEmail_(projectId, true, note);
                    }
                    
                    return HtmlService.createHtmlOutput("<div style='text-align:center; padding:50px; font-family:sans-serif;'><h2 style='color:#28A745;'>Project Resmi Closed!</h2><p>Proyek " + projectId + " telah Anda setujui dan resmi berstatus Closed.</p></div>");
                
                } else {
                    // JIKA MANAGER BUKAN PAK BAYUN -> LANJUT KE TAHAP 2 (PAK BAYUN)
                    projectSheet.getRange(rowIndex, 18).setValue("Waiting Final Closing Approval");
                    
                    if (typeof sendFinalClosingApprovalEmail_ === "function") {
                        sendFinalClosingApprovalEmail_(projectId);
                    }
                    
                    return HtmlService.createHtmlOutput("<div style='text-align:center; padding:50px; font-family:sans-serif;'><h2 style='color:#28A745;'>Tahap 1 Disetujui!</h2><p>Proyek " + projectId + " telah Anda setujui dan otomatis diteruskan ke Pak Bayun untuk Final Approval.</p></div>");
                }
            
            } else if (currentStatus === "Waiting Final Closing Approval") {
                // TAHAP 2: Pak Bayun Approve -> RESMI CLOSED
                projectSheet.getRange(rowIndex, 18).setValue("Closed"); 
                
                if (typeof setupCIMSExecutiveDashboard === "function") {
                    setupCIMSExecutiveDashboard();
                }

                if (typeof sendClosingResultEmail_ === "function") {
                    sendClosingResultEmail_(projectId, true, note);
                }
                
                return HtmlService.createHtmlOutput("<div style='text-align:center; padding:50px; font-family:sans-serif;'><h2 style='color:#28A745;'>Final Approval Sukses!</h2><p>Proyek " + projectId + " telah resmi Closed. Dashboard telah diperbarui.</p></div>");
            
            } else if (currentStatus === "Closed") {
                return HtmlService.createHtmlOutput("<div style='text-align:center; padding:50px; font-family:sans-serif;'><h2 style='color:#17A2B8;'>Sudah Selesai</h2><p>Proyek " + projectId + " sudah berstatus Closed sebelumnya.</p></div>");
            }
            
        } else {
            // JIKA REJECT (Kembali ke Need Revision)
            projectSheet.getRange(rowIndex, 18).setValue("Need Revision"); 
            
            if (typeof setupCIMSExecutiveDashboard === "function") {
                setupCIMSExecutiveDashboard();
            }

            if (typeof sendClosingResultEmail_ === "function") {
                sendClosingResultEmail_(projectId, false, note);
            }
            
            return HtmlService.createHtmlOutput("<div style='text-align:center; padding:50px; font-family:sans-serif;'><h2 style='color:#DC3545;'>Ditolak (Need Revision)!</h2><p>Proyek " + projectId + " dikembalikan ke Inisiator untuk perbaikan bukti.</p></div>");
        }
    } catch (err) {
        return HtmlService.createHtmlOutput("<div style='text-align:center; padding:50px; font-family:sans-serif;'><h2 style='color:#DC3545;'>Error Internal!</h2><p>" + err.toString() + "</p></div>");
    }
}