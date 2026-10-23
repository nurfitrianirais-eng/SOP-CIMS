/*************************************************
 * UTILITY & HELPER FUNCTIONS (CIMS SYSTEM)
 *************************************************/

function generateProjectID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(typeof CONFIG !== "undefined" ? CONFIG.PROJECT_SHEET : "MASTER_PROJECT");
  const year = new Date().getFullYear();
  const running = Math.max(sheet.getLastRow(), 2) - 1;
  const prefix = (typeof CONFIG !== "undefined" && CONFIG.PROJECT_PREFIX) ? CONFIG.PROJECT_PREFIX : "CI";
  return prefix + "-" + year + "-" + Utilities.formatString("%04d", running);
}

function getManagers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = (typeof CONFIG !== "undefined" && CONFIG.MANAGER_SHEET) ? CONFIG.MANAGER_SHEET : "MASTER_MANAGER";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let managers = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4] || "").toUpperCase() === "YES") {
      managers.push({
        name: data[i][0],
        email: data[i][1],
        role: data[i][2],
        order: data[i][3]
      });
    }
  }
  managers.sort((a, b) => a.order - b.order);
  return managers;
}

function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy");
}

function sendEmail(to, subject, body) {
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: body });
}

function getApproverByOrder_(order) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MASTER_MANAGER");
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][3]) === order && String(data[i][4] || "").toLowerCase() === "yes") {
      return { name: data[i][0], email: data[i][1], role: data[i][2], order: Number(data[i][3]) };
    }
  }
  return null;
}

function getNextApprovalOrder_(approver) {
  switch (String(approver).toLowerCase()) {
    case "manager": return 2;
    case "qa": return 3;
    case "qc": return 0;
    default: return 0;
  }
}

function getInitiatorEmail_(projectId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MASTER_PROJECT");
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  var searchId = String(projectId || "").trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toLowerCase() === searchId) {
      return data[i][5] || data[i][3] || "";
    }
  }
  return "";
}

function getSupportingEvidenceLink_(projectId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MASTER_PROJECT");
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  var searchId = String(projectId || "").trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toLowerCase() === searchId) {
      return data[i][12] || "";
    }
  }
  return "";
}

/**
 * HELPER UTAMA: Mengambil Data Proyek Berdasarkan Project ID
 */
function getProjectById_(projectId) {
  if (!projectId) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MASTER_PROJECT");
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var searchId = String(projectId).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][0] || "").trim().toLowerCase();
    if (rowId === searchId) {
      return {
        id: data[i][0],
        title: data[i][2],
        initiator: data[i][3],
        managerName: data[i][10],
        status: data[i][17],
        dueDate: data[i][18],
        remarks: data[i][21],
        row: data[i]
      };
    }
  }
  return null;
}

// Alias pendukung
function getProjectId_(projectId) { return getProjectById_(projectId); }