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
    if (String(data[i][5] || "").toUpperCase() === "YES") {
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
