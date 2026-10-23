/*
 * =========================================================
 * CIMS Executive Dashboard (FULL FIX: SMART MATCH & SANITIZER)
 * - Dynamic Punctuation & Parentheses Sanitizer
 * - Flexible Multi-Word Fuzzy Matcher (Fix Putri, Fahmi, Angela)
 * - Auto-Unmerge Layout Safeguard
 * - Complete Gantt Timeline & Stacked Chart
 * File: Dashboard.gs
 * =========================================================
 */
function setupCIMSExecutiveDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName("DASHBOARD");
  var masterSheet = ss.getSheetByName("MASTER_PROJECT");
  var empSheet = ss.getSheetByName("MASTER_EMPLOYEE");

  if (!masterSheet) {
    Logger.log("Error: Sheet MASTER_PROJECT tidak ditemukan!");
    return;
  }

  if (!dashSheet) {
    dashSheet = ss.insertSheet("DASHBOARD", 0);
  } else {
    dashSheet.clear();
    dashSheet.clearConditionalFormatRules();
  }

  // 🟢 UNMERGE SELURUH AREA DASHBOARD UNTUK MENCEGAH RETAKNYA TABEL 3 KOLOM & CHART
  try {
    dashSheet.getRange("A1:AJ120").breakApart();
  } catch (err) {}

  dashSheet.setHiddenGridlines(false);

  // 1. AMBIL DAFTAR KARYAWAN RESMI DARI MASTER_EMPLOYEE
  var validEmployeeList = [];
  var qaqcTeam = [];
  if (empSheet && empSheet.getLastRow() > 1) {
    var empData = empSheet.getRange("A2:B" + empSheet.getLastRow()).getValues();
    for (var e = 0; e < empData.length; e++) {
      var empName = String(empData[e][0]).trim();
      var empRole = String(empData[e][1]).trim();
      if (empName !== "") {
        validEmployeeList.push(empName);
        qaqcTeam.push({ name: empName, role: empRole || "-" });
      }
    }
  }

  // 2. AMBIL DATA DARI MASTER_PROJECT
  var lastRow = masterSheet.getLastRow();
  var validProjects = [];
  if (lastRow > 1) {
    var rawMasterData = masterSheet.getRange(1, 1, lastRow, masterSheet.getLastColumn()).getValues();
    for (var i = 1; i < rawMasterData.length; i++) {
      var pId = String(rawMasterData[i][0] || "").trim();
      if (pId !== "" && pId !== "-") {
        validProjects.push({
          sourceRow: i + 1,
          data: rawMasterData[i]
        });
      }
    }
  }

  // =========================================================
  // 1. HEADER BANNER MAIN
  // =========================================================
  dashSheet.getRange("B2:AE2").merge()
    .setValue("📊 CIMS INTERNAL QA/QC DASHBOARD")
    .setFontSize(14)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground("#8B0000")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  dashSheet.setRowHeight(2, 40);

  // =========================================================
  // 2. KPI CARDS LAYOUT
  // =========================================================
  dashSheet.getRange("B4:E4").merge().setValue("TOTAL PROJECT").setFontSize(9).setFontWeight("bold").setFontColor("#555555").setHorizontalAlignment("center").setBackground("#F0F0F0");
  dashSheet.getRange("B5:E5").merge().setFormula("=COUNTA(MASTER_PROJECT!A2:A)").setFontSize(22).setFontWeight("bold").setFontColor("#8B0000").setHorizontalAlignment("center").setBackground("#FAFAFA");

  dashSheet.getRange("G4:J4").merge().setValue("APPROVED (ON PROGRESS)").setFontSize(9).setFontWeight("bold").setFontColor("#555555").setHorizontalAlignment("center").setBackground("#F0F0F0");
  dashSheet.getRange("G5:J5").merge().setFormula('=COUNTIF(MASTER_PROJECT!R2:R, "Approved")').setFontSize(22).setFontWeight("bold").setFontColor("#1B4F72").setHorizontalAlignment("center").setBackground("#FAFAFA");

  dashSheet.getRange("L4:O4").merge().setValue("CLOSED (DONE)").setFontSize(9).setFontWeight("bold").setFontColor("#555555").setHorizontalAlignment("center").setBackground("#F0F0F0");
  dashSheet.getRange("L5:O5").merge().setFormula('=COUNTIF(MASTER_PROJECT!R2:R, "Closed")').setFontSize(22).setFontWeight("bold").setFontColor("#28A745").setHorizontalAlignment("center").setBackground("#FAFAFA");

  dashSheet.getRange("Q4:T4").merge().setValue("PENDING APPROVAL").setFontSize(9).setFontWeight("bold").setFontColor("#555555").setHorizontalAlignment("center").setBackground("#F0F0F0");
  dashSheet.getRange("Q5:T5").merge().setFormula('=COUNTIF(MASTER_PROJECT!R2:R, "Waiting*")').setFontSize(22).setFontWeight("bold").setFontColor("#D97706").setHorizontalAlignment("center").setBackground("#FAFAFA");

  dashSheet.setRowHeight(4, 25);
  dashSheet.setRowHeight(5, 45);

  dashSheet.getRange("B4:E5").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  dashSheet.getRange("G4:J5").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  dashSheet.getRange("L4:O5").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  dashSheet.getRange("Q4:T5").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

  // =========================================================
  // 3. TABEL RINGKASAN STATUS & KONTRIBUSI TIM
  // =========================================================
  dashSheet.getRange("B8:C8").merge().setValue("RINGKASAN STATUS").setFontWeight("bold").setBackground("#8B0000").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  var statusLabels = [
    ["Waiting Approval", '=COUNTIF(MASTER_PROJECT!R2:R, "Waiting*")'],
    ["Approved", '=COUNTIF(MASTER_PROJECT!R2:R, "Approved")'],
    ["Need Revision", '=COUNTIF(MASTER_PROJECT!R2:R, "Need Revision")'],
    ["Rejected", '=COUNTIF(MASTER_PROJECT!R2:R, "Rejected")'],
    ["Closed", '=COUNTIF(MASTER_PROJECT!R2:R, "Closed")']
  ];
  for (var i = 0; i < statusLabels.length; i++) {
    var row = 9 + i;
    dashSheet.getRange("B" + row).setValue(statusLabels[i][0]).setFontSize(10);
    dashSheet.getRange("C" + row).setFormula(statusLabels[i][1]).setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center");
  }

  // REKAP KONTRIBUSI TIM
  dashSheet.getRange("E8:G8").merge().setValue("TOP 5 PROJECT PER TEAM MEMBER").setFontWeight("bold").setBackground("#8B0000").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  dashSheet.getRange("E9").setValue("Employee Name").setFontWeight("bold").setBackground("#F0F0F0").setFontSize(9);
  dashSheet.getRange("F9").setValue("PIC").setFontWeight("bold").setBackground("#F0F0F0").setHorizontalAlignment("center").setFontSize(9);
  dashSheet.getRange("G9").setValue("Member").setFontWeight("bold").setBackground("#F0F0F0").setHorizontalAlignment("center").setFontSize(9);

  var teamMap = {}; 

  for (var m = 0; m < validProjects.length; m++) {
    var pData = validProjects[m].data;
    var projStatus = String(pData[17]).trim().toLowerCase();
    
    if (projStatus.indexOf("reject") !== -1 || projStatus === "") continue;

    var picStr = String(pData[3]).trim();
    var membersStr = String(pData[7]).trim();

    if (picStr && picStr !== "-") {
      var rawPics = parseIndividualNames_(picStr);
      for (var r = 0; r < rawPics.length; r++) {
        var cleanPic = cleanAndFilterName_(rawPics[r], validEmployeeList);
        if (cleanPic) {
          if (!teamMap[cleanPic]) teamMap[cleanPic] = { pic: 0, member: 0 };
          teamMap[cleanPic].pic += 1;
        }
      }
    }

    if (membersStr && membersStr !== "-") {
      var rawMembers = parseIndividualNames_(membersStr);
      for (var r = 0; r < rawMembers.length; r++) {
        var cleanMem = cleanAndFilterName_(rawMembers[r], validEmployeeList);
        if (cleanMem) {
          if (!teamMap[cleanMem]) teamMap[cleanMem] = { pic: 0, member: 0 };
          teamMap[cleanMem].member += 1;
        }
      }
    }
  }

  var sortedTeam = Object.keys(teamMap).map(function(name) {
    var pCount = teamMap[name].pic;
    var mCount = teamMap[name].member;
    return [name, pCount, mCount, pCount + mCount];
  }).sort(function(a, b) {
    if (b[1] !== a[1]) return b[1] - a[1];
    return b[2] - a[2];
  }); 

  dashSheet.getRange("E10:G14").clearContent();

  if (sortedTeam.length === 0) {
    dashSheet.getRange("E10").setValue("Belum ada data");
    dashSheet.getRange("F10").setValue(0);
    dashSheet.getRange("G10").setValue(0);
  } else {
    var topLimit = Math.min(sortedTeam.length, 5);
    for (var t = 0; t < topLimit; t++) {
      var rIdx = 10 + t;
      dashSheet.getRange("E" + rIdx).setValue(sortedTeam[t][0]).setFontSize(9);
      dashSheet.getRange("F" + rIdx).setValue(sortedTeam[t][1]).setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
      dashSheet.getRange("G" + rIdx).setValue(sortedTeam[t][2]).setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
    }
  }

  dashSheet.getRange("B8:C13").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  dashSheet.getRange("E8:G14").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

  // TABEL BANTUAN CHART STACKED (AG:AI)
  dashSheet.getRange("AG:AI").clearContent(); 
  var chartHelperData = [["Employee", "PIC", "Member"]];
  for (var k = 0; k < sortedTeam.length; k++) {
    chartHelperData.push([sortedTeam[k][0], sortedTeam[k][1], sortedTeam[k][2]]);
  }

  var stackedChartRange;
  if (chartHelperData.length > 1) {
    dashSheet.getRange(1, 33, chartHelperData.length, 3).setValues(chartHelperData);
    stackedChartRange = dashSheet.getRange(1, 33, chartHelperData.length, 3);
  } else {
    dashSheet.getRange("AG1:AI2").setValues([["Employee", "PIC", "Member"], ["Belum Ada Data", 0, 0]]);
    stackedChartRange = dashSheet.getRange("AG1:AI2");
  }

  // =========================================================
  // 4. MONITORING SELURUH TIM QA/QC (APPROVED & CLOSED)
  // =========================================================
  var startUnassignedRow = 16;
  dashSheet.getRange("B" + startUnassignedRow + ":AE" + startUnassignedRow).merge()
    .setValue("👥 QA/QC STAFF PROJECT ASSIGNMENT MONITORING")
    .setFontWeight("bold")
    .setBackground("#2C3E50")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  
  var active2026List = [];
  var active2027List = [];

  for (var m = 0; m < validProjects.length; m++) {
    var pData = validProjects[m].data;
    var projStatus = String(pData[17] || "").trim().toLowerCase();

    // PERBAIKAN: Sertakan kata "waiting" dan "approval" agar status closing tetap terbaca aktif
    var isApprovedOrClosed = projStatus.indexOf("approved") !== -1 || 
                             projStatus.indexOf("closed") !== -1 || 
                             projStatus.indexOf("waiting") !== -1 || 
                             projStatus.indexOf("approval") !== -1;
    
    if (!isApprovedOrClosed) continue;

    var rawStart = pData[1];
    var pDate = rawStart ? new Date(rawStart) : new Date(2026, 0, 1);
    var pYear = isNaN(pDate.getFullYear()) ? 2026 : pDate.getFullYear();

    var picStr = String(pData[3]).trim();
    var membersStr = String(pData[7]).trim();

    var rawPeople = [];
    if (picStr && picStr !== "-") rawPeople = rawPeople.concat(parseIndividualNames_(picStr));
    if (membersStr && membersStr !== "-") rawPeople = rawPeople.concat(parseIndividualNames_(membersStr));

    for (var pIdx = 0; pIdx < rawPeople.length; pIdx++) {
      var cleanedName = cleanAndFilterName_(rawPeople[pIdx], validEmployeeList);
      if (!cleanedName) continue;
      
      var matchLower = cleanedName.toLowerCase();
      if (pYear <= 2026) active2026List.push(matchLower);
      if (pYear >= 2027) active2027List.push(matchLower);
    }
  }

  var nextRow = startUnassignedRow + 1;

  if (qaqcTeam.length === 0) {
    dashSheet.getRange("B" + nextRow + ":AE" + nextRow).merge()
      .setValue("⚠️ Tab 'MASTER_EMPLOYEE' belum diisi data karyawan.")
      .setFontStyle('italic').setHorizontalAlignment("center").setFontColor("#D97706");
    nextRow++;
  } else {
    dashSheet.getRange("B" + nextRow).setValue("No").setFontWeight("bold").setBackground("#F0F0F0").setHorizontalAlignment("center");
    dashSheet.getRange("C" + nextRow + ":E" + nextRow).merge().setValue("Employee Name").setFontWeight("bold").setBackground("#F0F0F0");
    dashSheet.getRange("F" + nextRow + ":H" + nextRow).merge().setValue("Role / Position").setFontWeight("bold").setBackground("#F0F0F0");
    
    dashSheet.getRange("I" + nextRow + ":J" + nextRow).merge().setValue("Status (2026)").setFontWeight("bold").setBackground("#F0F0F0").setHorizontalAlignment("center");
    dashSheet.getRange("K" + nextRow + ":L" + nextRow).merge().setValue("Status (2027)").setFontWeight("bold").setBackground("#F0F0F0").setHorizontalAlignment("center");
    nextRow++;

    for (var u = 0; u < qaqcTeam.length; u++) {
      var staff = qaqcTeam[u];
      var staffLower = staff.name.toLowerCase().trim();

      var has2026 = active2026List.some(function(name) { return staffLower === name; });
      var has2027 = active2027List.some(function(name) { return staffLower === name; });

      dashSheet.getRange("B" + nextRow).setValue(u + 1).setHorizontalAlignment("center");
      dashSheet.getRange("C" + nextRow + ":E" + nextRow).merge().setValue(staff.name).setFontWeight("bold");
      dashSheet.getRange("F" + nextRow + ":H" + nextRow).merge().setValue(staff.role);
      
      var b2026 = dashSheet.getRange("I" + nextRow + ":J" + nextRow).merge();
      if (has2026) {
        b2026.setValue("Active Project").setFontWeight("bold").setFontColor("#155724").setBackground("#D4EDDA").setHorizontalAlignment("center");
      } else {
        b2026.setValue("No Project").setFontWeight("bold").setFontColor("#7f8c8d").setBackground("#F2F3F4").setHorizontalAlignment("center");
      }

      var b2027 = dashSheet.getRange("K" + nextRow + ":L" + nextRow).merge();
      if (has2027) {
        b2027.setValue("Active Project").setFontWeight("bold").setFontColor("#155724").setBackground("#D4EDDA").setHorizontalAlignment("center");
      } else {
        b2027.setValue("No Project").setFontWeight("bold").setFontColor("#7f8c8d").setBackground("#F2F3F4").setHorizontalAlignment("center");
      }

      nextRow++;
    }
    dashSheet.getRange("B" + (startUnassignedRow + 1) + ":L" + (nextRow - 1)).setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  }

  // =========================================================
  // 5. MATRIKS GANTT CHART TIMELINE
  // =========================================================
  var startGanttRow = nextRow + 1;
  dashSheet.getRange("B" + startGanttRow + ":AE" + startGanttRow).merge()
    .setValue("📅 GANTT CHART PROJECT TIMELINE")
    .setFontWeight("bold")
    .setBackground("#1A2530")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  
  var yearHeaderRow = startGanttRow + 1;
  dashSheet.getRange("B" + yearHeaderRow + ":G" + yearHeaderRow).merge()
    .setValue("PROJECT DETAILS")
    .setFontWeight("bold")
    .setBackground("#1B4F72")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  
  dashSheet.getRange("H" + yearHeaderRow + ":S" + yearHeaderRow).merge()
    .setValue("PERIODE 2026")
    .setFontWeight("bold")
    .setBackground("#2E86C1")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  
  dashSheet.getRange("T" + yearHeaderRow + ":AE" + yearHeaderRow).merge()
    .setValue("PERIODE 2027")
    .setFontWeight("bold")
    .setBackground("#1F618D")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  var ganttHeaderRow = yearHeaderRow + 1;
  dashSheet.getRange("B" + ganttHeaderRow).setValue("Project Title").setFontWeight("bold").setBackground("#2C3E50").setFontColor("#FFFFFF");
  dashSheet.getRange("C" + ganttHeaderRow).setValue("PIC").setFontWeight("bold").setBackground("#2C3E50").setFontColor("#FFFFFF");
  dashSheet.getRange("D" + ganttHeaderRow).setValue("Members").setFontWeight("bold").setBackground("#2C3E50").setFontColor("#FFFFFF");
  dashSheet.getRange("E" + ganttHeaderRow).setValue("Status").setFontWeight("bold").setBackground("#2C3E50").setFontColor("#FFFFFF");
  dashSheet.getRange("F" + ganttHeaderRow).setValue("Planned Timeline").setFontWeight("bold").setBackground("#2C3E50").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  dashSheet.getRange("G" + ganttHeaderRow).setValue("Actual Timeline").setFontWeight("bold").setBackground("#2C3E50").setFontColor("#FFFFFF").setHorizontalAlignment("center");

  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  for (var mIdx = 0; mIdx < 24; mIdx++) {
    var colNum = 8 + mIdx;
    var monthText = monthNames[mIdx % 12];
    dashSheet.getRange(ganttHeaderRow, colNum)
      .setValue(monthText)
      .setFontWeight("bold")
      .setBackground("#2C3E50")
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");
  }

  var totalProjects = validProjects.length;
  var projectStartRow = ganttHeaderRow + 1;
  var projectEndRow = projectStartRow + Math.max(1, totalProjects) - 1;

  if (totalProjects <= 0) {
    dashSheet.getRange("B" + projectStartRow + ":AE" + projectStartRow).merge().setValue("Belum ada data project terdaftar.").setFontStyle('italic').setHorizontalAlignment("center");
  } else {
    for (var p = 0; p < totalProjects; p++) {
      var rowIdx = projectStartRow + p;
      var sourceRow = validProjects[p].sourceRow;

      dashSheet.getRange("B" + rowIdx).setFormula("=MASTER_PROJECT!C" + sourceRow);
      dashSheet.getRange("C" + rowIdx).setFormula("=MASTER_PROJECT!D" + sourceRow);
      dashSheet.getRange("D" + rowIdx).setFormula("=IF(OR(MASTER_PROJECT!H" + sourceRow + '="", MASTER_PROJECT!H' + sourceRow + '="-"), "-", MASTER_PROJECT!H' + sourceRow + ")");
      dashSheet.getRange("E" + rowIdx).setFormula("=MASTER_PROJECT!R" + sourceRow);
      
      var plannedFormula = '=LET(' +
        'rawStart, MASTER_PROJECT!$B' + sourceRow + ', ' +
        'startDate, IF(ISBLANK(rawStart), "-", IF(ISNUMBER(rawStart), rawStart, IFERROR(DATEVALUE(rawStart), "-"))), ' +
        'estStr, MASTER_PROJECT!$P' + sourceRow + ', ' +
        'addM, IFERROR(IFS(' +
          'ISNUMBER(SEARCH("6 Bul", estStr)), 5, ' +
          'ISNUMBER(SEARCH("12 Bul", estStr)), 11, ' +
          'ISNUMBER(SEARCH("3 Bul", estStr)), 2, ' +
          'ISNUMBER(SEARCH("Bul", estStr)), 1, ' +
          'ISNUMBER(SEARCH("Tahun", estStr)), 11' +
        '), 0), ' +
        'planEnd, EDATE(startDate, addM), ' +
        'IF(startDate="-", "-", ' +
          'TEXT(startDate, "dd mmm yyyy") & " - " & TEXT(planEnd, "dd mmm yyyy") & IF(ISBLANK(estStr), "", " (Est: " & estStr & ")")' +
        ')' +
      ')';
      dashSheet.getRange("F" + rowIdx).setFormula(plannedFormula).setHorizontalAlignment("center").setFontSize(9);

      var actualFormula = '=LET(' +
        'rawStart, MASTER_PROJECT!$B' + sourceRow + ', ' +
        'startDate, IF(ISBLANK(rawStart), "-", IF(ISNUMBER(rawStart), rawStart, IFERROR(DATEVALUE(rawStart), "-"))), ' +
        'rawDue, MASTER_PROJECT!$S' + sourceRow + ', ' +
        'rawComp, MASTER_PROJECT!$T' + sourceRow + ', ' +
        'pStatus, MASTER_PROJECT!$R' + sourceRow + ', ' +
        'actualEnd, IF(ISBLANK(rawComp), IF(ISBLANK(rawDue), startDate, rawDue), rawComp), ' +
        'IF(startDate="-", "-", ' +
          'IFS(' +
            'pStatus="Closed", TEXT(startDate, "dd mmm yyyy") & " - " & TEXT(actualEnd, "dd mmm yyyy") & " (Closed)", ' +
            'pStatus="Approved", "On Progress", ' +
            'ISNUMBER(SEARCH("Waiting", pStatus)), "Pending Approval", ' +
            'pStatus="Need Revision", "Need Revision", ' +
            'TRUE, pStatus' +
          ')' +
        ')' +
      ')';
      dashSheet.getRange("G" + rowIdx).setFormula(actualFormula).setHorizontalAlignment("center").setFontSize(9);

      for (var mIdx = 0; mIdx < 24; mIdx++) {
        var colNum = 8 + mIdx;
        var yearVal = 2026 + Math.floor(mIdx / 12);
        var monthVal = (mIdx % 12) + 1;
        var curMIndex = yearVal * 12 + monthVal;

        var dotBarFormula = '=LET(' +
          'pStatus, MASTER_PROJECT!$R' + sourceRow + ', ' +
          'isActive, OR(pStatus="Approved", pStatus="Closed"), ' +
          'rawStart, MASTER_PROJECT!$B' + sourceRow + ', ' +
          'startDate, IF(ISBLANK(rawStart), DATE(2026,1,1), IF(ISNUMBER(rawStart), rawStart, IFERROR(DATEVALUE(rawStart), DATE(2026,1,1)))), ' +
          'estStr, MASTER_PROJECT!$P' + sourceRow + ', ' +
          'addM, IFERROR(IFS(' +
            'ISNUMBER(SEARCH("6 Bul", estStr)), 5, ' +
            'ISNUMBER(SEARCH("12 Bul", estStr)), 11, ' +
            'ISNUMBER(SEARCH("3 Bul", estStr)), 2, ' +
            'ISNUMBER(SEARCH("Bul", estStr)), 1, ' +
            'ISNUMBER(SEARCH("Tahun", estStr)), 11' +
          '), 0), ' +
          'rawDue, MASTER_PROJECT!$S' + sourceRow + ', ' +
          'rawComp, MASTER_PROJECT!$T' + sourceRow + ', ' +
          'actualEnd, IF(ISBLANK(rawComp), IF(ISBLANK(rawDue), startDate, rawDue), rawComp), ' +
          'endDate, IF(pStatus="Closed", actualEnd, EDATE(startDate, addM)), ' +
          'sMIndex, YEAR(startDate)*12 + MONTH(startDate), ' +
          'eMIndex, YEAR(endDate)*12 + MONTH(endDate), ' +
          'curMIndex, ' + curMIndex + ', ' +
          'inRange, AND(curMIndex >= sMIndex, curMIndex <= eMIndex), ' +
          'IFS(' +
            'NOT(isActive), "", ' +
            'NOT(inRange), "", ' +
            'AND(curMIndex = sMIndex, curMIndex = eMIndex), "🔵 🔴", ' +
            'curMIndex = sMIndex, "🔵", ' +
            'curMIndex = eMIndex, "🔴", ' +
            'TRUE, "━━━━"' +
          ')' +
        ')';

        dashSheet.getRange(rowIdx, colNum).setFormula(dotBarFormula).setHorizontalAlignment("center").setFontSize(10);
      }
    }

    dashSheet.getRange(startGanttRow, 2, projectEndRow - startGanttRow + 1, 30)
      .setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

    var ganttRange = dashSheet.getRange(projectStartRow, 8, Math.max(1, totalProjects), 24);
    var ruleStart = SpreadsheetApp.newConditionalFormatRule().whenTextContains("🔵").setBackground("#D4E6F1").setFontColor("#1B4F72").setBold(true).setRanges([ganttRange]).build();
    var ruleProgress = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("━━━━").setBackground("#3498DB").setFontColor("#FFFFFF").setBold(true).setRanges([ganttRange]).build();
    var ruleEnd = SpreadsheetApp.newConditionalFormatRule().whenTextContains("🔴").setBackground("#FADBD8").setFontColor("#78281F").setBold(true).setRanges([ganttRange]).build();

    var rules = dashSheet.getConditionalFormatRules();
    rules.push(ruleStart, ruleProgress, ruleEnd);
    dashSheet.setConditionalFormatRules(rules);
  }

  // Set Width Kolom
  dashSheet.setColumnWidth(1, 20);
  dashSheet.setColumnWidth(2, 220);
  dashSheet.setColumnWidth(3, 110);
  dashSheet.setColumnWidth(4, 130);
  dashSheet.setColumnWidth(5, 120);
  dashSheet.setColumnWidth(6, 180);
  dashSheet.setColumnWidth(7, 180);
  
  for (var mIdx = 0; mIdx < 24; mIdx++) {
    dashSheet.setColumnWidth(8 + mIdx, 55);
  }

  // =========================================================
  // 6. CHARTS VISUAL
  // =========================================================
  var existingCharts = dashSheet.getCharts();
  for (var c = 0; c < existingCharts.length; c++) {
    dashSheet.removeChart(existingCharts[c]);
  }

  // Chart 1: Status Distribution
  var chart1 = dashSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dashSheet.getRange("B9:C13"))
      .setPosition(projectEndRow + 3, 2, 0, 0)
      .setOption('title', 'Distribusi Status Project QA/QC')
      .setOption('legend', {position: 'none'})
      .setOption('colors', ['#8B0000'])
      .build();
  dashSheet.insertChart(chart1);

  // Chart 2: Stacked Column Chart (PIC vs Member)
  var chart2 = dashSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(stackedChartRange)
      .setNumHeaders(1)
      .setPosition(projectEndRow + 3, 7, 0, 0)
      .setOption('title', 'Kontribusi Project per Karyawan (PIC & Member)')
      .setOption('isStacked', true)
      .setOption('legend', {position: 'right', textStyle: {fontSize: 11}})
      .setOption('colors', ['#1B4F72', '#28A745'])
      .setOption('vAxis', {title: 'Jumlah Project', format: '0'})
      .build();
  dashSheet.insertChart(chart2);

  Logger.log("✅ Final Dashboard CIMS berhasil diperbarui!");
}

/**
 * HELPER FUNCTION: Smart Name Parser (Membersihkan Tanda Baca & Pemisah)
 */
function parseIndividualNames_(rawText) {
  if (!rawText) return [];
  
  var cleaned = String(rawText)
    .replace(/\(.*?\)/g, "") // Hapus teks dalam kurung misal (Global QAQC)
    .replace(/\s+dan\s+/gi, ",")
    .replace(/\s+&\s+/g, ",")
    .replace(/;/g, ",")
    .replace(/\//g, ",")
    .replace(/\n/g, ",");
    
  var rawNames = cleaned.split(",");
  var nameList = [];

  for (var i = 0; i < rawNames.length; i++) {
    var name = rawNames[i].replace(/^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$/g, "").trim();
    if (name.length > 0 && name !== "-") {
      nameList.push(name);
    }
  }
  return nameList;
}

/**
 * HELPER FUNCTION: Smart Fuzzy Employee Matcher (Bebas Error Keterangan/Strip)
 */
function cleanAndFilterName_(rawName, validEmployeeList) {
  if (!rawName || !validEmployeeList || validEmployeeList.length === 0) return null;

  // 1. Pembersihan karakter pengganggu (kurung, strip, karakter khusus)
  var cleaned = String(rawName)
    .replace(/\(.*?\)/g, "")        // Hapus isi dalam kurung
    .replace(/[^a-zA-Z0-9\s]/g, " ")  // Ganti semua simbol/tanda baca (- / . dll) dengan spasi
    .replace(/\s+/g, " ")           // Rapikan spasi ganda
    .trim();

  if (!cleaned || cleaned === "-") return null;
  var cleanLower = cleaned.toLowerCase();

  // Abaikan kata-kata keterangan non-nama
  if (cleanLower.indexOf("soon") !== -1 || cleanLower.indexOf("akan") !== -1 || cleanLower.indexOf("dengan") !== -1 || cleanLower.indexOf("juga") !== -1) {
    return null;
  }

  var noiseWords = ["procurement", "bd", "rnd", "r&d", "qaqc", "qa/qc", "bersama", "tim", "department", "dept", "team", "pak", "bu", "global", "specialist"];
  var inputWords = cleanLower.split(" ").filter(function(w) { 
    return w.length > 0 && noiseWords.indexOf(w) === -1; 
  });

  if (inputWords.length === 0) return null;

  // 2. EXACT MATCH CHECK (100% Persis)
  for (var i = 0; i < validEmployeeList.length; i++) {
    var officialName = validEmployeeList[i];
    var officialClean = officialName.replace(/[^a-zA-Z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
    if (cleanLower === officialClean.toLowerCase()) {
      return officialName; // Kembalikan Nama Resmi dari MASTER_EMPLOYEE
    }
  }

  // 3. FLEXIBLE TOKEN / FUZZY MATCH
  var bestMatch = null;
  var maxScore = 0;

  for (var i = 0; i < validEmployeeList.length; i++) {
    var officialName = validEmployeeList[i];
    var officialClean = officialName.replace(/[^a-zA-Z0-9\s]/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
    var officialWords = officialClean.split(" ").filter(function(w) { 
      return w.length > 0 && noiseWords.indexOf(w) === -1; 
    });

    if (officialWords.length === 0) continue;

    var matchedWordCount = 0;
    for (var w = 0; w < inputWords.length; w++) {
      var word = inputWords[w];
      if (word.length >= 2 && officialWords.indexOf(word) !== -1) {
        matchedWordCount++;
      }
    }

    // Syarat Lolos Match:
    // A. Semua kata input ada di nama resmi (misal "Putri Awwalia" -> "Putri Awwalia Adistyaningrum")
    // B. Minimal 2 kata cocok (misal "Adistyaningrum Putri" -> "Putri Awwalia Adistyaningrum")
    // C. Jika cuma 1 kata cocok (misal "Putri"), pastikan kata tersebut cukup panjang (>=4 huruf)
    var isAllInputMatched = (matchedWordCount === inputWords.length && matchedWordCount > 0);
    var isMultiWordMatched = (matchedWordCount >= 2);
    var isSingleLongWordMatched = (matchedWordCount === 1 && inputWords.length === 1 && inputWords[0].length >= 4);

    if (isAllInputMatched || isMultiWordMatched || isSingleLongWordMatched) {
      var score = matchedWordCount * 10 - Math.abs(officialWords.length - inputWords.length);
      if (score > maxScore) {
        maxScore = score;
        bestMatch = officialName;
      }
    }
  }

  return bestMatch;
}

/**
 * MEMUNCULKAN MENU DI SPREADSHEET
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 CIMS Menu')
    .addItem('⚡ Refresh Dashboard Now', 'setupCIMSExecutiveDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 Email Reminders')
      .addItem('Send Closing Reminders (All)', 'menuSendClosingReminders')
      .addItem('Send Monthly Reminders (All)', 'menuSendMonthlyReminders')
      .addItem('Send Pending Approval Reminders', 'menuSendPendingApprovalReminders')
      .addSeparator()
      .addItem('Re-send Approval Email (by ID)', 'menuResendApprovalEmail')
      .addItem('Re-send Result Email (by ID)', 'menuResendResultEmail')
      .addItem('Re-send Closing Approval (by ID)', 'menuResendClosingApproval')
      .addItem('Re-send Final Approval (by ID)', 'menuResendFinalApproval'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📩 Closing Links')
      .addItem('Kirim Link Closing (1 Project)', 'sendManualClosingLink')
      .addItem('Blast Link Closing (Ke SEMUA Project Approved)', 'blastClosingLinkToApprovedProjects'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📢 Member Notifications')
      .addItem('Kirim Info Proyek ke Member (1 Project)', 'sendManualInfoToMembers'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Utilities')
      .addItem('Generate Next Project ID', 'menuGenerateProjectID')
      .addItem('List Active Managers', 'menuListManagers'))
    .addSeparator()
    .addSubMenu(ui.createMenu('ℹ️ Help')
      .addItem('Documentation', 'menuDocumentation')
      .addItem('How to Use', 'menuHowToUse'))
    .addToUi();
}

/*********************************************
 * MENU HANDLER FUNCTIONS
 *********************************************/

function menuSendClosingReminders() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Send Closing Reminders', 'This will send deadline reminder emails to all initiators with projects due within 7 days.\n\nContinue?', ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    try {
      sendAutoClosingReminders();
      ui.alert('Success', 'Closing reminders sent successfully!', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Error', 'Failed to send reminders: ' + e.message, ui.ButtonSet.OK);
    }
  }
}

function menuSendMonthlyReminders() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Send Monthly Reminders', 'This will send monthly status update emails to all active project initiators.\n\nContinue?', ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    try {
      sendMonthlyInitiatorReminders();
      ui.alert('Success', 'Monthly reminders sent successfully!', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Error', 'Failed to send reminders: ' + e.message, ui.ButtonSet.OK);
    }
  }
}

function menuSendPendingApprovalReminders() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Send Pending Approval Reminders', 'This will send reminder emails to all approvers with pending closing approvals.\n\nContinue?', ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    try {
      sendPendingClosingApprovalReminders();
      ui.alert('Success', 'Pending approval reminders sent successfully!', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Error', 'Failed to send reminders: ' + e.message, ui.ButtonSet.OK);
    }
  }
}

function menuResendApprovalEmail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Re-send Approval Email', 'Enter Project ID (e.g., CI-2026-0001):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var projectId = response.getResponseText().trim();
    if (projectId) {
      try {
        sendApprovalEmail_(projectId, 1);
        ui.alert('Success', 'Approval email re-sent for ' + projectId, ui.ButtonSet.OK);
      } catch (e) {
        ui.alert('Error', 'Failed to send email: ' + e.message, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Error', 'Project ID cannot be empty', ui.ButtonSet.OK);
    }
  }
}

function menuResendResultEmail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Re-send Result Email', 'Enter Project ID:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var projectId = response.getResponseText().trim();
    if (projectId) {
      try {
        var project = getProjectDetailsByID_(projectId);
        if (project) {
          sendResultEmail_(projectId, project.status);
          ui.alert('Success', 'Result email re-sent for ' + projectId, ui.ButtonSet.OK);
        } else {
          ui.alert('Error', 'Project not found: ' + projectId, ui.ButtonSet.OK);
        }
      } catch (e) {
        ui.alert('Error', 'Failed to send email: ' + e.message, ui.ButtonSet.OK);
      }
    }
  }
}

function menuResendClosingApproval() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Re-send Closing Approval', 'Enter Project ID:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var projectId = response.getResponseText().trim();
    if (projectId) {
      try {
        var project = getProjectDetailsByID_(projectId);
        if (project) {
          var formData = {
            projectId: projectId,
            finalEvidence: project.remarks || "-",
            closingSummary: project.remarks || "-",
            benefitCategory: project.impact || "Uncategorized",
            claimedSaving: project.impact || "0",
            completionDate: project.dueDate || new Date().toLocaleDateString()
          };
          sendClosingApprovalEmail_(projectId, formData);
          ui.alert('Success', 'Closing approval email re-sent for ' + projectId, ui.ButtonSet.OK);
        } else {
          ui.alert('Error', 'Project not found: ' + projectId, ui.ButtonSet.OK);
        }
      } catch (e) {
        ui.alert('Error', 'Failed to send email: ' + e.message, ui.ButtonSet.OK);
      }
    }
  }
}

function menuResendFinalApproval() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Re-send Final Approval', 'Enter Project ID:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var projectId = response.getResponseText().trim();
    if (projectId) {
      try {
        sendFinalClosingApprovalEmail_(projectId);
        ui.alert('Success', 'Final approval email re-sent to Pak Bayun for ' + projectId, ui.ButtonSet.OK);
      } catch (e) {
        ui.alert('Error', 'Failed to send email: ' + e.message, ui.ButtonSet.OK);
      }
    }
  }
}

function menuGenerateProjectID() {
  var ui = SpreadsheetApp.getUi();
  try {
    var nextId = generateProjectID();
    var message = 'The next project ID will be: ' + nextId;
    ui.alert('Next Project ID', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to generate project ID: ' + e.message, ui.ButtonSet.OK);
  }
}

function menuListManagers() {
  var ui = SpreadsheetApp.getUi();
  try {
    var managers = getManagers();
    if (managers.length === 0) {
      ui.alert('No Active Managers', 'No active managers found in MASTER_MANAGER sheet.', ui.ButtonSet.OK);
    } else {
      var list = managers.map(function(m, i) { return (i+1) + '. ' + m.name + ' (' + m.role + ') - ' + m.email; }).join('\n');
      ui.alert('Active Managers (' + managers.length + ')', list, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Failed to list managers: ' + e.message, ui.ButtonSet.OK);
  }
}
function menuDocumentation() {
  var ui = SpreadsheetApp.getUi();
  var msg = 'CIMS Documentation' + String.fromCharCode(10) + String.fromCharCode(10) +
    'Documentation:' + String.fromCharCode(10) + String.fromCharCode(10) +
    '- Project submission via Google Form' + String.fromCharCode(10) +
    '- Automated approval workflow (L1, L2, Final)' + String.fromCharCode(10) +
    '- Email notifications for all stakeholders' + String.fromCharCode(10) +
    '- Real-time dashboard monitoring' + String.fromCharCode(10) +
    '- Automated reminders for deadlines' + String.fromCharCode(10) + String.fromCharCode(10) +
    'For detailed documentation, contact QAQC Department.';
  ui.alert('CIMS Documentation', msg, ui.ButtonSet.OK);
}

function menuHowToUse() {
  var ui = SpreadsheetApp.getUi();
  var msg = 'CIMS Menu Guide:' + String.fromCharCode(10) + String.fromCharCode(10) +
    'EMAIL REMINDERS:' + String.fromCharCode(10) +
    '- Send bulk reminder emails' + String.fromCharCode(10) +
    '- Re-send specific emails by Project ID' + String.fromCharCode(10) + String.fromCharCode(10) +
    'CLOSING LINKS:' + String.fromCharCode(10) +
    '- Send closing form link to specific project' + String.fromCharCode(10) +
    '- Blast closing links to all approved projects' + String.fromCharCode(10) + String.fromCharCode(10) +
    'MEMBER NOTIFICATIONS:' + String.fromCharCode(10) +
    '- Send project info to team members' + String.fromCharCode(10) + String.fromCharCode(10) +
    'UTILITIES:' + String.fromCharCode(10) +
    '- Generate next Project ID' + String.fromCharCode(10) +
    '- List all active managers' + String.fromCharCode(10) + String.fromCharCode(10) +
    'Need help? Contact QAQC Department.';
  ui.alert('How to Use CIMS Menu', msg, ui.ButtonSet.OK);
}