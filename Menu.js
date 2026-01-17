/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER — Menu.gs v2.7
 * ══════════════════════════════════════════════════════════
 * 
 * GRANOT HELPER FUNCTIONS ONLY
 * 
 * NOTE: The onOpen() menu is now in MainMenu.gs
 *       DO NOT add onOpen here - it will conflict!
 * 
 **************************************************************/

// ═══════════════════════════════════════════════════════════════
// ⚠️ NO onOpen() HERE - See MainMenu.gs for the menu system
// ═══════════════════════════════════════════════════════════════


/**
 * ═══════════════════════════════════════════════════════════════
 * GRANOT QUICK ALERT FUNCTIONS
 * ═══════════════════════════════════════════════════════════════
 */

function GRANOT_quickNotWorked() {
  var result = GRANOT_runAlert({
    workStatuses: ["NOT_WORKED"],
    sendMode: "SUMMARY",
    testMode: true
  });
  SpreadsheetApp.getUi().alert(result.message);
}

function GRANOT_quickHotNotWorked() {
  var result = GRANOT_runAlert({
    hotOnly: true,
    sendMode: "SUMMARY",
    testMode: true
  });
  SpreadsheetApp.getUi().alert(result.message);
}

function GRANOT_quickHighValueNotWorked() {
  var result = GRANOT_runAlert({
    highValueOnly: true,
    minCalls: 0,
    maxCalls: 2,
    sendMode: "SUMMARY",
    testMode: true
  });
  SpreadsheetApp.getUi().alert(result.message);
}

function GRANOT_quickStale() {
  var result = GRANOT_runAlert({
    staleOnly: true,
    sendMode: "SUMMARY",
    testMode: true
  });
  SpreadsheetApp.getUi().alert(result.message);
}

function GRANOT_quickZeroCalls() {
  var result = GRANOT_runAlert({
    maxCalls: 0,
    sendMode: "SUMMARY",
    testMode: true
  });
  SpreadsheetApp.getUi().alert(result.message);
}

function GRANOT_managerSummary() {
  var result = GRANOT_runAlert({
    sendMode: "SUMMARY",
    testMode: true
  });
  SpreadsheetApp.getUi().alert(result.message);
}

function GRANOT_teamReports() {
  var result = GRANOT_runAlert({
    sendMode: "BY_TEAM",
    testMode: true
  });
  SpreadsheetApp.getUi().alert(result.message);
}

function GRANOT_adminHealthReport() {
  var result = GRANOT_runAlert({
    sendMode: "ADMIN",
    testMode: true
  });
  SpreadsheetApp.getUi().alert(result.message);
}

function GRANOT_exportAllLeads() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var leads = GRANOT_enrichLeadsWithRC_(GRANOT_loadLeads_());
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var exportName = "GRANOT Export " + Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm");
    var sheet = ss.insertSheet(exportName);
    
    var headers = [
      "Job #", "Rep", "Priority", "Status", "First Name", "Last Name", "Phone", "Email",
      "From State", "To State", "Move Date", "Days to Move", "Est Total", "Source", "Lead Age (days)",
      "Work Status", "Calls", "SMS", "VM", "Inbound", "Longest Call",
      "Is Hot", "Is High Value", "Is Stale"
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#8B1538").setFontColor("#fff").setFontWeight("bold");
    
    var rows = leads.map(function(l) {
      return [
        l.jobNo, l.user || l.jobRep, l.priority, l.status, l.firstName, l.lastName,
        l.phone1, l.email, l.fromState, l.toState, l.moveDateStr, l.daysUntilMove,
        l.estTotal, l.source, l.leadAgeDays,
        l.rc.workStatus, l.rc.callsTotal, l.rc.smsTotal, l.rc.vmTotal, l.rc.inboundTotal,
        l.rc.longestCallDisplay,
        l.rc.isHot ? "YES" : "", l.rc.isHighValue ? "YES" : "", l.rc.isStale ? "YES" : ""
      ];
    });
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    sheet.autoResizeColumns(1, headers.length);
    ui.alert("✅ Exported " + rows.length + " leads to sheet: " + exportName);
  } catch (e) {
    ui.alert("Error: " + e.message);
  }
}