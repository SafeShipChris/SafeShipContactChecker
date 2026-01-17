/**************************************************************
 * 🔍 TEST: Check Cubic Feet for specific job
 * Run: testCubicFeetForJob()
 **************************************************************/

function testCubicFeetForJob() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var CFG = getConfig_();
  
  var lines = [];
  lines.push("📦 CUBIC FEET LIVE TEST");
  lines.push("═".repeat(50));
  
  // Read CALL & VOICEMAIL TRACKER using same method as scanner
  var sh = ss.getSheetByName("CALL & VOICEMAIL TRACKER");
  var values = sh.getRange(CFG.TRACKER_RANGE_A1).getValues();
  
  lines.push("\nRange: " + CFG.TRACKER_RANGE_A1);
  lines.push("Rows read: " + values.length);
  lines.push("CUBIC_FEET_IDX: " + CFG.TRACKER_COLS.CUBIC_FEET_IDX);
  
  // Check first 5 data rows
  lines.push("\n=== First 5 rows (Job, User, CF) ===");
  for (var i = 0; i < Math.min(5, values.length); i++) {
    var row = values[i];
    var job = row[CFG.TRACKER_COLS.JOB_IDX];
    var user = row[CFG.TRACKER_COLS.USERNAME_IDX];
    var cf = row[CFG.TRACKER_COLS.CUBIC_FEET_IDX];
    
    lines.push("Row " + (i+3) + ": Job=" + job + ", User=" + user + ", CF='" + cf + "' (type: " + typeof cf + ")");
  }
  
  // Look for job 7136378 specifically
  lines.push("\n=== Looking for Job 7136378 ===");
  var found = false;
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var job = row[CFG.TRACKER_COLS.JOB_IDX];
    
    if (String(job).indexOf("7136378") > -1 || job == 7136378) {
      var user = row[CFG.TRACKER_COLS.USERNAME_IDX];
      var cf = row[CFG.TRACKER_COLS.CUBIC_FEET_IDX];
      lines.push("FOUND at row " + (i+3) + ":");
      lines.push("  Job: " + job + " (type: " + typeof job + ")");
      lines.push("  User: " + user);
      lines.push("  CF: '" + cf + "' (type: " + typeof cf + ")");
      lines.push("  CF truthy? " + (cf ? "YES" : "NO"));
      found = true;
      break;
    }
  }
  
  if (!found) {
    lines.push("NOT FOUND in CALL & VOICEMAIL TRACKER");
  }
  
  // Show raw Column K values
  lines.push("\n=== Raw Column K (first 10 values) ===");
  var colK = sh.getRange("K3:K12").getValues();
  for (var i = 0; i < colK.length; i++) {
    lines.push("K" + (i+3) + ": '" + colK[i][0] + "' (type: " + typeof colK[i][0] + ")");
  }
  
  ui.alert(lines.join("\n"));
}

/**************************************************************
 * 🔧 RC_Enrichment PATCH — SMS Status Fix
 * 
 * PROBLEM: SMS counts show 1→1 but Last 5 SMS shows "Never"
 * CAUSE: Status filter is too strict - only accepts "Delivered" or "Sent"
 * 
 * INSTRUCTIONS:
 * 1. First run RC_debugSMSMatching() to see what statuses exist
 * 2. Then apply the fix to RC_Enrichment.gs
 **************************************************************/



/**
 * 🔍 DIAGNOSTIC: Test a specific phone number
 */
function RC_debugSpecificPhone() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var cell = ss.getActiveCell();
  var phone = String(cell.getValue() || "").trim();
  
  if (!phone) {
    ui.alert("Please select a cell with a phone number first.");
    return;
  }
  
  // Normalize phone
  var digits = phone.replace(/\D/g, "");
  if (digits.length === 11 && digits[0] === "1") digits = digits.substring(1);
  if (digits.length > 10) digits = digits.slice(-10);
  var normalized = digits.length === 10 ? digits : "";
  
  var lines = [];
  lines.push("🔍 DEBUG FOR PHONE: " + phone);
  lines.push("   Normalized: " + normalized);
  lines.push("═".repeat(50));
  lines.push("");
  
  // Search in SMS sheets
  var sheets = ["RC SMS LOG", "RC SMS LOG YESTERDAY"];
  
  sheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    var found = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var direction = String(row[0] || "").toLowerCase();
      var recipient = String(row[5] || "").replace(/\D/g, "");
      var sender = String(row[3] || "").replace(/\D/g, "");
      var status = String(row[9] || "");
      var datetime = row[7];
      
      // Normalize for comparison
      if (recipient.length === 11 && recipient[0] === "1") recipient = recipient.substring(1);
      if (sender.length === 11 && sender[0] === "1") sender = sender.substring(1);
      
      var matches = (direction === "outbound" && recipient === normalized) ||
                    (direction === "inbound" && sender === normalized);
      
      if (matches) {
        found.push({
          direction: direction,
          status: status,
          datetime: datetime,
          row: i + 2
        });
      }
    }
    
    lines.push("📋 " + sheetName + ": " + found.length + " records found");
    
    if (found.length > 0) {
      found.slice(0, 5).forEach(function(f) {
        var wouldMatch = f.status.toLowerCase().indexOf("delivered") >= 0 || 
                         f.status.toLowerCase().indexOf("sent") >= 0 || 
                         f.status === "";
        var icon = wouldMatch ? "✅" : "❌";
        lines.push("   Row " + f.row + ": " + f.direction + " | Status=\"" + f.status + "\" " + icon);
      });
      
      if (found.length > 5) {
        lines.push("   ... and " + (found.length - 5) + " more");
      }
    }
    lines.push("");
  });
  
  lines.push("══════════════════════════════════════════════════");
  lines.push("❌ = Won't appear in Last 5 SMS due to status filter");
  lines.push("✅ = Will appear in Last 5 SMS");
  
  ui.alert(lines.join("\n"));
}