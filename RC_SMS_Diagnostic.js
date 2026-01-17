/**************************************************************
 * 🔍 SMS DIAGNOSTIC — Run: RC_debugSMS_SheetStructure()
 **************************************************************/

function RC_debugSMS_SheetStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var lines = [];
  lines.push("🔍 SMS MATCHING DIAGNOSTIC");
  lines.push("═".repeat(60));
  
  // 1. Check RC SMS LOG structure
  lines.push("\n💬 RC SMS LOG STRUCTURE:");
  var smsSheet = ss.getSheetByName("RC SMS LOG");
  if (!smsSheet) {
    lines.push("❌ Sheet 'RC SMS LOG' not found!");
    ui.alert(lines.join("\n"));
    return;
  }
  
  var headers = smsSheet.getRange(1, 1, 1, 15).getValues()[0];
  lines.push("Headers (A-O): " + headers.map(function(h, i) { 
    return String.fromCharCode(65 + i) + "=" + h; 
  }).join(", "));
  
  // 2. Check first few data rows
  lines.push("\n📋 FIRST 3 DATA ROWS:");
  var sampleData = smsSheet.getRange(2, 1, 3, 12).getValues();
  for (var i = 0; i < sampleData.length; i++) {
    var row = sampleData[i];
    lines.push("Row " + (i+2) + ":");
    lines.push("  A (Direction): '" + row[0] + "'");
    lines.push("  D (Sender): '" + row[3] + "' → normalized: '" + RC_normalizePhone_(row[3]) + "'");
    lines.push("  F (Recipient): '" + row[5] + "' → normalized: '" + RC_normalizePhone_(row[5]) + "'");
    lines.push("  H (DateTime): '" + row[7] + "'");
    lines.push("  J (Status): '" + row[9] + "'");
  }
  
  // 3. Count by direction
  lines.push("\n📊 DIRECTION ANALYSIS:");
  var allData = smsSheet.getDataRange().getValues();
  var directions = {};
  for (var i = 1; i < allData.length; i++) {
    var dir = String(allData[i][0] || "").trim();
    directions[dir] = (directions[dir] || 0) + 1;
  }
  for (var d in directions) {
    lines.push("  '" + d + "': " + directions[d] + " rows");
  }
  
  // 4. Check a specific lead phone
  lines.push("\n🔍 LEAD PHONE LOOKUP TEST:");
  var smsTracker = ss.getSheetByName("SMS TRACKER");
  if (smsTracker) {
    // Get first lead phone from SMS TRACKER (column G = index 6)
    var testPhoneRaw = smsTracker.getRange("G3").getValue();
    var testPhoneNorm = RC_normalizePhone_(testPhoneRaw);
    lines.push("Test lead phone (G3): '" + testPhoneRaw + "' → '" + testPhoneNorm + "'");
    
    // Search for this phone in SMS log
    var foundCount = 0;
    var foundRows = [];
    for (var i = 1; i < allData.length && foundRows.length < 3; i++) {
      var row = allData[i];
      var direction = String(row[0] || "").trim();
      var senderNorm = RC_normalizePhone_(row[3]);
      var recipientNorm = RC_normalizePhone_(row[5]);
      
      if (senderNorm === testPhoneNorm || recipientNorm === testPhoneNorm) {
        foundCount++;
        foundRows.push({
          row: i + 1,
          direction: direction,
          sender: senderNorm,
          recipient: recipientNorm,
          datetime: row[7]
        });
      }
    }
    
    lines.push("Found " + foundCount + " total matches for this phone");
    if (foundRows.length > 0) {
      lines.push("Sample matches:");
      foundRows.forEach(function(f) {
        lines.push("  Row " + f.row + ": dir='" + f.direction + "', sender=" + f.sender + ", recip=" + f.recipient);
      });
    }
  }
  
  // 5. Check the current RC_CONFIG values
  lines.push("\n⚙️ CURRENT RC_CONFIG.SMS:");
  lines.push("  DIRECTION: column index " + RC_CONFIG.SMS.DIRECTION + " (should be A=0)");
  lines.push("  SENDER: column index " + RC_CONFIG.SMS.SENDER + " (should be D=3)");
  lines.push("  RECIPIENT: column index " + RC_CONFIG.SMS.RECIPIENT + " (should be F=5)");
  lines.push("  DATETIME: column index " + RC_CONFIG.SMS.DATETIME + " (should be H=7)");
  lines.push("  STATUS: column index " + RC_CONFIG.SMS.STATUS + " (should be J=9)");
  
  ui.alert(lines.join("\n"));
}