/**************************************************************
 * 🔢 RC_CountVerifier.gs - Outbound Count Verification Tool
 * ══════════════════════════════════════════════════════════════
 * 
 * Verifies that SMS and call counts only include:
 *   - OUTBOUND calls (not inbound)
 *   - OUTBOUND/SENT SMS (not inbound/received)
 * 
 * Menu: RingCentral API → Diagnostics → 🔢 Verify Outbound Counts
 * 
 **************************************************************/


/**
 * 🔢 Verify Outbound Counts
 * Shows detailed breakdown of what's being counted
 */
function RC_verifyOutboundCounts() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];
  
  results.push("═══════════════════════════════════════════════════════════");
  results.push("🔢 OUTBOUND COUNT VERIFICATION REPORT");
  results.push("═══════════════════════════════════════════════════════════");
  results.push("Time: " + new Date().toLocaleString("en-US", { timeZone: "America/New_York" }));
  results.push("");
  
  // ─────────────────────────────────────────────────────────────
  // 1. RC CALL LOG (TODAY) - Verify outbound calls only
  // ─────────────────────────────────────────────────────────────
  results.push("─────────────────────────────────────────────────────────────");
  results.push("📞 RC CALL LOG (TODAY)");
  results.push("─────────────────────────────────────────────────────────────");
  
  var callSheet = ss.getSheetByName("RC CALL LOG");
  if (callSheet) {
    var callLastRow = callSheet.getLastRow();
    if (callLastRow > 1) {
      // Read all data (Direction is column A)
      var callData = callSheet.getRange(2, 1, callLastRow - 1, 10).getValues();
      
      var callStats = {
        total: callData.length,
        outbound: 0,
        inbound: 0,
        other: 0,
        byDirection: {}
      };
      
      for (var i = 0; i < callData.length; i++) {
        var direction = String(callData[i][0] || "").trim().toUpperCase();
        
        // Track all direction values
        callStats.byDirection[direction] = (callStats.byDirection[direction] || 0) + 1;
        
        if (direction === "OUTBOUND") {
          callStats.outbound++;
        } else if (direction === "INBOUND") {
          callStats.inbound++;
        } else {
          callStats.other++;
        }
      }
      
      results.push("   Total rows: " + callStats.total);
      results.push("");
      results.push("   ✅ OUTBOUND (counted): " + callStats.outbound);
      results.push("   ❌ INBOUND (excluded): " + callStats.inbound);
      if (callStats.other > 0) {
        results.push("   ⚠️ OTHER (excluded): " + callStats.other);
      }
      results.push("");
      results.push("   Direction breakdown:");
      for (var dir in callStats.byDirection) {
        var marker = (dir === "OUTBOUND") ? "✅" : "❌";
        results.push("      " + marker + " " + (dir || "(empty)") + ": " + callStats.byDirection[dir]);
      }
      
      // Get timestamp
      var callTimestamp = callSheet.getRange(1, 12).getValue();
      if (callTimestamp) {
        results.push("");
        results.push("   Last sync: " + callTimestamp);
      }
      
    } else {
      results.push("   ⚠️ No data rows (only header)");
    }
  } else {
    results.push("   ❌ Sheet not found!");
  }
  
  results.push("");
  
  // ─────────────────────────────────────────────────────────────
  // 2. RC CALL LOG (YESTERDAY)
  // ─────────────────────────────────────────────────────────────
  results.push("─────────────────────────────────────────────────────────────");
  results.push("📞 RC CALL LOG YESTERDAY");
  results.push("─────────────────────────────────────────────────────────────");
  
  var callYestSheet = ss.getSheetByName("RC CALL LOG YESTERDAY");
  if (callYestSheet) {
    var callYestLastRow = callYestSheet.getLastRow();
    if (callYestLastRow > 1) {
      var callYestData = callYestSheet.getRange(2, 1, callYestLastRow - 1, 1).getValues();
      
      var outbound = 0, inbound = 0;
      for (var i = 0; i < callYestData.length; i++) {
        var dir = String(callYestData[i][0] || "").trim().toUpperCase();
        if (dir === "OUTBOUND") outbound++;
        else if (dir === "INBOUND") inbound++;
      }
      
      results.push("   Total rows: " + callYestData.length);
      results.push("   ✅ OUTBOUND (counted): " + outbound);
      results.push("   ❌ INBOUND (excluded): " + inbound);
    } else {
      results.push("   ⚠️ No data rows");
    }
  } else {
    results.push("   ⚠️ Sheet not found (may not be synced yet)");
  }
  
  results.push("");
  
  // ─────────────────────────────────────────────────────────────
  // 3. RC SMS LOG (TODAY) - Verify outbound SMS only
  // ─────────────────────────────────────────────────────────────
  results.push("─────────────────────────────────────────────────────────────");
  results.push("📱 RC SMS LOG (TODAY)");
  results.push("─────────────────────────────────────────────────────────────");
  
  var smsSheet = ss.getSheetByName("RC SMS LOG");
  if (smsSheet) {
    var smsLastRow = smsSheet.getLastRow();
    if (smsLastRow > 1) {
      // Read Direction (A) and Message Status (J)
      var smsData = smsSheet.getRange(2, 1, smsLastRow - 1, 10).getValues();
      
      var smsStats = {
        total: smsData.length,
        outbound: 0,
        inbound: 0,
        other: 0,
        byDirection: {},
        byStatus: {},
        outboundByStatus: {}
      };
      
      for (var i = 0; i < smsData.length; i++) {
        var direction = String(smsData[i][0] || "").trim().toUpperCase();
        var status = String(smsData[i][9] || "").trim();
        
        // Track all direction values
        smsStats.byDirection[direction] = (smsStats.byDirection[direction] || 0) + 1;
        
        if (direction === "OUTBOUND") {
          smsStats.outbound++;
          // Track outbound by status
          smsStats.outboundByStatus[status || "(no status)"] = (smsStats.outboundByStatus[status || "(no status)"] || 0) + 1;
        } else if (direction === "INBOUND") {
          smsStats.inbound++;
        } else {
          smsStats.other++;
        }
        
        // Track all statuses
        smsStats.byStatus[status || "(no status)"] = (smsStats.byStatus[status || "(no status)"] || 0) + 1;
      }
      
      results.push("   Total rows: " + smsStats.total);
      results.push("");
      results.push("   ✅ OUTBOUND (counted): " + smsStats.outbound);
      results.push("   ❌ INBOUND (excluded): " + smsStats.inbound);
      if (smsStats.other > 0) {
        results.push("   ⚠️ OTHER (excluded): " + smsStats.other);
      }
      results.push("");
      results.push("   Direction breakdown:");
      for (var dir in smsStats.byDirection) {
        var marker = (dir === "OUTBOUND") ? "✅" : "❌";
        results.push("      " + marker + " " + (dir || "(empty)") + ": " + smsStats.byDirection[dir]);
      }
      
      results.push("");
      results.push("   Outbound SMS by status:");
      for (var st in smsStats.outboundByStatus) {
        var statusMarker = (st.toUpperCase().indexOf("DELIVER") >= 0 || st.toUpperCase() === "SENT") ? "✅" : 
                           (st.toUpperCase().indexOf("FAIL") >= 0 || st.toUpperCase().indexOf("ERROR") >= 0) ? "⚠️" : "📤";
        results.push("      " + statusMarker + " " + st + ": " + smsStats.outboundByStatus[st]);
      }
      
      // Get timestamp
      var smsTimestamp = smsSheet.getRange(1, 17).getValue();
      if (smsTimestamp) {
        results.push("");
        results.push("   Last sync: " + smsTimestamp);
      }
      
    } else {
      results.push("   ⚠️ No data rows (only header)");
    }
  } else {
    results.push("   ❌ Sheet not found!");
  }
  
  results.push("");
  
  // ─────────────────────────────────────────────────────────────
  // 4. RC SMS LOG (YESTERDAY)
  // ─────────────────────────────────────────────────────────────
  results.push("─────────────────────────────────────────────────────────────");
  results.push("📱 RC SMS LOG YESTERDAY");
  results.push("─────────────────────────────────────────────────────────────");
  
  var smsYestSheet = ss.getSheetByName("RC SMS LOG YESTERDAY");
  if (smsYestSheet) {
    var smsYestLastRow = smsYestSheet.getLastRow();
    if (smsYestLastRow > 1) {
      var smsYestData = smsYestSheet.getRange(2, 1, smsYestLastRow - 1, 1).getValues();
      
      var outbound = 0, inbound = 0;
      for (var i = 0; i < smsYestData.length; i++) {
        var dir = String(smsYestData[i][0] || "").trim().toUpperCase();
        if (dir === "OUTBOUND") outbound++;
        else if (dir === "INBOUND") inbound++;
      }
      
      results.push("   Total rows: " + smsYestData.length);
      results.push("   ✅ OUTBOUND (counted): " + outbound);
      results.push("   ❌ INBOUND (excluded): " + inbound);
    } else {
      results.push("   ⚠️ No data rows");
    }
  } else {
    results.push("   ⚠️ Sheet not found (may not be synced yet)");
  }
  
  results.push("");
  
  // ─────────────────────────────────────────────────────────────
  // 5. SUMMARY
  // ─────────────────────────────────────────────────────────────
  results.push("═══════════════════════════════════════════════════════════");
  results.push("📊 SUMMARY - WHAT TRACKERS SEE");
  results.push("═══════════════════════════════════════════════════════════");
  results.push("");
  results.push("The tracker builder ONLY counts:");
  results.push("   • Calls where Direction = 'Outbound'");
  results.push("   • SMS where Direction = 'Outbound'");
  results.push("");
  results.push("Inbound calls/SMS are EXCLUDED from contact counts.");
  results.push("");
  results.push("═══════════════════════════════════════════════════════════");
  results.push("END OF VERIFICATION REPORT");
  results.push("═══════════════════════════════════════════════════════════");
  
  // Show results in dialog
  var output = results.join("
");
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<pre style="font-family: 'Consolas', monospace; font-size: 11px; background: #1e293b; color: #e2e8f0; padding: 16px; border-radius: 8px; overflow: auto; max-height: 550px; line-height: 1.4;">' + 
    output.replace(/✅/g, '<span style="color:#10b981">✅</span>')
          .replace(/❌/g, '<span style="color:#ef4444">❌</span>')
          .replace(/⚠️/g, '<span style="color:#f59e0b">⚠️</span>')
          .replace(/📞/g, '<span style="color:#60a5fa">📞</span>')
          .replace(/📱/g, '<span style="color:#a78bfa">📱</span>')
          .replace(/📤/g, '<span style="color:#60a5fa">📤</span>')
          .replace(/📊/g, '<span style="color:#f472b6">📊</span>') +
    '</pre>'
  )
  .setWidth(650)
  .setHeight(600);
  
  ui.showModalDialog(htmlOutput, "🔢 Outbound Count Verification");
}


/**
 * 🔍 Quick Outbound Count Check (Toast)
 * Fast check that shows counts in a toast notification
 */
function RC_quickOutboundCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var countOutbound = function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { total: 0, outbound: 0 };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { total: 0, outbound: 0 };
    
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var outbound = 0;
    
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toUpperCase() === "OUTBOUND") {
        outbound++;
      }
    }
    
    return { total: data.length, outbound: outbound };
  };
  
  var callToday = countOutbound("RC CALL LOG");
  var callYest = countOutbound("RC CALL LOG YESTERDAY");
  var smsToday = countOutbound("RC SMS LOG");
  var smsYest = countOutbound("RC SMS LOG YESTERDAY");
  
  var msg = "📞 Calls Today: " + callToday.outbound + " outbound / " + callToday.total + " total
" +
            "📞 Calls Yesterday: " + callYest.outbound + " outbound / " + callYest.total + " total
" +
            "📱 SMS Today: " + smsToday.outbound + " outbound / " + smsToday.total + " total
" +
            "📱 SMS Yesterday: " + smsYest.outbound + " outbound / " + smsYest.total + " total";
  
  ss.toast(msg, "🔢 Outbound Counts", 10);
}

