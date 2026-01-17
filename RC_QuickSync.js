/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER — RC_QuickSync.gs v3.1
 * ══════════════════════════════════════════════════════════════
 * 
 * v3.1 FIXES (from v3.0 rate limiting issues):
 * ────────────────────────────────────────────────────────────
 * ⚡ REDUCED BATCH SIZE - 4 pages at a time (vs 10)
 * ⚡ LONGER BATCH DELAYS - 1.5s between batches
 * ⚡ SMART BACKOFF - Exponential backoff on rate limit
 * ⚡ FIXED PROGRESS % - No more 425% bug
 * ⚡ HYBRID APPROACH - Serial first page, then parallel
 * 
 * BALANCE: Fast enough (~90-120s) without hitting rate limits
 * 
 **************************************************************/


/* ═══════════════════════════════════════════════════════════════
 * ⚡ QUICK SYNC v3.1 - Menu Entry Point
 * ═══════════════════════════════════════════════════════════════ */

function RC_API_quickSync() {
  var service = getRingCentralService_();
  if (!service.hasAccess()) {
    SpreadsheetApp.getUi().alert("⚠️ Not Authorized", "Please run 'Authorize RingCentral' first.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  RC_clearProgress_();
  var html = HtmlService.createHtmlOutput(getQuickSyncSidebarHtml_v3_())
    .setTitle("⚡ Quick Sync v3.1")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}


/* ═══════════════════════════════════════════════════════════════
 * ⚡ QUICK SYNC ENGINE v3.1 - Called from sidebar
 * ═══════════════════════════════════════════════════════════════ */

function RC_startQuickSync() {
  var startTime = new Date().getTime();
  var results = { calls: 0, sms: 0, callsNew: 0, smsNew: 0 };
  
  RC_setProgress_({
    percent: 0,
    stage: "Starting Quick Sync v3.1...",
    status: "Initializing...",
    callStatus: "waiting",
    smsStatus: "waiting",
    log: [{ time: getTimeStr_(), msg: "⚡ Quick Sync v3.1 - Balanced Mode", type: "info" }],
    complete: false
  });
  
  try {
    var range = RC_getESTDayRange_(0); // Today
    
    // ═══════════════════════════════════════════════════════════
    // STAGE 1: Today's Calls - BALANCED PARALLEL FETCH
    // ═══════════════════════════════════════════════════════════
    RC_updateQuickProgress_(5, "Syncing Calls...", "Starting fetch...", "active", "waiting");
    addLog_("📞 Fetching calls: " + range.displayDate, "info");
    
    var callResult = RC_quickSyncCalls_Balanced_(range);
    results.calls = callResult.total;
    results.callsNew = callResult.newCount;
    
    RC_updateQuickProgress_(45, "Calls done!", callResult.newCount + " new / " + callResult.total + " total", "complete", "waiting");
    addLog_("✅ Calls: " + callResult.newCount + " new in " + callResult.time + "s", "success");
    
    // ═══════════════════════════════════════════════════════════
    // STAGE 2: Today's SMS
    // ═══════════════════════════════════════════════════════════
    RC_updateQuickProgress_(50, "Syncing SMS...", "Creating export task...", "complete", "active");
    addLog_("💬 SMS export: " + range.displayDate, "info");
    
    var smsResult = RC_quickSyncSMS_(range);
    results.sms = smsResult.total;
    results.smsNew = smsResult.newCount;
    
    RC_updateQuickProgress_(95, "SMS done!", smsResult.newCount + " new / " + smsResult.total + " total", "complete", "complete");
    addLog_("✅ SMS: " + smsResult.newCount + " new", "success");
    
    // ═══════════════════════════════════════════════════════════
    // COMPLETE
    // ═══════════════════════════════════════════════════════════
    var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    var totalNew = results.callsNew + results.smsNew;
    var totalFetched = results.calls + results.sms;
    
    var progress = RC_getProgress();
    progress.percent = 100;
    progress.stage = "Complete!";
    progress.status = totalNew + " new records in " + elapsed + "s";
    progress.complete = true;
    progress.stats = { 
      calls: results.calls, 
      callsNew: results.callsNew,
      sms: results.sms, 
      smsNew: results.smsNew,
      total: totalFetched, 
      totalNew: totalNew,
      time: elapsed
    };
    progress.log.push({ time: getTimeStr_(), msg: "🎉 Done! " + totalNew + " new in " + elapsed + "s", type: "success" });
    RC_setProgress_(progress);
    
    return { success: true, total: totalFetched, newCount: totalNew, calls: results.calls, sms: results.sms };
    
  } catch (e) {
    Logger.log("Quick sync error: " + e);
    var progress = RC_getProgress() || { log: [] };
    progress.percent = 0;
    progress.stage = "Error";
    progress.status = e.message;
    progress.complete = true;
    progress.error = true;
    progress.log = progress.log || [];
    progress.log.push({ time: getTimeStr_(), msg: "❌ ERROR: " + e.message, type: "error" });
    RC_setProgress_(progress);
    return { success: false, error: e.message };
  }
}


/* ═══════════════════════════════════════════════════════════════
 * 📞 BALANCED CALL FETCH - v3.1
 * ──────────────────────────────────────────────────────────────
 * - 4 pages per batch (vs 10 in v3.0)
 * - 1.5 second delay between batches
 * - Exponential backoff on rate limit
 * - 1000 records per page
 * ═══════════════════════════════════════════════════════════════ */

function RC_quickSyncCalls_Balanced_(range) {
  var callStartTime = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RC_SMART_CONFIG.CALL_TODAY);
  if (!sheet) sheet = ss.insertSheet(RC_SMART_CONFIG.CALL_TODAY);
  
  var service = getRingCentralService_();
  var accessToken = service.getAccessToken();
  var baseUrl = RC_API_CONFIG.API_BASE;
  
  // STEP 1: Get existing call IDs for deduplication
  var existingIds = RC_getExistingRecordIds_(sheet, "call");
  addLog_("📊 Existing calls: " + Object.keys(existingIds).length, "info");
  
  var allCalls = [];
  var pageSize = 1000;   // Large pages
  var batchSize = 4;     // v3.1: Reduced from 10 to 4
  var batchDelay = 1500; // v3.1: 1.5 seconds between batches
  var maxPages = 50;
  var currentPage = 1;
  var hasMore = true;
  var rateLimitRetries = 0;
  var maxRateLimitRetries = 5;
  
  // STEP 2: First request (serial) to get record count
  addLog_("🔍 Getting record count...", "info");
  var firstUrl = baseUrl + "/account/~/call-log?" + buildCallParams_(range, pageSize, 1);
  
  var firstResponse = UrlFetchApp.fetch(firstUrl, {
    method: "GET",
    headers: { "Authorization": "Bearer " + accessToken },
    muteHttpExceptions: true
  });
  
  if (firstResponse.getResponseCode() === 429) {
    addLog_("⚠️ Rate limited on first request, waiting 60s...", "warn");
    Utilities.sleep(60000);
    firstResponse = UrlFetchApp.fetch(firstUrl, {
      method: "GET",
      headers: { "Authorization": "Bearer " + accessToken },
      muteHttpExceptions: true
    });
  }
  
  if (firstResponse.getResponseCode() !== 200) {
    throw new Error("API Error: " + firstResponse.getResponseCode());
  }
  
  var firstData = JSON.parse(firstResponse.getContentText());
  var firstRecords = firstData.records || [];
  allCalls = allCalls.concat(firstRecords);
  
  // Get paging info
  var paging = firstData.paging || {};
  var totalRecords = paging.totalElements || firstRecords.length;
  var totalPages = Math.ceil(totalRecords / pageSize);
  totalPages = Math.min(totalPages, maxPages);
  
  addLog_("📊 Total: ~" + totalRecords + " calls (" + totalPages + " pages)", "info");
  
  if (firstRecords.length < pageSize) {
    hasMore = false;
  } else {
    currentPage = 2;
  }
  
  // STEP 3: Fetch remaining pages in small parallel batches
  while (hasMore && currentPage <= totalPages) {
    var endPage = Math.min(currentPage + batchSize - 1, totalPages);
    var requests = [];
    
    // Build batch of requests
    for (var p = currentPage; p <= endPage; p++) {
      requests.push({
        url: baseUrl + "/account/~/call-log?" + buildCallParams_(range, pageSize, p),
        method: "GET",
        headers: { "Authorization": "Bearer " + accessToken },
        muteHttpExceptions: true
      });
    }
    
    // Update progress (capped at 40%)
    var progressPct = Math.min(5 + Math.round((currentPage / totalPages) * 35), 40);
    RC_updateQuickProgress_(
      progressPct,
      "Syncing Calls...",
      "Pages " + currentPage + "-" + endPage + " of " + totalPages,
      "active", "waiting"
    );
    
    // Execute batch
    var responses = UrlFetchApp.fetchAll(requests);
    var batchRecordCount = 0;
    var foundEnd = false;
    var gotRateLimited = false;
    
    for (var i = 0; i < responses.length; i++) {
      var respCode = responses[i].getResponseCode();
      
      if (respCode === 200) {
        var data = JSON.parse(responses[i].getContentText());
        var records = data.records || [];
        allCalls = allCalls.concat(records);
        batchRecordCount += records.length;
        
        if (records.length < pageSize) {
          foundEnd = true;
        }
      } else if (respCode === 429) {
        gotRateLimited = true;
      }
    }
    
    // Handle rate limiting with exponential backoff
    if (gotRateLimited) {
      rateLimitRetries++;
      if (rateLimitRetries > maxRateLimitRetries) {
        addLog_("❌ Too many rate limits, aborting", "error");
        break;
      }
      
      var waitTime = Math.min(30 * Math.pow(2, rateLimitRetries - 1), 120); // 30s, 60s, 120s max
      addLog_("⚠️ Rate limited, waiting " + waitTime + "s... (retry " + rateLimitRetries + ")", "warn");
      Utilities.sleep(waitTime * 1000);
      
      // Don't advance page, retry this batch
      continue;
    }
    
    // Success - log and advance
    addLog_("📄 Pages " + currentPage + "-" + endPage + ": +" + batchRecordCount + " (" + allCalls.length + " total)", "info");
    
    if (foundEnd || endPage >= totalPages) {
      hasMore = false;
    } else {
      currentPage = endPage + 1;
      // Delay between batches to avoid rate limiting
      Utilities.sleep(batchDelay);
    }
  }
  
  addLog_("📥 Fetched " + allCalls.length + " calls total", "info");
  
  // STEP 4: Filter to only NEW calls
  var newCalls = [];
  for (var i = 0; i < allCalls.length; i++) {
    var call = allCalls[i];
    var callKey = buildCallKey_(call);
    
    if (callKey && !existingIds[callKey]) {
      newCalls.push(call);
    }
  }
  
  addLog_("🆕 New calls: " + newCalls.length, "info");
  
  // STEP 5: Write to sheet
  if (newCalls.length > 0) {
    var headers = RC_API_CONFIG.CALL_HEADERS;
    var rows = newCalls.map(function(call) { return formatCallRow_(call); });
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground(SAFE_SHIP_BRAND.primaryColor).setFontColor("#FFFFFF").setFontWeight("bold");
      lastRow = 1;
    }
    
    sheet.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
    
    var timestamp = Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "MM/dd/yyyy hh:mm:ss a");
    sheet.getRange(1, headers.length + 2).setValue("Quick Sync v3.1: " + timestamp + " | +" + newCalls.length + " new");
  }
  
  var callElapsed = ((new Date().getTime() - callStartTime) / 1000).toFixed(1);
  return { total: allCalls.length, newCount: newCalls.length, time: callElapsed };
}


/* ═══════════════════════════════════════════════════════════════
 * 📞 HELPER: Build Call API Query Parameters
 * ═══════════════════════════════════════════════════════════════ */

function buildCallParams_(range, pageSize, page) {
  var params = {
    dateFrom: range.dateFrom,
    dateTo: range.dateTo,
    perPage: pageSize,
    page: page,
    view: "Simple"
  };
  
  return Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + "=" + encodeURIComponent(params[key]);
  }).join("&");
}


/* ═══════════════════════════════════════════════════════════════
 * 📞 HELPER: Build Call Deduplication Key
 * ═══════════════════════════════════════════════════════════════ */

function buildCallKey_(call) {
  var callDate = "";
  var callTime = "";
  var callPhone = "";
  
  if (call.startTime) {
    var dt = new Date(call.startTime);
    callDate = Utilities.formatDate(dt, RC_SMART_CONFIG.TZ, "MM/dd/yyyy");
    callTime = Utilities.formatDate(dt, RC_SMART_CONFIG.TZ, "hh:mm:ss a");
  }
  if (call.to && call.to.phoneNumber) {
    callPhone = call.to.phoneNumber.replace(/\D/g, "").slice(-10);
  } else if (call.from && call.from.phoneNumber) {
    callPhone = call.from.phoneNumber.replace(/\D/g, "").slice(-10);
  }
  
  var key = callDate + "_" + callTime + "_" + callPhone;
  return (key !== "__") ? key : null;
}


/* ═══════════════════════════════════════════════════════════════
 * 💬 QUICK SYNC SMS - Message Store Export
 * ═══════════════════════════════════════════════════════════════ */

function RC_quickSyncSMS_(range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RC_SMART_CONFIG.SMS_TODAY);
  if (!sheet) sheet = ss.insertSheet(RC_SMART_CONFIG.SMS_TODAY);
  
  var existingIds = RC_getExistingRecordIds_(sheet, "sms");
  addLog_("📊 Existing SMS: " + Object.keys(existingIds).length, "info");
  
  var allSMS = [];
  
  try {
    allSMS = RC_fetchSMSViaExport_(range);
    addLog_("📥 Export returned: " + allSMS.length + " SMS", "info");
  } catch (e) {
    addLog_("⚠️ Export failed: " + e.message, "warn");
    allSMS = [];
  }
  
  if (allSMS.length === 0) {
    return { total: 0, newCount: 0 };
  }
  
  // Filter to NEW only
  var newSMS = [];
  for (var i = 0; i < allSMS.length; i++) {
    var msg = allSMS[i];
    var msgKey = buildSMSKey_(msg);
    
    if (msgKey && !existingIds[msgKey]) {
      newSMS.push(msg);
    }
  }
  
  addLog_("🆕 New SMS: " + newSMS.length, "info");
  
  if (newSMS.length > 0) {
    var headers = RC_API_CONFIG.SMS_HEADERS;
    var rows = newSMS.map(function(msg) { return formatSMSRow_(msg); });
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground(SAFE_SHIP_BRAND.primaryColor).setFontColor("#FFFFFF").setFontWeight("bold");
      lastRow = 1;
    }
    
    sheet.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
    
    var timestamp = Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "MM/dd/yyyy hh:mm:ss a");
    sheet.getRange(1, headers.length + 2).setValue("Quick Sync v3.1: " + timestamp + " | +" + newSMS.length + " new");
  }
  
  return { total: allSMS.length, newCount: newSMS.length };
}


/* ═══════════════════════════════════════════════════════════════
 * 💬 HELPER: Build SMS Deduplication Key
 * ═══════════════════════════════════════════════════════════════ */

function buildSMSKey_(msg) {
  var msgDateTime = "";
  var msgSender = "";
  
  if (msg.creationTime) {
    var dt = new Date(msg.creationTime);
    msgDateTime = Utilities.formatDate(dt, RC_SMART_CONFIG.TZ, "MM/dd/yyyy hh:mm:ss a");
  }
  if (msg.from && msg.from.phoneNumber) {
    msgSender = msg.from.phoneNumber.replace(/\D/g, "").slice(-10);
  }
  
  var key = msgDateTime + "_" + msgSender;
  return (key !== "_") ? key : null;
}


/* ═══════════════════════════════════════════════════════════════
 * 💬 FETCH SMS VIA EXPORT API
 * ═══════════════════════════════════════════════════════════════ */

function RC_fetchSMSViaExport_(range) {
  var taskId = createMessageStoreReport_(range.dateFrom, range.dateTo);
  addLog_("📤 Export task created", "info");
  
  var maxWait = 90;
  var waited = 0;
  var pollInterval = 2000;
  var status = "";
  
  while (waited < maxWait) {
    Utilities.sleep(pollInterval);
    waited += pollInterval / 1000;
    
    status = checkMessageStoreReportStatus_(taskId);
    
    if (status === "Completed") {
      addLog_("✅ Export ready: " + Math.round(waited) + "s", "success");
      break;
    } else if (status === "Failed" || status === "AttemptFailed") {
      throw new Error("Export failed: " + status);
    }
    
    if (waited > 10) pollInterval = 3000;
    if (waited > 30) pollInterval = 5000;
    
    if (Math.round(waited) % 15 === 0) {
      RC_updateQuickProgress_(50 + Math.round(waited / 3), "Syncing SMS...", "Export: " + status + " (" + Math.round(waited) + "s)", "complete", "active");
    }
  }
  
  if (status !== "Completed") {
    throw new Error("Export timed out");
  }
  
  return downloadAndParseMessageStoreArchive_(taskId);
}


/* ═══════════════════════════════════════════════════════════════
 * 🔍 GET EXISTING RECORD IDs FOR DEDUPLICATION
 * ═══════════════════════════════════════════════════════════════ */

function RC_getExistingRecordIds_(sheet, type) {
  var existingIds = {};
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return existingIds;
  
  try {
    var numRows = lastRow - 1;
    var dataRange = sheet.getRange(2, 1, numRows, 8);
    var data = dataRange.getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var key;
      
      if (type === "call") {
        var date = String(row[4] || "").trim();
        var time = String(row[5] || "").trim();
        var phone = String(row[2] || "").replace(/\D/g, "").slice(-10);
        key = date + "_" + time + "_" + phone;
      } else {
        var dateTime = String(row[7] || "").trim();
        var sender = String(row[3] || "").replace(/\D/g, "").slice(-10);
        key = dateTime + "_" + sender;
      }
      
      if (key && key !== "_" && key !== "__") {
        existingIds[key] = true;
      }
    }
  } catch (e) {
    Logger.log("Error reading existing IDs: " + e);
  }
  
  return existingIds;
}


/* ═══════════════════════════════════════════════════════════════
 * 🔧 HELPER FUNCTIONS
 * ═══════════════════════════════════════════════════════════════ */

function RC_updateQuickProgress_(pct, stage, status, callState, smsState) {
  var progress = RC_getProgress();
  if (!progress) return;
  progress.percent = Math.min(pct, 100);  // v3.1: Cap at 100%
  progress.stage = stage;
  progress.status = status;
  progress.callStatus = callState;
  progress.smsStatus = smsState;
  RC_setProgress_(progress);
}


/* ═══════════════════════════════════════════════════════════════
 * 🎨 QUICK SYNC SIDEBAR HTML v3.1
 * ═══════════════════════════════════════════════════════════════ */

function getQuickSyncSidebarHtml_v3_() {
  var brand = UNIFIED_SIDEBAR_BRAND;
  
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>${getUnifiedSidebarCSS_()}</style>
</head>
<body>
  <div class="header">
    <div class="brand">🚢 Safe Ship Contact Checker</div>
    <div class="title">⚡ Quick Sync</div>
    <div class="subtitle">Balanced parallel + incremental sync</div>
  </div>
  
  <div class="content">
    <div style="text-align: center;">
      <span class="mode-badge live" id="modeBadge">v3.1 BALANCED</span>
    </div>
    
    <div class="ring-section">
      <div class="ring-wrap">
        <svg class="ring-svg" viewBox="0 0 120 120">
          <defs>
            <linearGradient id="ringGrad" x1="0%" y1="0%" x2="100%" y2="0%">
              <stop offset="0%" stop-color="${brand.PRIMARY}"/>
              <stop offset="50%" stop-color="${brand.GOLD}"/>
              <stop offset="100%" stop-color="${brand.PRIMARY}"/>
            </linearGradient>
          </defs>
          <circle class="ring-bg" cx="60" cy="60" r="55"/>
          <circle class="ring-fill" id="ringFill" cx="60" cy="60" r="55"/>
        </svg>
        <div class="ring-center">
          <div class="ring-pct" id="ringPct">0%</div>
          <div class="ring-label">Complete</div>
        </div>
      </div>
      <div class="stage-name" id="mainStatus">Ready</div>
      <div class="stage-status" id="mainSubStatus">Click Start to begin</div>
    </div>
    
    <div class="section-title">📊 Sync Stats</div>
    <div class="stats-grid stats-grid-2" id="statsGrid" style="display:none">
      <div class="stat-box"><div class="stat-val new" id="statNew">0</div><div class="stat-lbl">New Records</div></div>
      <div class="stat-box"><div class="stat-val time" id="statTime">0s</div><div class="stat-lbl">Time</div></div>
    </div>
    
    <div class="section-title">📋 Pipeline</div>
    <div class="pipeline">
      <div class="stage" id="stCall"><div class="stage-icon">📞</div><div class="stage-info"><div class="stage-name-text">Calls (4-page batches)</div><div class="stage-detail" id="stCallStatus">Waiting...</div></div><div class="stage-count" id="stCallCount"></div></div>
      <div class="stage" id="stSMS"><div class="stage-icon">💬</div><div class="stage-info"><div class="stage-name-text">SMS (Export)</div><div class="stage-detail" id="stSMSStatus">Waiting...</div></div><div class="stage-count" id="stSMSCount"></div></div>
    </div>
    
    <div class="btn-group">
      <button class="btn btn-live" id="btnStart">⚡ Start Quick Sync</button>
      <button class="btn btn-close" id="btnClose">Close</button>
    </div>
    
    <div class="section-title">📝 Activity Log</div>
    <div class="log-section" id="logSection">
      <div class="log-entry"><span class="log-time">--:--:--</span><span class="log-msg">Ready to sync...</span></div>
    </div>
  </div>
  
  <div class="footer">Safe Ship Contact Checker • Quick Sync v3.1</div>
  
  <script>
    var poll, circ = 2 * Math.PI * 55, ring = document.getElementById("ringFill");
    ring.style.strokeDasharray = circ; ring.style.strokeDashoffset = circ;
    
    document.getElementById("btnStart").onclick = function() {
      this.disabled = true; this.textContent = "⚡ Syncing...";
      google.script.run.withSuccessHandler(function(r){}).withFailureHandler(function(e){alert(e.message)}).RC_startQuickSync();
      poll = setInterval(function() { google.script.run.withSuccessHandler(updateUI).RC_getProgress(); }, 500);
    };
    
    document.getElementById("btnClose").onclick = function() { google.script.host.close(); };
    
    function updateUI(d) {
      if (!d) return;
      var pct = Math.min(d.percent || 0, 100);
      ring.style.strokeDashoffset = circ - (pct / 100) * circ;
      document.getElementById("ringPct").textContent = pct + "%";
      document.getElementById("mainStatus").textContent = d.stage || "Processing...";
      document.getElementById("mainSubStatus").textContent = d.status || "";
      
      var cs = d.callStatus || "waiting", ss = d.smsStatus || "waiting";
      document.getElementById("stCall").className = "stage " + (cs === "complete" ? "complete" : cs === "active" ? "active" : "");
      document.getElementById("stSMS").className = "stage " + (ss === "complete" ? "complete" : ss === "active" ? "active" : "");
      
      if (d.log && d.log.length) {
        var h = "";
        d.log.forEach(function(l) { h += '<div class="log-entry ' + (l.type || "") + '"><span class="log-time">' + l.time + '</span><span class="log-msg">' + l.msg + '</span></div>'; });
        document.getElementById("logSection").innerHTML = h;
        document.getElementById("logSection").scrollTop = 99999;
      }
      
      if (d.complete) {
        clearInterval(poll);
        document.getElementById("btnStart").style.display = "none";
        document.getElementById("btnClose").style.display = "block";
        if (d.stats) {
          document.getElementById("statsGrid").style.display = "grid";
          document.getElementById("statNew").textContent = d.stats.totalNew || 0;
          document.getElementById("statTime").textContent = (d.stats.time || 0) + "s";
          document.getElementById("stCallStatus").textContent = d.stats.callsNew + " new / " + d.stats.calls + " fetched";
          document.getElementById("stSMSStatus").textContent = d.stats.smsNew + " new / " + d.stats.sms + " fetched";
          document.getElementById("stCallCount").textContent = d.stats.callsNew; document.getElementById("stCallCount").style.display = "block";
          document.getElementById("stSMSCount").textContent = d.stats.smsNew; document.getElementById("stSMSCount").style.display = "block";
        }
      }
    }
  </script>
</body>
</html>`;
}


/* ═══════════════════════════════════════════════════════════════
 * 3. TODAY SYNC SIDEBAR
 * Replace getTodaySyncSidebarHtml_() in RC_TodaySync.gs
 * Location: Line 374
 * ═══════════════════════════════════════════════════════════════ */

function getTodaySyncSidebarHtml_() {
  var brand = UNIFIED_SIDEBAR_BRAND;
  
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>${getUnifiedSidebarCSS_()}</style>
</head>
<body>
  <div class="header">
    <div class="brand">🚢 Safe Ship Contact Checker</div>
    <div class="title">⚡ Today Sync</div>
    <div class="subtitle">Full export of today's data</div>
  </div>
  
  <div class="content">
    <div style="text-align: center;">
      <span class="mode-badge" style="background:${brand.WARNING};color:#000">⚠️ TODAY ONLY</span>
    </div>
    
    <div class="ring-section">
      <div class="ring-wrap">
        <svg class="ring-svg" viewBox="0 0 120 120">
          <defs>
            <linearGradient id="ringGrad" x1="0%" y1="0%" x2="100%" y2="0%">
              <stop offset="0%" stop-color="${brand.PRIMARY}"/>
              <stop offset="50%" stop-color="${brand.GOLD}"/>
              <stop offset="100%" stop-color="${brand.PRIMARY}"/>
            </linearGradient>
          </defs>
          <circle class="ring-bg" cx="60" cy="60" r="55"/>
          <circle class="ring-fill" id="ringFill" cx="60" cy="60" r="55"/>
        </svg>
        <div class="ring-center">
          <div class="ring-pct" id="ringPct">0%</div>
          <div class="ring-label">Complete</div>
        </div>
      </div>
      <div class="stage-name" id="mainStatus">Ready</div>
      <div class="stage-status" id="mainSubStatus">Click Start to begin</div>
    </div>
    
    <div class="section-title">📊 Sync Stats</div>
    <div class="stats-grid stats-grid-2" id="statsGrid" style="display:none">
      <div class="stat-box"><div class="stat-val new" id="statTotal">0</div><div class="stat-lbl">Total Records</div></div>
      <div class="stat-box"><div class="stat-val time" id="statTime">0s</div><div class="stat-lbl">Time</div></div>
    </div>
    
    <div class="section-title">📋 Pipeline</div>
    <div class="pipeline">
      <div class="stage" id="st1"><div class="stage-icon">💬</div><div class="stage-info"><div class="stage-name-text">Today's SMS</div><div class="stage-detail" id="st1s">Waiting...</div></div><div class="stage-count" id="st1c"></div></div>
      <div class="stage" id="st2"><div class="stage-icon">📞</div><div class="stage-info"><div class="stage-name-text">Today's Calls (Outbound)</div><div class="stage-detail" id="st2s">Waiting...</div></div><div class="stage-count" id="st2c"></div></div>
    </div>
    
    <div class="btn-group">
      <button class="btn btn-live" id="startBtn">⚡ Start Today Sync</button>
      <button class="btn btn-close" id="closeBtn">Close</button>
    </div>
    
    <div class="section-title">📝 Activity Log</div>
    <div class="log-section" id="logSection">
      <div class="log-entry"><span class="log-time">--:--:--</span><span class="log-msg">Ready to sync...</span></div>
    </div>
  </div>
  
  <div class="footer">Safe Ship Contact Checker • Today Sync v2.0</div>
  
  <script>
    var polling = null, circ = 2 * Math.PI * 55, ring = document.getElementById("ringFill");
    ring.style.strokeDasharray = circ; ring.style.strokeDashoffset = circ;
    
    document.getElementById("startBtn").onclick = function() {
      this.disabled = true; this.textContent = "Syncing...";
      document.getElementById("mainStatus").textContent = "Starting...";
      document.getElementById("mainSubStatus").textContent = "Connecting to RingCentral...";
      google.script.run.withSuccessHandler(function(r){}).withFailureHandler(function(e){alert("Error: "+e.message)}).RC_startTodaySync();
      polling = setInterval(checkProgress, 700);
    };
    
    document.getElementById("closeBtn").onclick = function() { google.script.host.close(); };
    
    function checkProgress() { google.script.run.withSuccessHandler(updateUI).withFailureHandler(function(e){}).RC_getProgress(); }
    
    function updateUI(d) {
      if (!d) return;
      var pct = d.percent || 0;
      ring.style.strokeDashoffset = circ - (pct / 100) * circ;
      document.getElementById("ringPct").textContent = pct + "%";
      document.getElementById("mainStatus").textContent = d.currentStage || "Processing...";
      document.getElementById("mainSubStatus").textContent = d.status || "";
      
      for (var i = 1; i <= 2; i++) {
        var s = d["stage" + i];
        if (s) {
          document.getElementById("st" + i).className = "stage " + (s.state || "");
          document.getElementById("st" + i + "s").textContent = s.status || "Waiting...";
          if (s.count != null) { document.getElementById("st" + i + "c").textContent = s.count.toLocaleString(); document.getElementById("st" + i + "c").style.display = "block"; }
        }
      }
      
      if (d.log && d.log.length) {
        var html = "";
        d.log.forEach(function(l) { html += '<div class="log-entry ' + (l.type || "") + '"><span class="log-time">' + l.time + '</span><span class="log-msg">' + l.msg + '</span></div>'; });
        document.getElementById("logSection").innerHTML = html;
        document.getElementById("logSection").scrollTop = 99999;
      }
      
      if (d.complete) {
        clearInterval(polling);
        document.getElementById("startBtn").style.display = "none";
        document.getElementById("closeBtn").style.display = "block";
        if (d.stats) { document.getElementById("statsGrid").style.display = "grid"; document.getElementById("statTotal").textContent = (d.stats.total || 0).toLocaleString(); document.getElementById("statTime").textContent = (d.stats.time || 0) + "s"; }
      }
    }
  </script>
</body>
</html>`;
}