/**************************************************************
 * SAFE SHIP CONTACT CHECKER - RC_TodaySync.gs v1.2
 * 
 * Quick sync for TODAY's data only (no yesterday rewrite)
 * With Premium Sidebar Progress UI!
 * 
 * FEATURES:
 * - Syncs today's SMS (all texts - inbound + outbound)
 * - Syncs today's Calls (outbound only)
 * - Preserves yesterday's cached data
 * - Fast execution (~30-60 seconds)
 * - Premium progress sidebar with live updates
 * 
 * MENU ITEM: RingCentral API -> Sync Today Only
 * 
 * NOTE: getTodaySyncSidebarHtml_() is now in RC_QuickSync.gs
 *       Do NOT duplicate it here.
 **************************************************************/

/**
 * SYNC TODAY ONLY - Menu entry point
 * Shows sidebar and user clicks Start to begin
 */
function RC_API_syncTodayOnly() {
  var ui = SpreadsheetApp.getUi();
  
  // Check authorization first
  try {
    var service = getRingCentralService_();
    if (!service.hasAccess()) {
      ui.alert("[!] Not Authorized", "Please run 'Authorize RingCentral' first.\n\nGo to: RingCentral API -> Authorize", ui.ButtonSet.OK);
      return;
    }
  } catch (e) {
    ui.alert("[X] Error", "Failed to check authorization: " + e.message, ui.ButtonSet.OK);
    return;
  }
  
  // Clear any stale progress
  RC_clearProgress_();
  
  // Show sidebar (uses getTodaySyncSidebarHtml_() from RC_QuickSync.gs)
  var html = HtmlService.createHtmlOutput(getTodaySyncSidebarHtml_())
    .setTitle("Today Sync")
    .setWidth(380);
  
  ui.showSidebar(html);
}

/**
 * START TODAY SYNC - Called from sidebar when user clicks Start
 * This initializes progress FIRST, then calls sync functions
 */
function RC_startTodaySync() {
  var startTime = new Date().getTime();
  var results = { sms: 0, calls: 0 };
  
  // 
  // CRITICAL: Initialize progress FIRST before any sync calls
  // 
  RC_setProgress_({
    percent: 0,
    currentStage: "Initializing...",
    status: "Starting today-only sync...",
    stage1: { state: "waiting", status: "Waiting..." },
    stage2: { state: "waiting", status: "Waiting..." },
    log: [{ time: getTimeStr_(), msg: "Today-Only Sync started", type: "info" }],
    complete: false,
    startTime: startTime
  });
  
  Utilities.sleep(500);
  
  try {
    // 
    // STAGE 1: Today's SMS (all texts)
    // 
    SS_toast_("Stage 1/2: Today's SMS...", "Today Sync", 60);
    
    var progress = RC_getProgress();
    progress.stage1 = { state: "active", status: "Creating export..." };
    progress.currentStage = "Today's SMS";
    progress.status = "Creating export task...";
    progress.percent = 10;
    RC_setProgress_(progress);
    addLog_("[SMS] Stage 1: Today's SMS", "info");
    
    // Call the existing sync function (it will update progress internally)
    results.sms = RC_API_syncSMS_ToSheet_WithProgress_(0, RC_SMART_CONFIG.SMS_TODAY, "TODAY", 1);
    
    // Mark stage 1 complete
    progress = RC_getProgress();
    progress.stage1 = { state: "complete", status: "Synced!", count: results.sms };
    progress.percent = 50;
    progress.currentStage = "Today's SMS Complete";
    progress.status = results.sms + " messages synced";
    RC_setProgress_(progress);
    addLog_("[OK] Synced: " + results.sms + " SMS", "success");
    
    Utilities.sleep(300);
    
    // 
    // STAGE 2: Today's Calls (Outbound Only)
    // 
    SS_toast_("Stage 2/2: Today's Calls...", "Today Sync", 60);
    
    progress = RC_getProgress();
    progress.stage2 = { state: "active", status: "Fetching..." };
    progress.currentStage = "Today's Calls";
    progress.status = "Fetching outbound calls...";
    progress.percent = 55;
    RC_setProgress_(progress);
    addLog_("[PHONE] Stage 2: Today's Outbound Calls", "info");
    
    // Call the existing sync function
    results.calls = RC_API_syncCalls_ToSheet_WithProgress_(0, RC_SMART_CONFIG.CALL_TODAY, "TODAY", 2);
    
    // Mark stage 2 complete
    progress = RC_getProgress();
    progress.stage2 = { state: "complete", status: "Synced!", count: results.calls };
    progress.percent = 95;
    progress.currentStage = "Today's Calls Complete";
    progress.status = results.calls + " calls synced";
    RC_setProgress_(progress);
    addLog_("[OK] Synced: " + results.calls + " calls", "success");
    
    // 
    // COMPLETE!
    // 
    var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    var total = results.sms + results.calls;
    
    progress = RC_getProgress();
    progress.percent = 100;
    progress.currentStage = "Complete!";
    progress.status = "Today's data synced successfully!";
    progress.complete = true;
    progress.stats = { sms: results.sms, calls: results.calls, total: total, time: elapsed };
    progress.log.push({ time: getTimeStr_(), msg: "COMPLETE! " + total + " records in " + elapsed + "s", type: "success" });
    RC_setProgress_(progress);
    
    SS_toast_("[OK] Synced " + total + " records in " + elapsed + "s!", "Complete", 10);
    
    return { success: true, total: total, sms: results.sms, calls: results.calls, time: elapsed };
    
  } catch (e) {
    Logger.log("Today sync error: " + e);
    var progress = RC_getProgress() || {
      percent: 0,
      stage1: { state: "waiting", status: "Waiting..." },
      stage2: { state: "waiting", status: "Waiting..." },
      log: []
    };
    progress.percent = 0;
    progress.currentStage = "Error";
    progress.status = e.message;
    progress.complete = true;
    progress.error = true;
    progress.log = progress.log || [];
    progress.log.push({ time: getTimeStr_(), msg: "[X] ERROR: " + e.message, type: "error" });
    RC_setProgress_(progress);
    
    return { success: false, error: e.message };
  }
}


/**
 * [SMS] SYNC TODAY'S SMS ONLY - Menu entry point
 */
function RC_API_syncTodaySMSOnly() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var service = getRingCentralService_();
    if (!service.hasAccess()) {
      ui.alert("[!] Not Authorized", "Please authorize RingCentral first.", ui.ButtonSet.OK);
      return;
    }
  } catch (e) {
    ui.alert("[X] Error", e.message, ui.ButtonSet.OK);
    return;
  }
  
  RC_clearProgress_();
  
  var html = HtmlService.createHtmlOutput(getSMSOnlySidebarHtml_())
    .setTitle("[SMS] SMS Sync")
    .setWidth(380);
  
  ui.showSidebar(html);
}

/**
 * Start SMS-only sync (called from sidebar)
 */
function RC_startSMSOnlySync() {
  var startTime = new Date().getTime();
  
  // Initialize progress FIRST
  RC_setProgress_({
    percent: 0,
    currentStage: "Initializing...",
    status: "Starting SMS sync...",
    stage1: { state: "waiting", status: "Waiting..." },
    log: [{ time: getTimeStr_(), msg: "[SMS] SMS-Only Sync started", type: "info" }],
    complete: false,
    startTime: startTime
  });
  
  Utilities.sleep(500);
  
  try {
    var progress = RC_getProgress();
    progress.stage1 = { state: "active", status: "Creating export..." };
    progress.currentStage = "Today's SMS";
    progress.status = "Creating export task...";
    progress.percent = 10;
    RC_setProgress_(progress);
    addLog_("[SEND] Creating SMS export task...", "info");
    
    var count = RC_API_syncSMS_ToSheet_WithProgress_(0, RC_SMART_CONFIG.SMS_TODAY, "TODAY", 1);
    
    var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    
    progress = RC_getProgress();
    progress.percent = 100;
    progress.currentStage = "Complete!";
    progress.status = count + " SMS messages synced!";
    progress.stage1 = { state: "complete", status: "Synced!", count: count };
    progress.complete = true;
    progress.stats = { sms: count, calls: 0, total: count, time: elapsed };
    progress.log.push({ time: getTimeStr_(), msg: "COMPLETE! " + count + " SMS in " + elapsed + "s", type: "success" });
    RC_setProgress_(progress);
    
    return { success: true, total: count, sms: count, calls: 0, time: elapsed };
    
  } catch (e) {
    var progress = RC_getProgress() || { log: [] };
    progress.percent = 0;
    progress.currentStage = "Error";
    progress.status = e.message;
    progress.complete = true;
    progress.error = true;
    progress.log = progress.log || [];
    progress.log.push({ time: getTimeStr_(), msg: "[X] ERROR: " + e.message, type: "error" });
    RC_setProgress_(progress);
    return { success: false, error: e.message };
  }
}


/**
 * [PHONE] SYNC TODAY'S CALLS ONLY - Menu entry point
 */
function RC_API_syncTodayCallsOnly() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var service = getRingCentralService_();
    if (!service.hasAccess()) {
      ui.alert("[!] Not Authorized", "Please authorize RingCentral first.", ui.ButtonSet.OK);
      return;
    }
  } catch (e) {
    ui.alert("[X] Error", e.message, ui.ButtonSet.OK);
    return;
  }
  
  RC_clearProgress_();
  
  var html = HtmlService.createHtmlOutput(getCallsOnlySidebarHtml_())
    .setTitle("[PHONE] Call Sync")
    .setWidth(380);
  
  ui.showSidebar(html);
}

/**
 * Start Calls-only sync (called from sidebar)
 */
function RC_startCallsOnlySync() {
  var startTime = new Date().getTime();
  
  // Initialize progress FIRST
  RC_setProgress_({
    percent: 0,
    currentStage: "Initializing...",
    status: "Starting calls sync...",
    stage1: { state: "waiting", status: "Waiting..." },
    log: [{ time: getTimeStr_(), msg: "[PHONE] Calls-Only Sync started", type: "info" }],
    complete: false,
    startTime: startTime
  });
  
  Utilities.sleep(500);
  
  try {
    var progress = RC_getProgress();
    progress.stage1 = { state: "active", status: "Fetching..." };
    progress.currentStage = "Today's Calls";
    progress.status = "Fetching outbound calls...";
    progress.percent = 10;
    RC_setProgress_(progress);
    addLog_("[PHONE] Fetching outbound calls...", "info");
    
    var count = RC_API_syncCalls_ToSheet_WithProgress_(0, RC_SMART_CONFIG.CALL_TODAY, "TODAY", 1);
    
    var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    
    progress = RC_getProgress();
    progress.percent = 100;
    progress.currentStage = "Complete!";
    progress.status = count + " outbound calls synced!";
    progress.stage1 = { state: "complete", status: "Synced!", count: count };
    progress.complete = true;
    progress.stats = { sms: 0, calls: count, total: count, time: elapsed };
    progress.log.push({ time: getTimeStr_(), msg: "COMPLETE! " + count + " calls in " + elapsed + "s", type: "success" });
    RC_setProgress_(progress);
    
    return { success: true, total: count, sms: 0, calls: count, time: elapsed };
    
  } catch (e) {
    var progress = RC_getProgress() || { log: [] };
    progress.percent = 0;
    progress.currentStage = "Error";
    progress.status = e.message;
    progress.complete = true;
    progress.error = true;
    progress.log = progress.log || [];
    progress.log.push({ time: getTimeStr_(), msg: "[X] ERROR: " + e.message, type: "error" });
    RC_setProgress_(progress);
    return { success: false, error: e.message };
  }
}


/* 
 * SIDEBAR HTML TEMPLATES
 * ===============================================================
 * NOTE: getTodaySyncSidebarHtml_() has been REMOVED from this file.
 * It is now defined ONLY in RC_QuickSync.gs to avoid duplication.
 * =============================================================== */


/* 
 * [SMS] SMS ONLY SIDEBAR
 * =============================================================== */

function getSMSOnlySidebarHtml_() {
  var brand = typeof UNIFIED_SIDEBAR_BRAND !== 'undefined' ? UNIFIED_SIDEBAR_BRAND : {
    PRIMARY: "#8B1538",
    GOLD: "#D4AF37",
    WARNING: "#f59e0b"
  };
  
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>${typeof getUnifiedSidebarCSS_ === 'function' ? getUnifiedSidebarCSS_() : getDefaultSidebarCSS_()}</style>
</head>
<body>
  <div class="header">
    <div class="brand">Safe Ship Contact Checker</div>
    <div class="title">[SMS] SMS Sync</div>
    <div class="subtitle">Sync today's SMS messages</div>
  </div>
  
  <div class="content">
    <div style="text-align: center;">
      <span class="mode-badge" style="background:${brand.WARNING};color:#000">[!] SMS ONLY</span>
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
    
    <div class="section-title">[LIST] Pipeline</div>
    <div class="pipeline">
      <div class="stage" id="st1"><div class="stage-icon">[MSG]</div><div class="stage-info"><div class="stage-name-text">Today's SMS</div><div class="stage-detail" id="st1s">Waiting...</div></div><div class="stage-count" id="st1c"></div></div>
    </div>
    
    <div class="btn-group">
      <button class="btn btn-live" id="startBtn">[MSG] Start SMS Sync</button>
      <button class="btn btn-close" id="closeBtn">Close</button>
    </div>
    
    <div class="section-title">[LOG] Activity Log</div>
    <div class="log-section" id="logSection">
      <div class="log-entry"><span class="log-time">--:--:--</span><span class="log-msg">Ready to sync...</span></div>
    </div>
  </div>
  
  <div class="footer">Safe Ship Contact Checker - SMS Sync v2.0</div>
  
  <script>
    var polling = null, circ = 2 * Math.PI * 55, ring = document.getElementById("ringFill");
    ring.style.strokeDasharray = circ; ring.style.strokeDashoffset = circ;
    
    document.getElementById("startBtn").onclick = function() {
      this.disabled = true; this.textContent = "Syncing...";
      google.script.run.withSuccessHandler(function(r){}).withFailureHandler(function(e){alert("Error: "+e.message)}).RC_startSMSOnlySync();
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
      
      var s = d.stage1;
      if (s) {
        document.getElementById("st1").className = "stage " + (s.state || "");
        document.getElementById("st1s").textContent = s.status || "Waiting...";
        if (s.count != null) { document.getElementById("st1c").textContent = s.count.toLocaleString(); document.getElementById("st1c").style.display = "block"; }
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
      }
    }
  </script>
</body>
</html>`;
}


/* 
 * [PHONE] CALLS ONLY SIDEBAR
 * =============================================================== */

function getCallsOnlySidebarHtml_() {
  var brand = typeof UNIFIED_SIDEBAR_BRAND !== 'undefined' ? UNIFIED_SIDEBAR_BRAND : {
    PRIMARY: "#8B1538",
    GOLD: "#D4AF37",
    WARNING: "#f59e0b"
  };
  
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>${typeof getUnifiedSidebarCSS_ === 'function' ? getUnifiedSidebarCSS_() : getDefaultSidebarCSS_()}</style>
</head>
<body>
  <div class="header">
    <div class="brand">Safe Ship Contact Checker</div>
    <div class="title">[PHONE] Call Sync</div>
    <div class="subtitle">Sync today's outbound calls</div>
  </div>
  
  <div class="content">
    <div style="text-align: center;">
      <span class="mode-badge" style="background:${brand.WARNING};color:#000">[!] CALLS ONLY</span>
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
    
    <div class="section-title">[LIST] Pipeline</div>
    <div class="pipeline">
      <div class="stage" id="st1"><div class="stage-icon">[PHONE]</div><div class="stage-info"><div class="stage-name-text">Today's Calls (Outbound)</div><div class="stage-detail" id="st1s">Waiting...</div></div><div class="stage-count" id="st1c"></div></div>
    </div>
    
    <div class="btn-group">
      <button class="btn btn-live" id="startBtn">[PHONE] Start Call Sync</button>
      <button class="btn btn-close" id="closeBtn">Close</button>
    </div>
    
    <div class="section-title">[LOG] Activity Log</div>
    <div class="log-section" id="logSection">
      <div class="log-entry"><span class="log-time">--:--:--</span><span class="log-msg">Ready to sync...</span></div>
    </div>
  </div>
  
  <div class="footer">Safe Ship Contact Checker - Call Sync v2.0</div>
  
  <script>
    var polling = null, circ = 2 * Math.PI * 55, ring = document.getElementById("ringFill");
    ring.style.strokeDasharray = circ; ring.style.strokeDashoffset = circ;
    
    document.getElementById("startBtn").onclick = function() {
      this.disabled = true; this.textContent = "Syncing...";
      google.script.run.withSuccessHandler(function(r){}).withFailureHandler(function(e){alert("Error: "+e.message)}).RC_startCallsOnlySync();
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
      
      var s = d.stage1;
      if (s) {
        document.getElementById("st1").className = "stage " + (s.state || "");
        document.getElementById("st1s").textContent = s.status || "Waiting...";
        if (s.count != null) { document.getElementById("st1c").textContent = s.count.toLocaleString(); document.getElementById("st1c").style.display = "block"; }
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
      }
    }
  </script>
</body>
</html>`;
}


/* 
 * Default CSS fallback if unified CSS not available
 * =============================================================== */

function getDefaultSidebarCSS_() {
  return `
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh;padding:20px}
    .header{text-align:center;margin-bottom:16px}
    .header .brand{font-size:10px;color:#64748b;text-transform:uppercase;letter-spacing:1px}
    .header .title{font-size:20px;color:#D4AF37;margin:4px 0}
    .header .subtitle{font-size:11px;color:#64748b}
    .content{padding:0}
    .mode-badge{display:inline-block;padding:4px 12px;border-radius:12px;font-size:10px;font-weight:700}
    .ring-section{text-align:center;margin:20px 0}
    .ring-wrap{width:120px;height:120px;margin:0 auto 12px;position:relative}
    .ring-svg{transform:rotate(-90deg)}
    .ring-bg{fill:none;stroke:#334155;stroke-width:6}
    .ring-fill{fill:none;stroke:url(#ringGrad);stroke-width:6;stroke-linecap:round;stroke-dasharray:345.58;stroke-dashoffset:345.58;transition:stroke-dashoffset 0.3s}
    .ring-center{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center}
    .ring-pct{font-size:28px;font-weight:800;color:#fff}
    .ring-label{font-size:10px;color:#64748b}
    .stage-name{font-size:14px;font-weight:600;color:#fff}
    .stage-status{font-size:11px;color:#64748b}
    .section-title{font-size:10px;color:#D4AF37;font-weight:600;margin:16px 0 8px;text-transform:uppercase;letter-spacing:1px}
    .pipeline{background:#1e293b;border-radius:12px;padding:8px}
    .stage{display:flex;align-items:center;padding:10px;border-radius:8px;margin-bottom:4px;background:rgba(255,255,255,0.02);border-left:3px solid #334155}
    .stage:last-child{margin-bottom:0}
    .stage.active{background:rgba(212,175,55,0.15);border-left-color:#D4AF37}
    .stage.complete{background:rgba(16,185,129,0.15);border-left-color:#10b981}
    .stage-icon{font-size:16px;margin-right:10px;width:28px;text-align:center}
    .stage-info{flex:1}
    .stage-name-text{font-size:12px;font-weight:600;color:#fff}
    .stage-detail{font-size:10px;color:#64748b}
    .stage-count{font-size:11px;font-weight:700;color:#10b981;display:none}
    .btn-group{margin:16px 0}
    .btn{width:100%;padding:14px;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;margin-bottom:8px}
    .btn-live{background:linear-gradient(135deg,#D4AF37,#F4D03F);color:#000}
    .btn-live:disabled{background:#334155;color:#64748b;cursor:not-allowed}
    .btn-close{background:#334155;color:#fff;display:none}
    .log-section{background:#1e293b;border-radius:12px;padding:12px;max-height:120px;overflow-y:auto}
    .log-entry{font-size:10px;padding:4px 0;border-bottom:1px solid rgba(255,255,255,0.05);display:flex;gap:8px}
    .log-entry:last-child{border-bottom:none}
    .log-time{color:#475569;font-family:monospace;min-width:55px}
    .log-msg{color:#94a3b8}
    .log-entry.success .log-msg{color:#10b981}
    .log-entry.error .log-msg{color:#ef4444}
    .log-entry.warn .log-msg{color:#f59e0b}
    .footer{text-align:center;font-size:9px;color:#475569;margin-top:16px}
  `;
}