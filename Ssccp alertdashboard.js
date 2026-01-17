/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER PRO — ALERT DASHBOARD
 * File: SSCCP_AlertDashboard.gs
 *
 * Comprehensive alert management dashboard with:
 *   - Fancy burgundy/gold Safe Ship branding
 *   - Auto-refresh trackers before sending
 *   - Live lead counter per tracker
 *   - Email preview mode (popup dialog)
 *   - Per-rep progress during sending
 *   - Tracker freshness & parameter display
 *   - Lead/rep breakdown
 *   - Dry run mode
 *   - Rate limit display
 *   - Email logging for recall
 *   - Working cancel button
 *   - Quick parameter edit (dropdowns)
 *   - Jump to tracker links
 *   - **NEW v1.1: Report type selector**
 *   - **NEW v1.1: RC enrichment integration**
 *   - **NEW v1.1: Premium email templates**
 *   - **NEW v1.1: Manager summary reports**
 *
 * v1.1 — Full Pipeline Integration
 **************************************************************/

/* ═══════════════════════════════════════════════════════════════
 * CONFIGURATION
 * ═══════════════════════════════════════════════════════════════ */

var SSCCP_AD_CONFIG = {
  // Tracker sheets to include
  TRACKERS: [
    { key: "SMS", name: "SMS TRACKER", type: "SMS", priorityCell: "B7", countCell: "B6", icon: "📱" },
    { key: "CALL", name: "CALL & VOICEMAIL TRACKER", type: "CALL", priorityCell: "A8", countCell: "B8", icon: "📞" },
    { key: "P1", name: "PRIORITY 1 CALL & VOICEMAIL TRACKER", type: "CALL", priorityCell: "A8", countCell: "B8", icon: "🔥" },
    { key: "P35", name: "PRIORITY 3,5 CALL & VOICEMAIL TRACKER", type: "CALL_MULTI", priorityCell: "A8", countCell: "B8", icon: "⚡" }
  ],
  
  // v1.1: Report types available
  REPORT_TYPES: [
    { 
      key: "UNCONTACTED_SMS_CALL", 
      name: "📉 Uncontacted Leads (SMS + Call/VM)", 
      trackers: ["SMS", "CALL"],
      description: "Leads not meeting SMS or Call/VM contact thresholds"
    },
    { 
      key: "PRIORITY_1", 
      name: "🔥 Priority 1 Call/VM", 
      trackers: ["P1"],
      description: "High-priority leads requiring immediate attention"
    },
    { 
      key: "PRIORITY_35", 
      name: "⚡ Priority 3,5 Call/VM", 
      trackers: ["P35"],
      description: "Medium-priority leads needing follow-up"
    },
    { 
      key: "QUOTED_FOLLOWUP", 
      name: "💰 Quoted Follow-Up", 
      trackers: ["SMS", "CALL"],
      description: "Leads with quotes requiring follow-up contact"
    },
    { 
      key: "ALL_TRACKERS", 
      name: "📊 All Trackers Combined", 
      trackers: ["SMS", "CALL", "P1", "P35"],
      description: "All leads from all trackers"
    }
  ],
  
  // Email quota (Gmail daily limit)
  DAILY_EMAIL_LIMIT: 1500,
  
  // Cache keys
  CACHE_KEYS: {
    RUN_STATE: "SSCCP_AD_RUN_STATE",
    CANCEL_FLAG: "SSCCP_AD_CANCEL",
    LAST_REFRESH: "SSCCP_AD_LAST_REFRESH",
    EMAIL_LOG: "SSCCP_AD_EMAIL_LOG"
  },
  
  // Email log sheet
  EMAIL_LOG_SHEET: "Email_Log",
  
  // Branding colors
  BRAND: {
    PRIMARY: "#8B1538",      // Burgundy
    PRIMARY_DARK: "#6B1028", // Dark burgundy
    GOLD: "#D4AF37",         // Gold accent
    DARK_BG: "#0f172a",      // Dark background
    CARD_BG: "#1e293b",      // Card background
    SUCCESS: "#10b981",      // Green
    WARNING: "#f59e0b",      // Yellow
    ERROR: "#ef4444",        // Red
    MUTED: "#94a3b8"         // Gray text
  }
};

/* ═══════════════════════════════════════════════════════════════
 * MAIN ENTRY POINT - OPEN DASHBOARD
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Open the Alert Dashboard sidebar
 */
function SSCCP_openAlertDashboard() {
  var html = HtmlService.createHtmlOutput(SSCCP_AD_getSidebarHTML_())
    .setTitle("🚢 Alert Dashboard")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Menu entry for Uncontacted Leads Alert (replaces old function)
 */
function SSCCP_MENU_UncontactedLeadsAlert() {
  SSCCP_openAlertDashboard();
}

/* ═══════════════════════════════════════════════════════════════
 * SIDEBAR HTML GENERATOR
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_AD_getSidebarHTML_() {
  var brand = SSCCP_AD_CONFIG.BRAND;
  
  // Build report type options HTML
  var reportOptionsHtml = SSCCP_AD_CONFIG.REPORT_TYPES.map(function(rt) {
    return '<option value="' + rt.key + '">' + rt.name + '</option>';
  }).join('');
  
  return '<!DOCTYPE html>' +
'<html>' +
'<head>' +
'  <base target="_top">' +
'  <style>' +
'    * { box-sizing: border-box; margin: 0; padding: 0; }' +
'    ' +
'    :root {' +
'      --primary: ' + brand.PRIMARY + ';' +
'      --primary-dark: ' + brand.PRIMARY_DARK + ';' +
'      --gold: ' + brand.GOLD + ';' +
'      --dark-bg: ' + brand.DARK_BG + ';' +
'      --card-bg: ' + brand.CARD_BG + ';' +
'      --success: ' + brand.SUCCESS + ';' +
'      --warning: ' + brand.WARNING + ';' +
'      --error: ' + brand.ERROR + ';' +
'      --muted: ' + brand.MUTED + ';' +
'    }' +
'    ' +
'    body {' +
'      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;' +
'      background: var(--dark-bg);' +
'      color: #fff;' +
'      font-size: 12px;' +
'      min-height: 100vh;' +
'    }' +
'    ' +
'    /* Header */' +
'    .header {' +
'      background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%);' +
'      padding: 16px;' +
'      text-align: center;' +
'      border-bottom: 3px solid var(--gold);' +
'    }' +
'    .header .brand {' +
'      font-size: 11px;' +
'      color: var(--gold);' +
'      letter-spacing: 2px;' +
'      text-transform: uppercase;' +
'      font-weight: 700;' +
'    }' +
'    .header .title {' +
'      font-size: 18px;' +
'      font-weight: 800;' +
'      margin-top: 4px;' +
'    }' +
'    .header .subtitle {' +
'      font-size: 10px;' +
'      color: rgba(255,255,255,0.7);' +
'      margin-top: 2px;' +
'    }' +
'    ' +
'    /* Content area */' +
'    .content {' +
'      padding: 12px;' +
'    }' +
'    ' +
'    /* Mode badge */' +
'    .mode-badge {' +
'      display: inline-block;' +
'      padding: 6px 12px;' +
'      border-radius: 20px;' +
'      font-size: 11px;' +
'      font-weight: 700;' +
'      margin-bottom: 12px;' +
'    }' +
'    .mode-badge.live { background: var(--success); }' +
'    .mode-badge.test { background: var(--warning); color: #000; }' +
'    .mode-badge.safe { background: #3b82f6; }' +
'    ' +
'    /* v1.1: Report type selector */' +
'    .report-selector {' +
'      background: var(--card-bg);' +
'      border-radius: 10px;' +
'      padding: 12px;' +
'      margin-bottom: 12px;' +
'      border: 2px solid var(--gold);' +
'    }' +
'    .report-selector label {' +
'      font-size: 10px;' +
'      font-weight: 700;' +
'      color: var(--gold);' +
'      letter-spacing: 1px;' +
'      text-transform: uppercase;' +
'      display: block;' +
'      margin-bottom: 6px;' +
'    }' +
'    .report-selector select {' +
'      width: 100%;' +
'      padding: 10px;' +
'      border-radius: 6px;' +
'      border: 1px solid rgba(255,255,255,0.2);' +
'      background: var(--dark-bg);' +
'      color: #fff;' +
'      font-size: 12px;' +
'      font-weight: 600;' +
'      cursor: pointer;' +
'    }' +
'    .report-selector select:focus {' +
'      outline: none;' +
'      border-color: var(--gold);' +
'    }' +
'    .report-description {' +
'      font-size: 10px;' +
'      color: var(--muted);' +
'      margin-top: 6px;' +
'      padding: 6px;' +
'      background: rgba(0,0,0,0.2);' +
'      border-radius: 4px;' +
'    }' +
'    ' +
'    /* v1.1: Manager report toggle */' +
'    .manager-toggle {' +
'      display: flex;' +
'      align-items: center;' +
'      gap: 8px;' +
'      margin-top: 10px;' +
'      padding: 8px;' +
'      background: rgba(0,0,0,0.2);' +
'      border-radius: 6px;' +
'    }' +
'    .manager-toggle input[type="checkbox"] {' +
'      width: 18px;' +
'      height: 18px;' +
'      cursor: pointer;' +
'    }' +
'    .manager-toggle label {' +
'      font-size: 11px;' +
'      color: #fff;' +
'      cursor: pointer;' +
'      margin: 0;' +
'      letter-spacing: 0;' +
'      text-transform: none;' +
'    }' +
'    ' +
'    /* Ring progress */' +
'    .ring-section {' +
'      text-align: center;' +
'      margin-bottom: 16px;' +
'    }' +
'    .ring-wrap {' +
'      position: relative;' +
'      width: 100px;' +
'      height: 100px;' +
'      margin: 0 auto;' +
'    }' +
'    .ring-svg {' +
'      transform: rotate(-90deg);' +
'      width: 100%;' +
'      height: 100%;' +
'    }' +
'    .ring-bg {' +
'      fill: none;' +
'      stroke: var(--card-bg);' +
'      stroke-width: 8;' +
'    }' +
'    .ring-fill {' +
'      fill: none;' +
'      stroke: url(#ringGrad);' +
'      stroke-width: 8;' +
'      stroke-linecap: round;' +
'      stroke-dasharray: 345.4;' +
'      stroke-dashoffset: 345.4;' +
'      transition: stroke-dashoffset 0.5s ease;' +
'    }' +
'    .ring-center {' +
'      position: absolute;' +
'      top: 50%;' +
'      left: 50%;' +
'      transform: translate(-50%, -50%);' +
'      text-align: center;' +
'    }' +
'    .ring-pct {' +
'      font-size: 24px;' +
'      font-weight: 800;' +
'      color: var(--gold);' +
'    }' +
'    .ring-label {' +
'      font-size: 9px;' +
'      color: var(--muted);' +
'      text-transform: uppercase;' +
'    }' +
'    .stage-name {' +
'      font-size: 14px;' +
'      font-weight: 700;' +
'      margin-top: 8px;' +
'    }' +
'    .stage-status {' +
'      font-size: 11px;' +
'      color: var(--muted);' +
'    }' +
'</style>' +
'</head>' +
'<body>' +
'  <div class="header">' +
'    <div class="brand">🚢 Safe Ship Contact Checker</div>' +
'    <div class="title">📧 Alert Dashboard</div>' +
'    <div class="subtitle">v1.1 — Full Pipeline Integration</div>' +
'  </div>' +
'  ' +
'  <div class="content">' +
'    <!-- Mode Badge -->' +
'    <div style="text-align: center;">' +
'      <span class="mode-badge" id="modeBadge">Loading...</span>' +
'    </div>' +
'    ' +
'    <!-- v1.1: Report Type Selector -->' +
'    <div class="report-selector">' +
'      <label>📋 Report Type</label>' +
'      <select id="reportType" onchange="onReportTypeChange()">' +
'        ' + reportOptionsHtml + '' +
'      </select>' +
'      <div class="report-description" id="reportDescription">Leads not meeting SMS or Call/VM contact thresholds</div>' +
'      <div class="manager-toggle">' +
'        <input type="checkbox" id="sendManagerReports" />' +
'        <label for="sendManagerReports">📊 Also send Manager Summary Reports</label>' +
'      </div>' +
'    </div>' +
'    ' +
'    <!-- Ring Progress -->' +
'    <div class="ring-section">' +
'      <div class="ring-wrap">' +
'        <svg class="ring-svg" viewBox="0 0 120 120">' +
'          <defs>' +
'            <linearGradient id="ringGrad" x1="0%" y1="0%" x2="100%" y2="0%">' +
'              <stop offset="0%" stop-color="' + brand.PRIMARY + '"/>' +
'              <stop offset="50%" stop-color="' + brand.GOLD + '"/>' +
'              <stop offset="100%" stop-color="' + brand.PRIMARY + '"/>' +
'            </linearGradient>' +
'          </defs>' +
'          <circle class="ring-bg" cx="60" cy="60" r="55"/>' +
'          <circle class="ring-fill" id="ringFill" cx="60" cy="60" r="55"/>' +
'        </svg>' +
'        <div class="ring-center">' +
'          <div class="ring-pct" id="ringPct">0%</div>' +
'          <div class="ring-label">Complete</div>' +
'        </div>' +
'      </div>' +
'      <div class="stage-name" id="mainStatus">Ready</div>' +
'      <div class="stage-status" id="mainSubStatus">Select report type and click Refresh</div>' +
'    </div>' +
SSCCP_AD_getSidebarHTML_Part2_();
}

/**
 * Part 2 of sidebar HTML - Tracker cards, stats, buttons, and more styles
 */
function SSCCP_AD_getSidebarHTML_Part2_() {
  return '' +
'    <!-- Tracker Cards -->' +
'    <div class="section-title">📊 Tracker Status</div>' +
'    <div class="tracker-cards" id="trackerCards">' +
'      <!-- Populated by JS -->' +
'    </div>' +
'    ' +
'    <!-- Lead Breakdown -->' +
'    <div class="section-title">📈 Lead Breakdown</div>' +
'    <div class="stats-grid" id="priorityStats">' +
'      <div class="stat-box"><div class="stat-val p0" id="statP0">-</div><div class="stat-lbl">P0</div></div>' +
'      <div class="stat-box"><div class="stat-val p1" id="statP1">-</div><div class="stat-lbl">P1</div></div>' +
'      <div class="stat-box"><div class="stat-val p3" id="statP3">-</div><div class="stat-lbl">P3</div></div>' +
'      <div class="stat-box"><div class="stat-val p5" id="statP5">-</div><div class="stat-lbl">P5</div></div>' +
'    </div>' +
'    ' +
'    <div class="breakdown" id="repBreakdown">' +
'      <div class="breakdown-row">' +
'        <span class="breakdown-label">Total Reps:</span>' +
'        <span class="breakdown-value" id="totalReps">-</span>' +
'      </div>' +
'      <div class="breakdown-row">' +
'        <span class="breakdown-label">Total Leads:</span>' +
'        <span class="breakdown-value" id="totalLeads">-</span>' +
'      </div>' +
'      <div class="breakdown-row">' +
'        <span class="breakdown-label">Avg per Rep:</span>' +
'        <span class="breakdown-value" id="avgPerRep">-</span>' +
'      </div>' +
'    </div>' +
'    ' +
'    <!-- Email Quota -->' +
'    <div class="quota-section">' +
'      <div class="section-title" style="margin-bottom:0;">📧 Email Quota</div>' +
'      <div class="quota-bar-bg">' +
'        <div class="quota-bar-fill" id="quotaBar" style="width: 0%"></div>' +
'      </div>' +
'      <div class="quota-text">' +
'        <span id="quotaUsed">- used</span>' +
'        <span id="quotaRemaining">- remaining</span>' +
'      </div>' +
'    </div>' +
'    ' +
'    <!-- Pipeline Stages -->' +
'    <div class="section-title">📋 Pipeline</div>' +
'    <div class="pipeline" id="pipelineStages">' +
'      <div class="stage" id="stage1">' +
'        <div class="stage-icon">🔑</div>' +
'        <div class="stage-info">' +
'          <div class="stage-name-text">Gate Checks</div>' +
'          <div class="stage-detail" id="stage1Detail">Waiting...</div>' +
'        </div>' +
'      </div>' +
'      <div class="stage" id="stage2">' +
'        <div class="stage-icon">🔄</div>' +
'        <div class="stage-info">' +
'          <div class="stage-name-text">Refresh Trackers</div>' +
'          <div class="stage-detail" id="stage2Detail">Waiting...</div>' +
'        </div>' +
'      </div>' +
'      <div class="stage" id="stage3">' +
'        <div class="stage-icon">🔍</div>' +
'        <div class="stage-info">' +
'          <div class="stage-name-text">Scan Sources</div>' +
'          <div class="stage-detail" id="stage3Detail">Waiting...</div>' +
'        </div>' +
'        <div class="stage-count" id="stage3Count" style="display:none">0</div>' +
'      </div>' +
'      <div class="stage" id="stage4">' +
'        <div class="stage-icon">🔗</div>' +
'        <div class="stage-info">' +
'          <div class="stage-name-text">RC Enrichment</div>' +
'          <div class="stage-detail" id="stage4Detail">Waiting...</div>' +
'        </div>' +
'      </div>' +
'      <div class="stage" id="stage5">' +
'        <div class="stage-icon">📧</div>' +
'        <div class="stage-info">' +
'          <div class="stage-name-text">Send Alerts</div>' +
'          <div class="stage-detail" id="stage5Detail">Waiting...</div>' +
'        </div>' +
'        <div class="stage-count" id="stage5Count" style="display:none">0</div>' +
'      </div>' +
'    </div>' +
'    ' +
'    <!-- Action Buttons -->' +
'    <div class="btn-group">' +
'      <button class="btn btn-refresh" id="btnRefresh" onclick="refreshTrackers()">' +
'        🔄 Refresh Trackers' +
'      </button>' +
'      <div class="btn-row">' +
'        <button class="btn btn-preview" id="btnPreview" onclick="previewEmails()" disabled>' +
'          👁 Preview' +
'        </button>' +
'        <button class="btn btn-dry" id="btnDry" onclick="startDryRun()" disabled>' +
'          🧪 Dry Run' +
'        </button>' +
'      </div>' +
'      <button class="btn btn-live" id="btnLive" onclick="confirmLiveRun()" disabled>' +
'        ▶ Start Live Run' +
'      </button>' +
'      <button class="btn btn-cancel" id="btnCancel" onclick="cancelRun()">' +
'        ⛔ CANCEL' +
'      </button>' +
'    </div>' +
'    ' +
'    <!-- Action Log -->' +
'    <div class="section-title">📜 Action Log</div>' +
'    <div class="log-section" id="logContainer">' +
'      <div class="log-entry info">' +
'        <span class="log-time">--:--:--</span>' +
'        <span class="log-msg">Dashboard ready. Select report type and click Refresh.</span>' +
'      </div>' +
'    </div>' +
'  </div>' +
'  ' +
'  <div class="footer">' +
'    Safe Ship Contact Checker Pro • Alert Dashboard v1.1' +
'  </div>' +
'  ' +
'  <canvas class="confetti" id="confetti"></canvas>' +
'  ' +
SSCCP_AD_getSidebarHTML_Part3_();
}

/**
 * Part 3 of sidebar HTML - More CSS styles
 */
function SSCCP_AD_getSidebarHTML_Part3_() {
  return '' +
'<style>' +
'    /* Tracker cards */' +
'    .section-title {' +
'      font-size: 10px;' +
'      font-weight: 700;' +
'      color: var(--gold);' +
'      letter-spacing: 1px;' +
'      text-transform: uppercase;' +
'      margin-bottom: 8px;' +
'      display: flex;' +
'      align-items: center;' +
'      gap: 6px;' +
'    }' +
'    .tracker-cards {' +
'      display: flex;' +
'      flex-direction: column;' +
'      gap: 8px;' +
'      margin-bottom: 16px;' +
'    }' +
'    .tracker-card {' +
'      background: var(--card-bg);' +
'      border-radius: 10px;' +
'      padding: 10px;' +
'      border-left: 3px solid var(--primary);' +
'    }' +
'    .tracker-card.active { border-left-color: var(--success); }' +
'    .tracker-card.disabled { opacity: 0.5; border-left-color: var(--muted); }' +
'    .tracker-card.selected { border-left-color: var(--gold); box-shadow: 0 0 0 1px var(--gold); }' +
'    .tracker-header {' +
'      display: flex;' +
'      justify-content: space-between;' +
'      align-items: center;' +
'      margin-bottom: 6px;' +
'    }' +
'    .tracker-name {' +
'      font-weight: 700;' +
'      font-size: 11px;' +
'      display: flex;' +
'      align-items: center;' +
'      gap: 4px;' +
'    }' +
'    .tracker-params {' +
'      font-size: 10px;' +
'      color: var(--gold);' +
'    }' +
'    .tracker-stats {' +
'      display: flex;' +
'      justify-content: space-between;' +
'      align-items: center;' +
'      font-size: 10px;' +
'      color: var(--muted);' +
'    }' +
'    .tracker-leads {' +
'      font-weight: 700;' +
'      color: #fff;' +
'    }' +
'    .tracker-actions {' +
'      display: flex;' +
'      gap: 6px;' +
'      margin-top: 6px;' +
'    }' +
'    .tracker-btn {' +
'      background: rgba(255,255,255,0.1);' +
'      border: none;' +
'      color: #fff;' +
'      padding: 4px 8px;' +
'      border-radius: 4px;' +
'      font-size: 9px;' +
'      cursor: pointer;' +
'    }' +
'    .tracker-btn:hover { background: rgba(255,255,255,0.2); }' +
'    ' +
'    /* Breakdown section */' +
'    .breakdown {' +
'      background: var(--card-bg);' +
'      border-radius: 10px;' +
'      padding: 10px;' +
'      margin-bottom: 12px;' +
'    }' +
'    .breakdown-row {' +
'      display: flex;' +
'      justify-content: space-between;' +
'      padding: 4px 0;' +
'      border-bottom: 1px solid rgba(255,255,255,0.05);' +
'    }' +
'    .breakdown-row:last-child { border-bottom: none; }' +
'    .breakdown-label { color: var(--muted); }' +
'    .breakdown-value { font-weight: 700; }' +
'    ' +
'    /* Stats grid */' +
'    .stats-grid {' +
'      display: grid;' +
'      grid-template-columns: repeat(4, 1fr);' +
'      gap: 6px;' +
'      margin-bottom: 12px;' +
'    }' +
'    .stat-box {' +
'      background: var(--card-bg);' +
'      border-radius: 8px;' +
'      padding: 8px 4px;' +
'      text-align: center;' +
'    }' +
'    .stat-val {' +
'      font-size: 16px;' +
'      font-weight: 800;' +
'    }' +
'    .stat-val.p0 { color: #3b82f6; }' +
'    .stat-val.p1 { color: #ef4444; }' +
'    .stat-val.p3 { color: #f59e0b; }' +
'    .stat-val.p5 { color: #8b5cf6; }' +
'    .stat-lbl {' +
'      font-size: 8px;' +
'      color: var(--muted);' +
'      text-transform: uppercase;' +
'    }' +
'    ' +
'    /* Quota bar */' +
'    .quota-section {' +
'      background: var(--card-bg);' +
'      border-radius: 10px;' +
'      padding: 10px;' +
'      margin-bottom: 12px;' +
'    }' +
'    .quota-bar-bg {' +
'      background: var(--dark-bg);' +
'      border-radius: 4px;' +
'      height: 8px;' +
'      overflow: hidden;' +
'      margin-top: 6px;' +
'    }' +
'    .quota-bar-fill {' +
'      height: 100%;' +
'      background: linear-gradient(90deg, var(--success), var(--warning));' +
'      transition: width 0.3s;' +
'    }' +
'    .quota-text {' +
'      display: flex;' +
'      justify-content: space-between;' +
'      font-size: 10px;' +
'      margin-top: 4px;' +
'    }' +
'    ' +
'    /* Pipeline stages */' +
'    .pipeline {' +
'      background: var(--card-bg);' +
'      border-radius: 10px;' +
'      padding: 10px;' +
'      margin-bottom: 12px;' +
'    }' +
'    .stage {' +
'      display: flex;' +
'      align-items: center;' +
'      gap: 8px;' +
'      padding: 6px 0;' +
'      border-bottom: 1px solid rgba(255,255,255,0.05);' +
'    }' +
'    .stage:last-child { border-bottom: none; }' +
'    .stage-icon {' +
'      width: 24px;' +
'      height: 24px;' +
'      border-radius: 50%;' +
'      background: var(--card-bg);' +
'      border: 2px solid var(--muted);' +
'      display: flex;' +
'      align-items: center;' +
'      justify-content: center;' +
'      font-size: 10px;' +
'      flex-shrink: 0;' +
'    }' +
'    .stage.complete .stage-icon { background: var(--success); border-color: var(--success); }' +
'    .stage.active .stage-icon { background: var(--gold); border-color: var(--gold); animation: pulse 1s infinite; }' +
'    .stage.error .stage-icon { background: var(--error); border-color: var(--error); }' +
'    .stage-info { flex: 1; min-width: 0; }' +
'    .stage-name-text { font-size: 11px; font-weight: 600; }' +
'    .stage-detail { font-size: 9px; color: var(--muted); }' +
'    .stage-count {' +
'      font-size: 10px;' +
'      font-weight: 700;' +
'      color: var(--success);' +
'      background: rgba(16,185,129,0.2);' +
'      padding: 2px 6px;' +
'      border-radius: 6px;' +
'    }' +
'    ' +
'    @keyframes pulse {' +
'      0%, 100% { opacity: 1; }' +
'      50% { opacity: 0.5; }' +
'    }' +
'    ' +
'    /* Action buttons */' +
'    .btn-group {' +
'      display: flex;' +
'      flex-direction: column;' +
'      gap: 8px;' +
'      margin-bottom: 12px;' +
'    }' +
'    .btn-row {' +
'      display: flex;' +
'      gap: 8px;' +
'    }' +
'    .btn {' +
'      flex: 1;' +
'      padding: 12px;' +
'      border: none;' +
'      border-radius: 10px;' +
'      font-size: 12px;' +
'      font-weight: 700;' +
'      cursor: pointer;' +
'      transition: all 0.2s;' +
'      display: flex;' +
'      align-items: center;' +
'      justify-content: center;' +
'      gap: 6px;' +
'    }' +
'    .btn:disabled {' +
'      opacity: 0.5;' +
'      cursor: not-allowed;' +
'    }' +
'    .btn-refresh {' +
'      background: var(--card-bg);' +
'      color: #fff;' +
'      border: 1px solid rgba(255,255,255,0.2);' +
'    }' +
'    .btn-refresh:hover:not(:disabled) { background: rgba(255,255,255,0.1); }' +
'    .btn-preview {' +
'      background: #3b82f6;' +
'      color: #fff;' +
'    }' +
'    .btn-preview:hover:not(:disabled) { background: #2563eb; }' +
'    .btn-dry {' +
'      background: var(--warning);' +
'      color: #000;' +
'    }' +
'    .btn-dry:hover:not(:disabled) { background: #d97706; }' +
'    .btn-live {' +
'      background: linear-gradient(135deg, var(--primary), var(--primary-dark));' +
'      color: #fff;' +
'    }' +
'    .btn-live:hover:not(:disabled) { transform: scale(1.02); }' +
'    .btn-cancel {' +
'      background: var(--error);' +
'      color: #fff;' +
'      display: none;' +
'    }' +
'    .btn-cancel:hover:not(:disabled) { background: #dc2626; }' +
'    ' +
'    /* Action log */' +
'    .log-section {' +
'      background: var(--card-bg);' +
'      border-radius: 10px;' +
'      padding: 10px;' +
'      max-height: 150px;' +
'      overflow-y: auto;' +
'    }' +
'    .log-entry {' +
'      display: flex;' +
'      gap: 6px;' +
'      padding: 3px 0;' +
'      font-size: 10px;' +
'      border-bottom: 1px solid rgba(255,255,255,0.03);' +
'    }' +
'    .log-entry:last-child { border-bottom: none; }' +
'    .log-time {' +
'      color: #64748b;' +
'      font-family: monospace;' +
'      min-width: 50px;' +
'      flex-shrink: 0;' +
'    }' +
'    .log-msg { color: var(--muted); flex: 1; }' +
'    .log-entry.success .log-msg { color: var(--success); }' +
'    .log-entry.error .log-msg { color: var(--error); }' +
'    .log-entry.warn .log-msg { color: var(--warning); }' +
'    .log-entry.info .log-msg { color: #3b82f6; }' +
'    ' +
'    /* Confetti */' +
'    .confetti {' +
'      position: fixed;' +
'      top: 0;' +
'      left: 0;' +
'      right: 0;' +
'      bottom: 0;' +
'      pointer-events: none;' +
'      z-index: 1000;' +
'      display: none;' +
'    }' +
'    ' +
'    /* Footer */' +
'    .footer {' +
'      text-align: center;' +
'      padding: 8px;' +
'      font-size: 8px;' +
'      color: #475569;' +
'    }' +
'  </style>' +
SSCCP_AD_getSidebarHTML_Part4_();
}

/**
 * Part 4 of sidebar HTML - JavaScript state and UI functions
 */
function SSCCP_AD_getSidebarHTML_Part4_() {
  return '' +
'  <script>' +
'    // ══════════════════════════════════════════════════════════' +
'    // STATE' +
'    // ══════════════════════════════════════════════════════════' +
'    var state = {' +
'      isRunning: false,' +
'      trackers: [],' +
'      totalLeads: 0,' +
'      totalReps: 0,' +
'      pollInterval: null,' +
'      selectedReportType: "UNCONTACTED_SMS_CALL",' +
'      sendManagerReports: false' +
'    };' +
'    ' +
'    // Report type descriptions' +
'    var reportDescriptions = {' +
'      "UNCONTACTED_SMS_CALL": "Leads not meeting SMS or Call/VM contact thresholds",' +
'      "PRIORITY_1": "High-priority leads requiring immediate attention",' +
'      "PRIORITY_35": "Medium-priority leads needing follow-up",' +
'      "QUOTED_FOLLOWUP": "Leads with quotes requiring follow-up contact",' +
'      "ALL_TRACKERS": "All leads from all trackers"' +
'    };' +
'    ' +
'    // Tracker keys per report type' +
'    var reportTrackers = {' +
'      "UNCONTACTED_SMS_CALL": ["SMS", "CALL"],' +
'      "PRIORITY_1": ["P1"],' +
'      "PRIORITY_35": ["P35"],' +
'      "QUOTED_FOLLOWUP": ["SMS", "CALL"],' +
'      "ALL_TRACKERS": ["SMS", "CALL", "P1", "P35"]' +
'    };' +
'    ' +
'    // ══════════════════════════════════════════════════════════' +
'    // INITIALIZATION' +
'    // ══════════════════════════════════════════════════════════' +
'    document.addEventListener("DOMContentLoaded", function() {' +
'      loadInitialData();' +
'    });' +
'    ' +
'    function loadInitialData() {' +
'      log("Loading dashboard data...", "info");' +
'      ' +
'      // Load mode' +
'      google.script.run' +
'        .withSuccessHandler(function(mode) {' +
'          updateModeBadge(mode);' +
'        })' +
'        .SSCCP_AD_getMode();' +
'      ' +
'      // Load tracker status' +
'      google.script.run' +
'        .withSuccessHandler(function(data) {' +
'          updateTrackerCards(data.trackers);' +
'          updateBreakdown(data);' +
'          log("Dashboard loaded", "success");' +
'        })' +
'        .withFailureHandler(function(err) {' +
'          log("Failed to load: " + err, "error");' +
'        })' +
'        .SSCCP_AD_getTrackerStatus();' +
'      ' +
'      // Load quota' +
'      google.script.run' +
'        .withSuccessHandler(updateQuota)' +
'        .SSCCP_AD_getEmailQuota();' +
'    }' +
'    ' +
'    // ══════════════════════════════════════════════════════════' +
'    // REPORT TYPE HANDLER' +
'    // ══════════════════════════════════════════════════════════' +
'    function onReportTypeChange() {' +
'      var select = document.getElementById("reportType");' +
'      state.selectedReportType = select.value;' +
'      ' +
'      // Update description' +
'      document.getElementById("reportDescription").textContent = reportDescriptions[state.selectedReportType] || "";' +
'      ' +
'      // Highlight relevant trackers' +
'      var activeTrackers = reportTrackers[state.selectedReportType] || [];' +
'      var cards = document.querySelectorAll(".tracker-card");' +
'      cards.forEach(function(card) {' +
'        var key = card.getAttribute("data-key");' +
'        if (activeTrackers.indexOf(key) >= 0) {' +
'          card.classList.add("selected");' +
'        } else {' +
'          card.classList.remove("selected");' +
'        }' +
'      });' +
'      ' +
'      log("Report type: " + state.selectedReportType, "info");' +
'    }' +
'    ' +
'    // ══════════════════════════════════════════════════════════' +
'    // UI UPDATE FUNCTIONS' +
'    // ══════════════════════════════════════════════════════════' +
'    function updateModeBadge(mode) {' +
'      var badge = document.getElementById("modeBadge");' +
'      if (mode.testMode) {' +
'        badge.className = "mode-badge test";' +
'        badge.textContent = "🧪 TEST MODE - Admin Only";' +
'      } else if (mode.safeMode) {' +
'        badge.className = "mode-badge safe";' +
'        badge.textContent = "🔒 SAFE MODE - Admin Only";' +
'      } else {' +
'        badge.className = "mode-badge live";' +
'        badge.textContent = "🟢 LIVE MODE";' +
'      }' +
'    }' +
'    ' +
'    function updateTrackerCards(trackers) {' +
'      state.trackers = trackers;' +
'      var container = document.getElementById("trackerCards");' +
'      var activeTrackerKeys = reportTrackers[state.selectedReportType] || [];' +
'      var html = "";' +
'      ' +
'      trackers.forEach(function(t) {' +
'        var isSelected = activeTrackerKeys.indexOf(t.key) >= 0;' +
'        var cardClass = "tracker-card";' +
'        if (t.leads > 0) cardClass += " active";' +
'        if (isSelected) cardClass += " selected";' +
'        ' +
'        html += "<div class=\\"" + cardClass + "\\" data-key=\\"" + t.key + "\\">";' +
'        html += "<div class=\\"tracker-header\\">";' +
'        html += "<span class=\\"tracker-name\\">" + t.icon + " " + t.shortName + "</span>";' +
'        html += "<span class=\\"tracker-params\\">P:" + t.priority + " Req:<" + t.reqCount + "</span>";' +
'        html += "</div>";' +
'        html += "<div class=\\"tracker-stats\\">";' +
'        html += "<span>Updated: " + t.freshness + "</span>";' +
'        html += "<span class=\\"tracker-leads\\">" + t.leads + " leads</span>";' +
'        html += "</div>";' +
'        html += "<div class=\\"tracker-actions\\">";' +
'        html += "<button class=\\"tracker-btn\\" onclick=\\"jumpToTracker(\'" + t.key + "\')\\">📄 Open</button>";' +
'        html += "<button class=\\"tracker-btn\\" onclick=\\"editParams(\'" + t.key + "\')\\">⚙️ Edit</button>";' +
'        html += "</div>";' +
'        html += "</div>";' +
'      });' +
'      ' +
'      container.innerHTML = html;' +
'      ' +
'      // Enable buttons if we have data' +
'      var hasLeads = trackers.some(function(t) { return t.leads > 0; });' +
'      document.getElementById("btnPreview").disabled = !hasLeads;' +
'      document.getElementById("btnDry").disabled = !hasLeads;' +
'      document.getElementById("btnLive").disabled = !hasLeads;' +
'    }' +
'    ' +
'    function updateBreakdown(data) {' +
'      document.getElementById("statP0").textContent = data.byPriority.p0 || 0;' +
'      document.getElementById("statP1").textContent = data.byPriority.p1 || 0;' +
'      document.getElementById("statP3").textContent = data.byPriority.p3 || 0;' +
'      document.getElementById("statP5").textContent = data.byPriority.p5 || 0;' +
'      ' +
'      document.getElementById("totalReps").textContent = data.totalReps;' +
'      document.getElementById("totalLeads").textContent = data.totalLeads;' +
'      document.getElementById("avgPerRep").textContent = data.totalReps > 0 ' +
'        ? Math.round(data.totalLeads / data.totalReps) : "-";' +
'      ' +
'      state.totalLeads = data.totalLeads;' +
'      state.totalReps = data.totalReps;' +
'    }' +
'    ' +
'    function updateQuota(quota) {' +
'      var pct = Math.round((quota.used / quota.limit) * 100);' +
'      document.getElementById("quotaBar").style.width = pct + "%";' +
'      document.getElementById("quotaUsed").textContent = quota.used + " used";' +
'      document.getElementById("quotaRemaining").textContent = quota.remaining + " remaining";' +
'    }' +
'    ' +
'    function updateProgress(pct) {' +
'      document.getElementById("ringPct").textContent = Math.round(pct) + "%";' +
'      var offset = 345.4 * (1 - pct / 100);' +
'      document.getElementById("ringFill").style.strokeDashoffset = offset;' +
'    }' +
'    ' +
'    function setStage(stageNum, status, detail, count) {' +
'      for (var i = 1; i <= 5; i++) {' +
'        var el = document.getElementById("stage" + i);' +
'        el.classList.remove("active", "complete", "error");' +
'        if (i < stageNum) el.classList.add("complete");' +
'        else if (i === stageNum) el.classList.add(status);' +
'      }' +
'      if (detail) {' +
'        document.getElementById("stage" + stageNum + "Detail").textContent = detail;' +
'      }' +
'      if (count !== undefined) {' +
'        var countEl = document.getElementById("stage" + stageNum + "Count");' +
'        countEl.textContent = count;' +
'        countEl.style.display = "block";' +
'      }' +
'    }' +
SSCCP_AD_getSidebarHTML_Part5_();
}

/**
 * Part 5 of sidebar HTML - JavaScript action functions and polling
 */
function SSCCP_AD_getSidebarHTML_Part5_() {
  return '' +
'    ' +
'    // ══════════════════════════════════════════════════════════' +
'    // ACTION FUNCTIONS' +
'    // ══════════════════════════════════════════════════════════' +
'    function refreshTrackers() {' +
'      log("Refreshing all trackers...", "info");' +
'      document.getElementById("btnRefresh").disabled = true;' +
'      document.getElementById("mainStatus").textContent = "Refreshing Trackers...";' +
'      setStage(2, "active", "Running...");' +
'      ' +
'      google.script.run' +
'        .withSuccessHandler(function(result) {' +
'          log("Trackers refreshed: " + result.summary, "success");' +
'          setStage(2, "complete", result.summary);' +
'          document.getElementById("btnRefresh").disabled = false;' +
'          document.getElementById("mainStatus").textContent = "Trackers Refreshed";' +
'          document.getElementById("mainSubStatus").textContent = result.summary;' +
'          ' +
'          // Reload data' +
'          loadInitialData();' +
'        })' +
'        .withFailureHandler(function(err) {' +
'          log("Refresh failed: " + err, "error");' +
'          setStage(2, "error", "Failed");' +
'          document.getElementById("btnRefresh").disabled = false;' +
'        })' +
'        .SSCCP_AD_refreshTrackers();' +
'    }' +
'    ' +
'    function previewEmails() {' +
'      log("Opening email preview...", "info");' +
'      state.sendManagerReports = document.getElementById("sendManagerReports").checked;' +
'      google.script.run.SSCCP_AD_showEmailPreview(state.selectedReportType);' +
'    }' +
'    ' +
'    function editParams(trackerKey) {' +
'      google.script.run.SSCCP_AD_showParamEditor(trackerKey);' +
'    }' +
'    ' +
'    function jumpToTracker(trackerKey) {' +
'      google.script.run.SSCCP_AD_jumpToTracker(trackerKey);' +
'    }' +
'    ' +
'    function startDryRun() {' +
'      log("Starting DRY RUN (no emails will be sent)...", "warn");' +
'      startRun(true);' +
'    }' +
'    ' +
'    function confirmLiveRun() {' +
'      state.sendManagerReports = document.getElementById("sendManagerReports").checked;' +
'      google.script.run' +
'        .withSuccessHandler(function(confirmed) {' +
'          if (confirmed) {' +
'            log("Starting LIVE RUN...", "info");' +
'            startRun(false);' +
'          } else {' +
'            log("Live run cancelled by user", "warn");' +
'          }' +
'        })' +
'        .SSCCP_AD_confirmLiveRun(state.selectedReportType, state.sendManagerReports);' +
'    }' +
'    ' +
'    function startRun(isDryRun) {' +
'      state.isRunning = true;' +
'      state.sendManagerReports = document.getElementById("sendManagerReports").checked;' +
'      setButtonsRunning(true);' +
'      ' +
'      google.script.run' +
'        .withSuccessHandler(function(result) {' +
'          log("Run started: " + result.runId, "success");' +
'          startPolling();' +
'        })' +
'        .withFailureHandler(function(err) {' +
'          log("Failed to start: " + err, "error");' +
'          state.isRunning = false;' +
'          setButtonsRunning(false);' +
'        })' +
'        .SSCCP_AD_startRun(isDryRun, state.selectedReportType, state.sendManagerReports);' +
'    }' +
'    ' +
'    function cancelRun() {' +
'      log("Requesting cancel...", "warn");' +
'      document.getElementById("btnCancel").disabled = true;' +
'      google.script.run' +
'        .withSuccessHandler(function() {' +
'          log("Cancel signal sent", "warn");' +
'        })' +
'        .SSCCP_AD_cancelRun();' +
'    }' +
'    ' +
'    function setButtonsRunning(running) {' +
'      document.getElementById("btnRefresh").disabled = running;' +
'      document.getElementById("btnPreview").disabled = running;' +
'      document.getElementById("btnDry").disabled = running;' +
'      document.getElementById("btnLive").disabled = running;' +
'      document.getElementById("reportType").disabled = running;' +
'      document.getElementById("sendManagerReports").disabled = running;' +
'      document.getElementById("btnCancel").style.display = running ? "flex" : "none";' +
'      document.getElementById("btnCancel").disabled = false;' +
'    }' +
'    ' +
'    // ══════════════════════════════════════════════════════════' +
'    // POLLING' +
'    // ══════════════════════════════════════════════════════════' +
'    function startPolling() {' +
'      state.pollInterval = setInterval(function() {' +
'        google.script.run' +
'          .withSuccessHandler(handlePollResult)' +
'          .withFailureHandler(function(err) {' +
'            log("Poll error: " + err, "error");' +
'          })' +
'          .SSCCP_AD_getRunStatus();' +
'      }, 1500);' +
'    }' +
'    ' +
'    function stopPolling() {' +
'      if (state.pollInterval) {' +
'        clearInterval(state.pollInterval);' +
'        state.pollInterval = null;' +
'      }' +
'    }' +
'    ' +
'    function handlePollResult(status) {' +
'      // Update progress' +
'      updateProgress(status.progress || 0);' +
'      document.getElementById("mainStatus").textContent = status.stageName || "Running...";' +
'      document.getElementById("mainSubStatus").textContent = status.stageDetail || "";' +
'      ' +
'      // Update stage' +
'      if (status.currentStage) {' +
'        setStage(status.currentStage, "active", status.stageDetail, status.stageCount);' +
'      }' +
'      ' +
'      // Log entry' +
'      if (status.log) {' +
'        log(status.log, status.logType || "info");' +
'      }' +
'      ' +
'      // Check completion' +
'      if (status.complete || status.cancelled || status.error) {' +
'        stopPolling();' +
'        state.isRunning = false;' +
'        setButtonsRunning(false);' +
'        ' +
'        if (status.cancelled) {' +
'          log("⛔ Run cancelled", "warn");' +
'          document.getElementById("mainStatus").textContent = "Cancelled";' +
'        } else if (status.error) {' +
'          log("❌ Error: " + status.error, "error");' +
'          document.getElementById("mainStatus").textContent = "Error";' +
'          setStage(status.currentStage || 1, "error", status.error);' +
'        } else {' +
'          log("✅ Run completed! " + status.summary, "success");' +
'          document.getElementById("mainStatus").textContent = "Complete!";' +
'          document.getElementById("mainSubStatus").textContent = status.summary;' +
'          updateProgress(100);' +
'          showConfetti();' +
'          ' +
'          // Reload quota' +
'          google.script.run.withSuccessHandler(updateQuota).SSCCP_AD_getEmailQuota();' +
'        }' +
'      }' +
'    }' +
'    ' +
'    // ══════════════════════════════════════════════════════════' +
'    // UTILITIES' +
'    // ══════════════════════════════════════════════════════════' +
'    function log(msg, type) {' +
'      var container = document.getElementById("logContainer");' +
'      var entry = document.createElement("div");' +
'      entry.className = "log-entry " + (type || "");' +
'      ' +
'      var time = document.createElement("span");' +
'      time.className = "log-time";' +
'      time.textContent = new Date().toLocaleTimeString("en-US", { hour12: false });' +
'      ' +
'      var msgEl = document.createElement("span");' +
'      msgEl.className = "log-msg";' +
'      msgEl.textContent = msg;' +
'      ' +
'      entry.appendChild(time);' +
'      entry.appendChild(msgEl);' +
'      container.appendChild(entry);' +
'      container.scrollTop = container.scrollHeight;' +
'    }' +
'    ' +
'    function showConfetti() {' +
'      var canvas = document.getElementById("confetti");' +
'      canvas.style.display = "block";' +
'      var ctx = canvas.getContext("2d");' +
'      canvas.width = window.innerWidth;' +
'      canvas.height = window.innerHeight;' +
'      ' +
'      var particles = [];' +
'      var colors = ["#8B1538", "#D4AF37", "#10b981", "#3b82f6", "#f59e0b"];' +
'      ' +
'      for (var i = 0; i < 150; i++) {' +
'        particles.push({' +
'          x: Math.random() * canvas.width,' +
'          y: Math.random() * canvas.height - canvas.height,' +
'          r: Math.random() * 6 + 2,' +
'          d: Math.random() * 150,' +
'          color: colors[Math.floor(Math.random() * colors.length)],' +
'          tilt: Math.random() * 10 - 10' +
'        });' +
'      }' +
'      ' +
'      var frames = 0;' +
'      function draw() {' +
'        ctx.clearRect(0, 0, canvas.width, canvas.height);' +
'        particles.forEach(function(p) {' +
'          ctx.beginPath();' +
'          ctx.fillStyle = p.color;' +
'          ctx.fillRect(p.x, p.y, p.r, p.r * 1.5);' +
'          p.y += Math.cos(p.d) + 3;' +
'          p.x += Math.sin(frames / 50) * 2;' +
'          if (p.y > canvas.height) p.y = -10;' +
'        });' +
'        frames++;' +
'        if (frames < 200) {' +
'          requestAnimationFrame(draw);' +
'        } else {' +
'          canvas.style.display = "none";' +
'        }' +
'      }' +
'      draw();' +
'    }' +
'  </script>' +
'</body>' +
'</html>';
}

/* ═══════════════════════════════════════════════════════════════
 * BACKEND API FUNCTIONS
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Get current mode (test/safe/live)
 */
function SSCCP_AD_getMode() {
  var props = PropertiesService.getScriptProperties();
  return {
    testMode: props.getProperty("SSCCP_TEST_MODE") === "true",
    safeMode: props.getProperty("SSCCP_SAFE_MODE") === "true"
  };
}

/**
 * Get status of all trackers including freshness, params, and lead counts
 */
function SSCCP_AD_getTrackerStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cache = CacheService.getScriptCache();
  var lastRefresh = cache.get(SSCCP_AD_CONFIG.CACHE_KEYS.LAST_REFRESH);
  
  var trackers = [];
  var byPriority = { p0: 0, p1: 0, p3: 0, p5: 0 };
  var totalLeads = 0;
  var repSet = new Set();
  
  SSCCP_AD_CONFIG.TRACKERS.forEach(function(cfg) {
    var sheet = ss.getSheetByName(cfg.name);
    var tracker = {
      key: cfg.key,
      name: cfg.name,
      shortName: cfg.key === "SMS" ? "SMS" : 
                 cfg.key === "CALL" ? "CALL/VM" :
                 cfg.key === "P1" ? "PRIORITY 1" : "PRIORITY 3,5",
      icon: cfg.icon,
      type: cfg.type,
      leads: 0,
      priority: "-",
      reqCount: 1,
      freshness: lastRefresh ? SSCCP_AD_timeAgo_(new Date(lastRefresh)) : "Never"
    };
    
    if (sheet) {
      // Read params
      try {
        var priorityCell = sheet.getRange(cfg.priorityCell).getValue();
        var countCell = sheet.getRange(cfg.countCell).getValue();
        tracker.priority = String(priorityCell || "0");
        tracker.reqCount = parseInt(countCell, 10) || 0;
        tracker.reqCount = tracker.reqCount + 1; // Convert to "less than" threshold
      } catch (e) {}
      
      // Count leads (starting row 3, column F onwards)
      var lastRow = sheet.getLastRow();
      if (lastRow >= 3) {
        var data = sheet.getRange(3, 6, lastRow - 2, 6).getValues(); // F:K
        data.forEach(function(row) {
          if (row[1]) { // PHONE NUMBER exists (col G)
            tracker.leads++;
            var user = String(row[4] || "").toUpperCase().trim(); // USER col J
            if (user) repSet.add(user);
          }
        });
      }
      
      totalLeads += tracker.leads;
      
      // Count by priority from this tracker
      if (cfg.key === "SMS" || cfg.key === "CALL") {
        byPriority.p0 += tracker.leads;
      } else if (cfg.key === "P1") {
        byPriority.p1 += tracker.leads;
      } else if (cfg.key === "P35") {
        // Need to check actual priority values
        if (lastRow >= 3) {
          var priData = sheet.getRange(3, 27, lastRow - 2, 1).getValues(); // AA = priority
          priData.forEach(function(r) {
            var p = parseInt(r[0], 10);
            if (p === 3) byPriority.p3++;
            else if (p === 5) byPriority.p5++;
          });
        }
      }
    }
    
    trackers.push(tracker);
  });
  
  return {
    trackers: trackers,
    byPriority: byPriority,
    totalLeads: totalLeads,
    totalReps: repSet.size
  };
}

/**
 * Get email quota status
 */
function SSCCP_AD_getEmailQuota() {
  var remaining = MailApp.getRemainingDailyQuota();
  var limit = SSCCP_AD_CONFIG.DAILY_EMAIL_LIMIT;
  return {
    used: limit - remaining,
    remaining: remaining,
    limit: limit
  };
}

/**
 * Refresh all trackers by running SSCCP_buildTrackersFromGranotData
 */
function SSCCP_AD_refreshTrackers() {
  // Call the tracker builder
  var result = SSCCP_buildTrackersFromGranotData();
  
  // Store refresh time
  var cache = CacheService.getScriptCache();
  cache.put(SSCCP_AD_CONFIG.CACHE_KEYS.LAST_REFRESH, new Date().toISOString(), 21600);
  
  var summary = "SMS:" + (result.trackers.smsTracker ? result.trackers.smsTracker.rowsWritten : 0) +
                " CALL:" + (result.trackers.callTracker ? result.trackers.callTracker.rowsWritten : 0) +
                " P1:" + (result.trackers.priority1Tracker ? result.trackers.priority1Tracker.rowsWritten : 0) +
                " P3/5:" + (result.trackers.priority35Tracker ? result.trackers.priority35Tracker.rowsWritten : 0);
  
  return {
    success: true,
    summary: summary,
    duration: result.duration
  };
}

/**
 * Jump to a specific tracker sheet
 */
function SSCCP_AD_jumpToTracker(trackerKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = SSCCP_AD_CONFIG.TRACKERS.find(function(t) { return t.key === trackerKey; });
  if (cfg) {
    var sheet = ss.getSheetByName(cfg.name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      sheet.getRange("F3").activate();
    }
  }
}

/**
 * Show parameter editor dialog
 */
function SSCCP_AD_showParamEditor(trackerKey) {
  var cfg = SSCCP_AD_CONFIG.TRACKERS.find(function(t) { return t.key === trackerKey; });
  if (!cfg) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.name);
  if (!sheet) return;
  
  var currentPriority = sheet.getRange(cfg.priorityCell).getValue();
  var currentCount = sheet.getRange(cfg.countCell).getValue();
  
  var html = HtmlService.createHtmlOutput(SSCCP_AD_getParamEditorHTML_(trackerKey, cfg, currentPriority, currentCount))
    .setWidth(300)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, "Edit " + cfg.name + " Parameters");
}

function SSCCP_AD_getParamEditorHTML_(trackerKey, cfg, currentPriority, currentCount) {
  var priorityOptions = "";
  for (var i = 0; i <= 8; i++) {
    priorityOptions += '<option value="' + i + '"' + (i == currentPriority ? ' selected' : '') + '>' + i + '</option>';
  }
  // Multi-priority options
  var multiOptions = [
    { val: "1, 3", label: "1, 3" },
    { val: "3, 5", label: "3, 5" },
    { val: "1, 3, 5", label: "1, 3, 5" }
  ];
  multiOptions.forEach(function(opt) {
    priorityOptions += '<option value="' + opt.val + '"' + (opt.val == currentPriority ? ' selected' : '') + '>' + opt.label + '</option>';
  });
  
  var countOptions = "";
  for (var c = 0; c <= 5; c++) {
    countOptions += '<option value="' + c + '"' + (c == currentCount ? ' selected' : '') + '>' + c + ' (show <' + (c+1) + ')</option>';
  }
  
  return '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:sans-serif;padding:16px;background:#1e293b;color:#fff}' +
    '.field{margin-bottom:16px}' +
    'label{display:block;font-size:12px;color:#94a3b8;margin-bottom:4px}' +
    'select{width:100%;padding:8px;border-radius:6px;border:1px solid #475569;background:#0f172a;color:#fff;font-size:14px}' +
    '.btn{width:100%;padding:10px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;margin-top:8px}' +
    '.btn-save{background:#10b981;color:#fff}' +
    '.btn-cancel{background:#475569;color:#fff}' +
    '</style></head><body>' +
    '<div class="field"><label>Priority</label>' +
    '<select id="priority">' + priorityOptions + '</select></div>' +
    '<div class="field"><label>Required Count (leads with less than X)</label>' +
    '<select id="count">' + countOptions + '</select></div>' +
    '<button class="btn btn-save" onclick="save()">💾 Save & Refresh</button>' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<script>' +
    'function save(){' +
    'var p=document.getElementById("priority").value;' +
    'var c=document.getElementById("count").value;' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();}).SSCCP_AD_saveParams("' + trackerKey + '",p,c);' +
    '}' +
    '</script></body></html>';
}

/**
 * Save tracker parameters
 */
function SSCCP_AD_saveParams(trackerKey, priority, count) {
  var cfg = SSCCP_AD_CONFIG.TRACKERS.find(function(t) { return t.key === trackerKey; });
  if (!cfg) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.name);
  if (!sheet) return;
  
  sheet.getRange(cfg.priorityCell).setValue(priority);
  sheet.getRange(cfg.countCell).setValue(parseInt(count, 10));
  
  // Trigger refresh for just this tracker
  SpreadsheetApp.flush();
  
  // Show toast
  ss.toast("Parameters saved! Run Refresh Trackers to update.", "✔ Saved", 3);
}

/**
 * Show email preview dialog
 */
function SSCCP_AD_showEmailPreview(reportType) {
  var html = HtmlService.createHtmlOutput(SSCCP_AD_getPreviewHTML_(reportType))
    .setWidth(700)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "📧 Email Preview - " + (reportType || "All"));
}

function SSCCP_AD_getPreviewHTML_(reportType) {
  // Get relevant trackers for this report type
  var reportConfig = SSCCP_AD_CONFIG.REPORT_TYPES.find(function(r) { return r.key === reportType; });
  var trackerKeys = reportConfig ? reportConfig.trackers : ["SMS", "CALL", "P1", "P35"];
  
  // Gather lead data from selected trackers
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var repLeads = {};
  
  SSCCP_AD_CONFIG.TRACKERS.forEach(function(cfg) {
    // Skip trackers not in this report type
    if (trackerKeys.indexOf(cfg.key) < 0) return;
    
    var sheet = ss.getSheetByName(cfg.name);
    if (!sheet) return;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    
    var data = sheet.getRange(3, 6, lastRow - 2, 8).getValues(); // F:M
    data.forEach(function(row) {
      if (!row[1]) return; // No phone
      var user = String(row[4] || "").toUpperCase().trim();
      if (!user) return;
      
      if (!repLeads[user]) repLeads[user] = [];
      repLeads[user].push({
        tracker: cfg.key,
        phone: row[1],
        job: row[3],
        moveDate: row[0],
        source: row[6]
      });
    });
  });
  
  var html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:sans-serif;padding:0;margin:0;background:#0f172a;color:#fff}' +
    '.header{background:linear-gradient(135deg,#8B1538,#6B1028);padding:16px;border-bottom:3px solid #D4AF37}' +
    '.header h2{margin:0;font-size:16px}' +
    '.content{padding:16px;max-height:400px;overflow-y:auto}' +
    '.rep-card{background:#1e293b;border-radius:8px;padding:12px;margin-bottom:12px;border-left:3px solid #D4AF37}' +
    '.rep-name{font-weight:bold;font-size:14px;margin-bottom:8px;color:#D4AF37}' +
    '.lead-row{display:flex;gap:12px;padding:4px 0;border-bottom:1px solid rgba(255,255,255,0.1);font-size:11px}' +
    '.lead-row:last-child{border-bottom:none}' +
    '.tracker-badge{padding:2px 6px;border-radius:4px;font-size:9px;font-weight:bold}' +
    '.tracker-SMS{background:#3b82f6}' +
    '.tracker-CALL{background:#10b981}' +
    '.tracker-P1{background:#ef4444}' +
    '.tracker-P35{background:#8b5cf6}' +
    '.summary{background:#1e293b;padding:12px;margin:16px;border-radius:8px}' +
    '.info-banner{background:#3b82f6;color:#fff;padding:10px;margin:16px;border-radius:6px;font-size:11px}' +
    '</style></head><body>' +
    '<div class="header"><h2>📧 Email Preview - ' + (reportConfig ? reportConfig.name : "All") + '</h2></div>' +
    '<div class="info-banner">✨ <strong>Premium Format:</strong> Actual emails will include RC enrichment with call/SMS counts, stoplight indicators, last 5 calls/SMS history, and process direction.</div>' +
    '<div class="summary"><strong>' + Object.keys(repLeads).length + '</strong> reps will receive emails with <strong>' + 
    Object.values(repLeads).reduce(function(a, b) { return a + b.length; }, 0) + '</strong> total leads</div>' +
    '<div class="content">';
  
  Object.keys(repLeads).sort().forEach(function(rep) {
    var leads = repLeads[rep];
    html += '<div class="rep-card">';
    html += '<div class="rep-name">👤 ' + rep + ' (' + leads.length + ' leads)</div>';
    leads.slice(0, 5).forEach(function(lead) {
      html += '<div class="lead-row">';
      html += '<span class="tracker-badge tracker-' + lead.tracker + '">' + lead.tracker + '</span>';
      html += '<span>' + lead.phone + '</span>';
      html += '<span>Job: ' + lead.job + '</span>';
      html += '<span>' + lead.source + '</span>';
      html += '</div>';
    });
    if (leads.length > 5) {
      html += '<div style="font-size:10px;color:#94a3b8;padding:4px 0;">+ ' + (leads.length - 5) + ' more leads...</div>';
    }
    html += '</div>';
  });
  
  html += '</div></body></html>';
  return html;
}

/**
 * Show confirmation dialog for live run
 */
function SSCCP_AD_confirmLiveRun(reportType, sendManagerReports) {
  var status = SSCCP_AD_getTrackerStatus();
  var reportConfig = SSCCP_AD_CONFIG.REPORT_TYPES.find(function(r) { return r.key === reportType; });
  var ui = SpreadsheetApp.getUi();
  
  var msg = "You are about to send emails to " + status.totalReps + " reps with " + status.totalLeads + " total leads.\n\n" +
    "Report Type: " + (reportConfig ? reportConfig.name : "Unknown") + "\n" +
    "Manager Reports: " + (sendManagerReports ? "YES" : "NO") + "\n\n" +
    "This will send REAL emails using the PREMIUM format with RC enrichment.\n\n" +
    "Are you sure you want to proceed?";
  
  var response = ui.alert("⚠️ Confirm LIVE Run", msg, ui.ButtonSet.YES_NO);
  
  return response === ui.Button.YES;
}

/* ═══════════════════════════════════════════════════════════════
 * RUN MANAGEMENT - v1.1: Uses Premium Pipeline
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Start alert run (dry or live)
 * v1.1: Now accepts reportType and sendManagerReports
 */
function SSCCP_AD_startRun(isDryRun, reportType, sendManagerReports) {
  var cache = CacheService.getScriptCache();
  var runId = "RUN_" + new Date().getTime();
  
  // Initialize state
  var state = {
    runId: runId,
    isDryRun: isDryRun,
    reportType: reportType || "UNCONTACTED_SMS_CALL",
    sendManagerReports: sendManagerReports || false,
    startTime: new Date().toISOString(),
    currentStage: 1,
    stageName: "Gate Checks",
    stageDetail: "Initializing...",
    progress: 0,
    leadsProcessed: 0,
    emailsSent: 0,
    currentRep: "",
    repIndex: 0,
    totalReps: 0,
    complete: false,
    cancelled: false,
    error: null,
    logs: []
  };
  
  cache.put(SSCCP_AD_CONFIG.CACHE_KEYS.RUN_STATE, JSON.stringify(state), 3600);
  cache.remove(SSCCP_AD_CONFIG.CACHE_KEYS.CANCEL_FLAG);
  
  // Start async execution
  SSCCP_AD_executeRun_(runId, isDryRun, reportType, sendManagerReports);
  
  return { runId: runId };
}

/**
 * Get current run status
 */
function SSCCP_AD_getRunStatus() {
  var cache = CacheService.getScriptCache();
  var stateJson = cache.get(SSCCP_AD_CONFIG.CACHE_KEYS.RUN_STATE);
  
  if (!stateJson) {
    return { complete: true, error: "No active run" };
  }
  
  return JSON.parse(stateJson);
}

/**
 * Cancel current run
 */
function SSCCP_AD_cancelRun() {
  var cache = CacheService.getScriptCache();
  cache.put(SSCCP_AD_CONFIG.CACHE_KEYS.CANCEL_FLAG, "true", 300);
}

/**
 * Execute the alert pipeline - v1.1: ACTUALLY SENDS EMAILS with premium format
 */
function SSCCP_AD_executeRun_(runId, isDryRun, reportType, sendManagerReports) {
  var cache = CacheService.getScriptCache();
  var CFG = getConfig_();
  
  function updateState(updates) {
    var stateJson = cache.get(SSCCP_AD_CONFIG.CACHE_KEYS.RUN_STATE);
    var state = stateJson ? JSON.parse(stateJson) : {};
    Object.keys(updates).forEach(function(k) { state[k] = updates[k]; });
    cache.put(SSCCP_AD_CONFIG.CACHE_KEYS.RUN_STATE, JSON.stringify(state), 3600);
    return state;
  }
  
  function checkCancel() {
    return cache.get(SSCCP_AD_CONFIG.CACHE_KEYS.CANCEL_FLAG) === "true";
  }
  
  try {
    // Get mode settings
    var mode = SSCCP_AD_getMode();
    var routeToAdmin = mode.testMode || mode.safeMode || CFG.SAFE_MODE || CFG.FORWARD_ALL;
    var adminEmail = CFG.ADMIN_EMAIL;
    
    // Get report config
    var reportConfig = SSCCP_AD_CONFIG.REPORT_TYPES.find(function(r) { return r.key === reportType; });
    var trackerKeys = reportConfig ? reportConfig.trackers : ["SMS", "CALL"];
    var reportTitle = reportConfig ? reportConfig.name.replace(/^[^\s]+\s/, "") : "Uncontacted Leads Alert";
    
    // Stage 1: Gate Checks
    updateState({ currentStage: 1, stageName: "Gate Checks", stageDetail: "Verifying configuration...", progress: 5 });
    
    if (!adminEmail && routeToAdmin) {
      updateState({ error: "Missing ADMIN_EMAIL in config", stageName: "Error", stageDetail: "Missing ADMIN_EMAIL" });
      return;
    }
    
    if (checkCancel()) {
      updateState({ cancelled: true, stageName: "Cancelled" });
      return;
    }
    
    updateState({ stageDetail: "✔ Gate checks passed" + (routeToAdmin ? " (Admin mode)" : ""), progress: 10 });
    
    // Stage 2: Refresh Trackers
    updateState({ currentStage: 2, stageName: "Refresh Trackers", stageDetail: "Updating tracker data...", progress: 15 });
    
    var refreshResult = SSCCP_AD_refreshTrackers();
    
    if (checkCancel()) {
      updateState({ cancelled: true, stageName: "Cancelled" });
      return;
    }
    
    updateState({ stageDetail: "✔ " + refreshResult.summary, progress: 25 });
    
    // Stage 3: Scan Sources - Build lead data from selected trackers
    updateState({ currentStage: 3, stageName: "Scan Sources", stageDetail: "Loading lead data...", progress: 30 });
    
    var repData = SSCCP_AD_buildRepLeadData_(trackerKeys);
    var repCount = Object.keys(repData).length;
    var totalLeads = 0;
    for (var r in repData) {
      totalLeads += repData[r].leads.length;
    }
    
    if (checkCancel()) {
      updateState({ cancelled: true, stageName: "Cancelled" });
      return;
    }
    
    updateState({ 
      stageDetail: "✔ Found " + totalLeads + " leads",
      stageCount: totalLeads,
      progress: 40,
      log: "Found " + totalLeads + " leads across " + repCount + " reps",
      logType: "success"
    });
    
    // Stage 4: RC Enrichment - This is the key difference from v1.0!
    updateState({ currentStage: 4, stageName: "RC Enrichment", stageDetail: "Building RC lookup index...", progress: 45 });
    
    // Build RC lookup index using existing function from RC_Enrichment.js
    var rcIndex = {};
    try {
      rcIndex = RC_buildLookupIndex_();
      updateState({ stageDetail: "✔ RC index built (" + Object.keys(rcIndex).length + " phones)", progress: 55 });
    } catch (e) {
      updateState({ stageDetail: "⚠️ RC enrichment skipped: " + e.message, progress: 55 });
    }
    
    // Enrich leads with RC data
    updateState({ stageDetail: "Enriching leads with RC data...", progress: 58 });
    
    for (var repUser in repData) {
      var rep = repData[repUser];
      // Use existing RC_enrichLeads_ function
      try {
        RC_enrichLeads_(rep.leads, rcIndex);
      } catch (e) {
        // Continue even if enrichment fails for one rep
      }
    }
    
    // Get RC freshness info
    var rcFreshness = {};
    try {
      rcFreshness = SSCCP_getRCFreshness_();
    } catch (e) {
      rcFreshness = { callLogLabel: "Unknown", smsLogLabel: "Unknown" };
    }
    
    if (checkCancel()) {
      updateState({ cancelled: true, stageName: "Cancelled" });
      return;
    }
    
    updateState({ stageDetail: "✔ Enrichment complete", progress: 60 });
    
    // Stage 5: Send Alerts
    updateState({ 
      currentStage: 5, 
      stageName: "Send Alerts", 
      stageDetail: isDryRun ? "Simulating emails..." : "Sending emails...",
      progress: 65,
      totalReps: repCount
    });
    
    var reps = Object.keys(repData).sort();
    var emailsSent = 0;
    var emailsFailed = 0;
    
    for (var i = 0; i < reps.length; i++) {
      if (checkCancel()) {
        updateState({ cancelled: true, stageName: "Cancelled", stageDetail: "Stopped at rep " + (i+1) + "/" + reps.length });
        return;
      }
      
      var repUsername = reps[i];
      var rep = repData[repUsername];
      var progress = 65 + (30 * (i + 1) / reps.length);
      
      updateState({
        stageDetail: (isDryRun ? "Simulating: " : "Emailing: ") + repUsername + " (" + (i+1) + "/" + reps.length + ")",
        currentRep: repUsername,
        repIndex: i + 1,
        progress: progress
      });
      
      if (isDryRun) {
        // Dry run - just simulate
        updateState({
          log: "🧪 Simulated: " + repUsername + " (" + rep.leads.length + " leads)",
          logType: "warn"
        });
        Utilities.sleep(100);
        emailsSent++;
      } else {
        // LIVE RUN - Actually send email using PREMIUM format
        try {
          var toEmail = routeToAdmin ? adminEmail : rep.email;
          
          if (!toEmail) {
            updateState({ log: "⚠️ Skipped " + repUsername + " - no email", logType: "warn" });
            emailsFailed++;
            continue;
          }
          
          // Build payload for premium renderer
          var payload = {
            rep: {
              name: rep.name || repUsername,
              username: repUsername,
              email: rep.email
            },
            leads: rep.leads,
            reportTitle: reportTitle,
            runId: runId,
            rcFreshness: rcFreshness,
            isTestMode: routeToAdmin,
            window: { label: "Alert Dashboard v1.1" }
          };
          
          // Add tracker type to each lead for process direction
          rep.leads.forEach(function(lead) {
            if (!lead.trackerType) {
              lead.trackerType = lead.tracker === "SMS" ? "SMS" : "CALL";
            }
          });
          
          // Use the premium email renderer from code.js
          var htmlBody = SSCCP_renderRepEmailHtml_PREMIUM_(CFG, payload);
          
          var subject = (routeToAdmin ? "[TEST:" + rep.email + "] " : "") +
                        "🟢 Safe Ship — " + reportTitle + " — " + (rep.name || repUsername);
          
          MailApp.sendEmail({
            to: toEmail,
            subject: subject,
            htmlBody: htmlBody,
            body: "You have " + rep.leads.length + " leads requiring contact. Please check your email client for the full report."
          });
          
          emailsSent++;
          
          updateState({
            log: "📧 Sent: " + repUsername + " (" + rep.leads.length + " leads) → " + toEmail,
            logType: "success"
          });
          
          // Log to Email_Log sheet
          SSCCP_AD_logEmail_(repUsername, toEmail, rep.leads.length, runId);
          
          // Rate limiting
          Utilities.sleep(150);
          
        } catch (e) {
          emailsFailed++;
          updateState({
            log: "❌ Failed: " + repUsername + " - " + e.message,
            logType: "error"
          });
        }
      }
    }
    
    // TODO: Manager summaries if sendManagerReports is true
    // This would call the existing manager report logic from code.js
    
    // Complete
    var summaryText = (isDryRun ? "DRY RUN: " : "") + emailsSent + " emails " + (isDryRun ? "simulated" : "sent");
    if (emailsFailed > 0) summaryText += ", " + emailsFailed + " failed";
    if (routeToAdmin && !isDryRun) summaryText += " (routed to admin)";
    
    updateState({
      complete: true,
      stageName: "Complete",
      stageDetail: "✔ All done!",
      progress: 100,
      emailsSent: emailsSent,
      summary: summaryText,
      log: "✅ " + summaryText,
      logType: "success"
    });
    
  } catch (error) {
    updateState({
      error: error.message,
      stageName: "Error",
      stageDetail: error.message,
      log: "❌ Error: " + error.message,
      logType: "error"
    });
  }
}

/**
 * Build complete rep lead data with emails and lead details
 * v1.1: Now filters by tracker keys for report type
 */
function SSCCP_AD_buildRepLeadData_(trackerKeys) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var repData = {};
  
  // Default to all trackers if not specified
  if (!trackerKeys || trackerKeys.length === 0) {
    trackerKeys = ["SMS", "CALL", "P1", "P35"];
  }
  
  // Gather leads from selected trackers
  SSCCP_AD_CONFIG.TRACKERS.forEach(function(cfg) {
    // Skip trackers not in the selected report type
    if (trackerKeys.indexOf(cfg.key) < 0) return;
    
    var sheet = ss.getSheetByName(cfg.name);
    if (!sheet) return;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    
    // Read more columns for full lead data
    // F=MOVE_DATE, G=PHONE, H=Count, I=JOB#, J=USER, K=CF, L=SOURCE, M=EXCLUDE, N+=extra
    var data = sheet.getRange(3, 6, lastRow - 2, 22).getValues(); // F:AA
    
    data.forEach(function(row) {
      if (!row[1]) return; // No phone number
      
      var user = String(row[4] || "").toUpperCase().trim();
      if (!user) return;
      
      if (!repData[user]) {
        repData[user] = {
          username: user,
          name: user,
          email: "",
          leads: []
        };
      }
      
      // Build lead object matching the format expected by premium renderer
      var lead = {
        tracker: cfg.key,
        trackerType: cfg.type === "SMS" ? "SMS" : "CALL",
        moveDate: row[0] ? Utilities.formatDate(new Date(row[0]), "America/New_York", "MM/dd/yyyy") : "",
        phone: row[1],
        count: row[2],
        job: row[3],
        cf: row[5],
        source: row[6],
        // Extra fields for premium email
        email: row[8] || "",    // Column N = email (after EXCLUDE)
        firstName: row[9] || "",
        lastName: row[10] || "",
        fromState: row[12] || "",
        fromCity: row[11] || "",
        cubicFeet: "",  // Not in tracker, but needed for template
        // RC data will be added by enrichment
        rc: null
      };
      
      repData[user].leads.push(lead);
    });
  });
  
  // Get emails and names from Sales_Roster
  try {
    var roster = ss.getSheetByName("Sales_Roster");
    if (roster) {
      var rosterData = roster.getDataRange().getValues();
      var headers = rosterData[0];
      var usernameIdx = -1, emailIdx = -1, nameIdx = -1;
      
      for (var h = 0; h < headers.length; h++) {
        var hdr = String(headers[h]).toLowerCase().trim();
        if (hdr === "username") usernameIdx = h;
        if (hdr === "primary email" || hdr === "email") emailIdx = h;
        if (hdr === "full name" || hdr === "name") nameIdx = h;
      }
      
      if (usernameIdx >= 0) {
        for (var i = 1; i < rosterData.length; i++) {
          var rosterUser = String(rosterData[i][usernameIdx] || "").toUpperCase().trim();
          if (repData[rosterUser]) {
            if (emailIdx >= 0) repData[rosterUser].email = rosterData[i][emailIdx] || "";
            if (nameIdx >= 0) repData[rosterUser].name = rosterData[i][nameIdx] || rosterUser;
          }
        }
      }
    }
  } catch (e) {
    Logger.log("Roster lookup error: " + e);
  }
  
  return repData;
}

/**
 * Log email to Email_Log sheet
 */
function SSCCP_AD_logEmail_(repName, email, leadCount, runId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SSCCP_AD_CONFIG.EMAIL_LOG_SHEET);
  
  if (!sheet) {
    // Create the sheet
    sheet = ss.insertSheet(SSCCP_AD_CONFIG.EMAIL_LOG_SHEET);
    sheet.appendRow(["Timestamp", "Run ID", "Rep", "Email", "Lead Count", "Status"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
  
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([timestamp, runId, repName, email, leadCount, "SENT"]);
}

/* ═══════════════════════════════════════════════════════════════
 * UTILITY FUNCTIONS
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_AD_timeAgo_(date) {
  if (!date) return "Never";
  var seconds = Math.floor((new Date() - date) / 1000);
  
  if (seconds < 60) return seconds + "s ago";
  if (seconds < 3600) return Math.floor(seconds / 60) + "m ago";
  if (seconds < 86400) return Math.floor(seconds / 3600) + "h ago";
  return Math.floor(seconds / 86400) + "d ago";
}