/**************************************************************
 * CONTACT ACTIVITY v3.5 - Shows Failed SMS with Red Indicator
 * 
 * ENHANCED version - not yet integrated into Code.gs
 * 
 * CHANGES:
 * - Failed SMS now show with FAILED badge instead of checkmark
 * - All SMS attempts are visible (no more "Never" when attempts were made)
 * - Warning banner if ALL SMS attempts failed
 * 
 * NOTE: Function renamed to _STAGED_ to avoid collision with Config.js
 *       When ready to integrate, replace Config.js version and rename back.
 **************************************************************/

function renderLeadContactActivity_v2_STAGED_(lead) {
  var rc = lead.rc || {};
  var html = '';
  
  // Alert: Lead answered through text
  if (rc.hasReplied) {
    html += '<div style="background:linear-gradient(135deg,#10b981,#059669);color:#fff;padding:8px 14px;border-radius:6px;margin:10px 18px 0;display:flex;align-items:center;gap:8px;font-size:12px;">';
    html += '<span style="font-size:16px;">[MSG]</span>';
    html += '<span><strong>This Lead Answered Through Text!</strong> Check SMS conversation for their response</span></div>';
  }
  
  // Alert: Wrong Priority Risk (call > 4 minutes)
  if (rc.hasLongCall) {
    html += '<div style="background:linear-gradient(135deg,#f59e0b,#d97706);color:#000;padding:8px 14px;border-radius:6px;margin:10px 18px 0;display:flex;align-items:center;gap:8px;font-size:12px;">';
    html += '<span style="font-size:16px;">[!]</span>';
    html += '<span><strong>Wrong Priority Risk - Review File</strong> (Call duration: ' + (rc.longestCallEverDisplay || rc.longestCallDisplay || "4+ min") + ')</span></div>';
  }
  
  // v3.4 NEW: Alert if ALL SMS attempts failed
  if (rc.allSmsFailed && rc.smsTotal > 0) {
    html += '<div style="background:linear-gradient(135deg,#ef4444,#dc2626);color:#fff;padding:8px 14px;border-radius:6px;margin:10px 18px 0;display:flex;align-items:center;gap:8px;font-size:12px;">';
    html += '<span style="font-size:16px;">[X]</span>';
    html += '<span><strong>All SMS Attempts Failed!</strong> Check phone number validity or try a different method</span></div>';
  }
  
  html += '<div style="padding:14px 18px;">';
  
  // Stats Row (Yesterday -> Today)
  html += '<div style="font-size:10px;color:#64748b;font-weight:600;margin-bottom:8px;">[CHART] CONTACT ACTIVITY (YESTERDAY -> TODAY)</div>';
  html += '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;">';
  html += renderStatBubble_v2_("SMS SENT", rc.smsYesterday || 0, rc.smsToday || 0, "[MOBILE]", "#dbeafe", "#1e40af");
  html += renderStatBubble_v2_("CALLS", rc.callsYesterday || 0, rc.callsToday || 0, "[PHONE]", "#dcfce7", "#166534");
  html += renderStatBubble_v2_("VOICEMAILS", rc.vmYesterday || 0, rc.vmToday || 0, "[VM]", "#fef3c7", "#92400e");
  html += '<div style="background:#f1f5f9;border-radius:6px;padding:6px 10px;text-align:center;min-width:70px;">';
  html += '<div style="font-size:9px;color:#64748b;font-weight:600;">LONGEST</div>';
  html += '<div style="font-size:14px;font-weight:800;color:#334155;">' + (rc.longestCallDisplay || "0:00") + '</div></div>';
  html += '</div>';
  
  // 
  // BUILD CALLS HTML (with IN badge for inbound)
  // 
  var callsHtml = '';
  var callBg = (rc.last5Calls && rc.last5Calls.length > 0) ? "#dcfce7" : "#fee2e2";
  var callColor = (rc.last5Calls && rc.last5Calls.length > 0) ? "#166534" : "#991b1b";
  
  if (rc.last5Calls && rc.last5Calls.length > 0) {
    for (var i = 0; i < rc.last5Calls.length; i++) {
      var call = rc.last5Calls[i];
      
      // IN badge for inbound calls (cyan)
      var inBadge = '';
      if (call.isInbound) {
        inBadge = '<span style="display:inline-block;background:#06b6d4;color:#fff;font-size:9px;font-weight:700;padding:1px 4px;border-radius:3px;margin-right:4px;">IN</span>';
      }
      
      // VM badge for voicemails (orange) - only for outbound
      var vmBadge = '';
      if (call.isVM && !call.isInbound) {
        vmBadge = '<span style="display:inline-block;background:#f59e0b;color:#fff;font-size:9px;font-weight:700;padding:1px 4px;border-radius:3px;margin-left:4px;">VM</span>';
      }
      
      callsHtml += '<div style="font-size:10px;color:' + callColor + ';padding:1px 0;white-space:nowrap;">';
      callsHtml += inBadge + call.timeAgo + ' <span style="opacity:0.7;">(' + call.timeFormatted + ')</span> ' + vmBadge + '<strong>' + call.durationDisplay + '</strong></div>';
    }
  } else {
    callsHtml = '<div style="font-size:10px;color:' + callColor + ';font-weight:600;">Never</div>';
  }
  
  // 
  // BUILD SMS HTML - v3.4: Show ALL attempts, mark failures
  // 
  var smsHtml = '';
  
  // v3.4: Determine background color based on SMS history
  var smsBg = "#fee2e2";  // Default red (no SMS)
  var smsColor = "#991b1b";
  
  if (rc.last5SMS && rc.last5SMS.length > 0) {
    // Check if any SMS succeeded
    var hasSuccess = false;
    for (var k = 0; k < rc.last5SMS.length; k++) {
      if (!rc.last5SMS[k].isFailed) {
        hasSuccess = true;
        break;
      }
    }
    
    if (hasSuccess) {
      smsBg = "#dbeafe";  // Blue - has successful SMS
      smsColor = "#1e40af";
    } else {
      smsBg = "#fef3c7";  // Yellow - all failed (attempts made but failed)
      smsColor = "#92400e";
    }
    
    for (var j = 0; j < rc.last5SMS.length; j++) {
      var sms = rc.last5SMS[j];
      
      // v3.4: Show different indicator for failed vs successful
      var statusIndicator;
      if (sms.isFailed) {
        statusIndicator = '<span style="display:inline-block;background:#ef4444;color:#fff;font-size:8px;font-weight:700;padding:1px 4px;border-radius:3px;margin-left:4px;">FAILED</span>';
      } else {
        statusIndicator = '<span style="color:#10b981;margin-left:4px;">[OK]</span>';
      }
      
      smsHtml += '<div style="font-size:10px;color:' + smsColor + ';padding:1px 0;white-space:nowrap;">';
      smsHtml += sms.timeAgo + ' <span style="opacity:0.7;">(' + sms.timeFormatted + ')</span>' + statusIndicator + '</div>';
    }
  } else {
    smsHtml = '<div style="font-size:10px;color:' + smsColor + ';font-weight:600;">Never</div>';
  }
  
  // 
  // BUILD PROCESS DIRECTION (based on trackerType)
  // 
  var processDirection = getProcessDirection_(lead.trackerType, rc);
  
  // 
  // RENDER THREE BOXES: CALLS | SMS | PROCESS DIRECTION
  // 
  html += '<div style="display:flex;gap:8px;flex-wrap:wrap;">';
  
  // CALLS BOX (green)
  html += '<div style="background:' + callBg + ';border-radius:6px;padding:8px 10px;flex:1;min-width:120px;">';
  html += '<div style="font-size:9px;color:' + callColor + ';font-weight:700;margin-bottom:4px;">[PHONE] LAST 5 CALLS</div>';
  html += callsHtml;
  html += '</div>';
  
  // SMS BOX (blue/yellow/red based on status)
  html += '<div style="background:' + smsBg + ';border-radius:6px;padding:8px 10px;flex:1;min-width:120px;">';
  html += '<div style="font-size:9px;color:' + smsColor + ';font-weight:700;margin-bottom:4px;">[MSG] LAST 5 SMS</div>';
  html += smsHtml;
  html += '</div>';
  
  // PROCESS DIRECTION BOX (purple/action)
  html += '<div style="background:' + processDirection.bgColor + ';border-radius:6px;padding:8px 10px;flex:1;min-width:120px;border:2px solid ' + processDirection.borderColor + ';">';
  html += '<div style="font-size:9px;color:' + processDirection.textColor + ';font-weight:700;margin-bottom:4px;">[TARGET] PROCESS DIRECTION</div>';
  html += '<div style="font-size:12px;color:' + processDirection.textColor + ';font-weight:800;margin-bottom:4px;">' + processDirection.icon + ' ' + processDirection.action + '</div>';
  html += '<div style="font-size:9px;color:' + processDirection.textColor + ';opacity:0.8;">Source: ' + processDirection.source + '</div>';
  html += '</div>';
  
  html += '</div>'; // end flex container
  html += '</div>'; // end padding container
  return html;
}