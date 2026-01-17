/**************************************************************
 * Config.gs — Safe Ship Lead Contact Alert Bot (Enterprise)
 * Spec-compliant, deterministic, no silent failures.
 **************************************************************/

function SS_getConfig_() {
  const props = PropertiesService.getScriptProperties();

  return {
    TIMEZONE: 'America/New_York',

    // ===== Admin + Test Mode =====
    ADMIN_SUMMARY_EMAILS: ['Christopher@safeshipmoving.com', 'Jake@safeshipmoving.com'],
    ADMIN_TEST_EMAIL: String(props.getProperty('ADMIN_TEST_EMAIL') || 'Christopher@safeshipmoving.com').trim(),

    // ===== Slack =====
    SLACK_BOT_TOKEN: String(props.getProperty('SLACK_BOT_TOKEN') || '').trim(),
    SLACK_LOOKUP_CACHE_HOURS: Number(props.getProperty('SLACK_LOOKUP_CACHE_HOURS') || 24),

    // ===== Sheet Names =====
    SHEETS: {
      SMS: 'SMS TRACKER',
      CALLVM: 'CALL & VOICEMAIL TRACKER',
      P1CALLVM: 'PRIORITY 1 CALL & VOICEMAIL TRACKER',
      CONTACTED: 'CONTACTED LEADS',
      SAME_DAY: 'SAME DAY TRANSFERS',
      SALES_ROSTER: 'Sales_Roster',
      TEAM_ROSTER: 'Team_Roster',
      NOTIF_LOG: 'Notification_Log',
      ADMIN_REPORT_LOG: 'Admin_Report_Log'
    },

    // ===== Tracker Order (ALWAYS) =====
    TRACKER_ORDER: [
      'SMS TRACKER',
      'CALL & VOICEMAIL TRACKER',
      'PRIORITY 1 CALL & VOICEMAIL TRACKER',
      'CONTACTED LEADS',
      'SAME DAY TRANSFERS'
    ],

    // ===== Column Maps (absolute sheet columns; spec exact) =====
    COLMAP: {
      SMS: {
        scanRangeA1: 'F3:M',
        dataStartRow: 3,
        jobColLetter: 'I',
        userColLetter: 'J',
        excludeColLetter: 'M',
        priorityCellA1: 'B7'
      },
      CALLVM: {
        scanRangeA1: 'F3:M',
        dataStartRow: 3,
        jobColLetter: 'I',
        userColLetter: 'J',
        excludeColLetter: 'M',
        priorityCellA1: 'A8'
      },
      P1CALLVM: {
        scanRangeA1: 'F3:M',
        dataStartRow: 3,
        jobColLetter: 'I',
        userColLetter: 'J',
        excludeColLetter: 'M',
        priorityCellA1: 'A8'
      },
      CONTACTED: {
        scanRangeA1: 'E2:L',
        dataStartRow: 2,
        jobColLetter: 'H',
        userColLetter: 'I',
        excludeColLetter: 'L',
        priorityLabel: 'Already Quoted' // Spec fixed label
      },
      SAME_DAY: {
        headerRow: 2,
        dataStartRow: 4,
        scanRangeA1: 'A4:F',
        // A=DATE, B=JOB#, C=REP, D=move window, E=BOOKING AGENT, F=$$
        bookedAgentColIndexInRange: 5, // E within A:F
        bookedAmountColIndexInRange: 6  // F within A:F
      }
    },

    // ===== Output Limits =====
    LIMITS: {
      SLACK_PER_SECTION: 15,
      EMAIL_PER_SECTION: 25
    },

    // ===== Theme Tokens (shared across Gmail + Slack) =====
    THEME: SS_getTheme_()
  };
}

function SS_getTheme_() {
  return {
    brand: {
      name: 'SAFE SHIP',
      product: 'Lead Contact Alert Bot'
    },
    colors: {
      primary: '#4F81BD',
      text: '#111827',
      subtext: '#374151',
      muted: '#6B7280',
      bg: '#F6F8FB',
      card: '#FFFFFF',
      border: '#E5E7EB',
      success: '#16A34A',
      warn: '#D97706',
      error: '#DC2626'
    },
    badges: {
      NEEDS_ACTION: { label: 'NEEDS ACTION', icon: '⚡', color: '#D97706' },
      FOLLOW_UP: { label: 'FOLLOW-UP', icon: '📌', color: '#4F81BD' },
      OK: { label: 'OK', icon: '✅', color: '#16A34A' },
      ERROR: { label: 'ERROR', icon: '⛔', color: '#DC2626' }
    }
  };
}

/** Deterministic "Today" in America/New_York */
function SS_todayKey_() {
  const cfg = SS_getConfig_();
  return Utilities.formatDate(new Date(), cfg.TIMEZONE, 'yyyy-MM-dd');
}