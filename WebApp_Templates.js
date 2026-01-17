/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER PRO — REP WEB APP
 * File: WebApp_Templates.gs
 *
 * SMS templates with personalization engine.
 * Provides pre-built templates that can be customized per lead.
 *
 * Template Variables:
 *   {first_name} - Lead's first name
 *   {last_name} - Lead's last name
 *   {move_date} - Formatted move date
 *   {days_until} - Days until move
 *   {from_city} - Origin city
 *   {from_state} - Origin state
 *   {to_city} - Destination city
 *   {to_state} - Destination state
 *   {rep_name} - Rep's display name
 *   {rep_first} - Rep's first name
 *   {source} - Lead source
 *
 * v1.0 — Production Build
 **************************************************************/

/* ════════════════════════════════════════════════════════════════
 * TEMPLATE DEFINITIONS
 * ════════════════════════════════════════════════════════════════ */

var WEBAPP_TEMPLATES = [
  {
    id: 'initial_friendly',
    name: 'Initial Outreach (Friendly)',
    category: 'initial',
    icon: '👋',
    description: 'First contact - warm and welcoming',
    template: "Hi {first_name}! This is {rep_first} from Safe Ship Moving. I saw you're planning a move{move_context}. I'd love to help make your move stress-free! Would you have a few minutes for a quick quote?",
    tags: ['initial', 'friendly']
  },
  {
    id: 'initial_professional',
    name: 'Initial Outreach (Professional)',
    category: 'initial',
    icon: '💼',
    description: 'First contact - professional tone',
    template: "Hello {first_name}, this is {rep_first} with Safe Ship Moving Services. I'm reaching out regarding your upcoming move{move_context}. I'd be happy to provide you with a competitive quote. When would be a good time to discuss your moving needs?",
    tags: ['initial', 'professional']
  },
  {
    id: 'follow_up',
    name: 'Follow-Up',
    category: 'followup',
    icon: '🔄',
    description: 'Following up on previous contact',
    template: "Hi {first_name}, just checking in on your upcoming move{move_context}. Have you had a chance to think about your moving options? I'm here to answer any questions and help you get the best rate. - {rep_first}, Safe Ship Moving",
    tags: ['followup']
  },
  {
    id: 'quote_reminder',
    name: 'Quote Reminder',
    category: 'followup',
    icon: '💰',
    description: 'Remind about quote already sent',
    template: "Hi {first_name}! I wanted to follow up on the quote I sent for your move{move_context}. Do you have any questions? I'm happy to go over the details or adjust anything to fit your needs. - {rep_first}",
    tags: ['followup', 'quote']
  },
  {
    id: 'urgent_soon',
    name: 'Urgent - Move Soon',
    category: 'urgent',
    icon: '⚡',
    description: 'For moves within 3 days',
    template: "Hi {first_name}! I noticed your move is coming up very soon{move_context}. We still have availability and can help make this happen smoothly. Give me a call or reply here and I'll get you taken care of right away! - {rep_first}",
    tags: ['urgent', 'soon']
  },
  {
    id: 'same_day',
    name: 'Same Day Move',
    category: 'urgent',
    icon: '🚨',
    description: 'For same-day moves',
    template: "Hi {first_name}! I see you need help with a move TODAY. Good news - we have a crew available! Let me know ASAP and I'll get everything set up for you. Call or text me back right now. - {rep_first}, Safe Ship",
    tags: ['urgent', 'same-day']
  },
  {
    id: 'final_attempt',
    name: 'Final Attempt',
    category: 'closing',
    icon: '🏁',
    description: 'Last try before closing lead',
    template: "Hi {first_name}, I've tried reaching you a few times about your move{move_context}. I don't want you to miss out on our current rates. This will be my last message - just reply 'YES' if you'd still like help, or 'STOP' if you've made other arrangements. - {rep_first}",
    tags: ['closing', 'final']
  },
  {
    id: 'price_match',
    name: 'Price Match Offer',
    category: 'closing',
    icon: '🏆',
    description: 'Competitive pricing message',
    template: "Hi {first_name}! Still looking for the best deal on your move{move_context}? We offer price matching and flexible scheduling. Let me know what other quotes you've received and I'll see if we can do better! - {rep_first}, Safe Ship",
    tags: ['closing', 'price']
  },
  {
  id: 'my_custom',
  name: 'My Custom Template',
  category: 'custom',
  icon: '⭐',
  description: 'My custom message',
  template: "Hi {first_name}! Your move on {move_date} is coming up...",
  tags: ['custom']
  }  
  ];

/* ════════════════════════════════════════════════════════════════
 * PUBLIC FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get all available templates
 * @returns {Array} Template definitions (without full template text for security)
 */
function WebApp_Templates_getAll() {
  return WEBAPP_TEMPLATES.map(function(t) {
    return {
      id: t.id,
      name: t.name,
      category: t.category,
      icon: t.icon,
      description: t.description,
      tags: t.tags
    };
  });
}

/**
 * Generate a personalized SMS draft
 * @param {Object} lead - Lead data object
 * @param {string} templateId - Template ID to use
 * @param {Object} context - Rep context
 * @param {Object} options - Additional options (tone, custom fields)
 * @returns {Object} { message, charCount, segments, preview }
 */
function WebApp_Templates_generate(lead, templateId, context, options) {
  options = options || {};
  
  // Find template
  var template = WEBAPP_TEMPLATES.find(function(t) { return t.id === templateId; });
  
  if (!template) {
    throw new Error('Template not found: ' + templateId);
  }
  
  // Build variable map
  var vars = WebApp_Templates_buildVarMap_(lead, context, options);
  
  // Replace variables in template
  var message = WebApp_Templates_replaceVars_(template.template, vars);
  
  // Clean up extra spaces
  message = message.replace(/\s+/g, ' ').trim();
  
  // Calculate SMS stats
  var charCount = message.length;
  var segments = Math.ceil(charCount / 160);  // Standard SMS segment size
  
  return {
    message: message,
    charCount: charCount,
    segments: segments,
    templateId: templateId,
    templateName: template.name,
    preview: message.substring(0, 100) + (message.length > 100 ? '...' : '')
  };
}

/**
 * Generate custom message (no template, just personalization)
 * @param {string} customText - Custom message text with variables
 * @param {Object} lead - Lead data
 * @param {Object} context - Rep context
 * @returns {Object} Generated message details
 */
function WebApp_Templates_generateCustom(customText, lead, context) {
  var vars = WebApp_Templates_buildVarMap_(lead, context, {});
  var message = WebApp_Templates_replaceVars_(customText, vars);
  
  message = message.replace(/\s+/g, ' ').trim();
  
  return {
    message: message,
    charCount: message.length,
    segments: Math.ceil(message.length / 160),
    templateId: 'custom',
    templateName: 'Custom Message'
  };
}

/* ════════════════════════════════════════════════════════════════
 * INTERNAL FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Build variable replacement map
 * @param {Object} lead - Lead data
 * @param {Object} context - Rep context
 * @param {Object} options - Additional options
 * @returns {Object} Variable map
 */
function WebApp_Templates_buildVarMap_(lead, context, options) {
  // Format move date
  var moveDate = '';
  var daysUntil = null;
  
  if (lead.moveDate) {
    var md = lead.moveDate instanceof Date ? lead.moveDate : new Date(lead.moveDate);
    if (!isNaN(md.getTime())) {
      var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      moveDate = months[md.getMonth()] + ' ' + md.getDate();
      
      var today = new Date();
      today.setHours(0, 0, 0, 0);
      md.setHours(0, 0, 0, 0);
      daysUntil = Math.round((md - today) / (24 * 60 * 60 * 1000));
    }
  }
  
  // Build move context phrase
  var moveContext = '';
  if (moveDate) {
    if (daysUntil === 0) {
      moveContext = ' today';
    } else if (daysUntil === 1) {
      moveContext = ' tomorrow';
    } else if (daysUntil !== null && daysUntil > 0 && daysUntil <= 7) {
      moveContext = ' on ' + moveDate + ' (just ' + daysUntil + ' days away!)';
    } else if (moveDate) {
      moveContext = ' on ' + moveDate;
    }
  }
  
  // Add location context if available
  if (lead.fromState && lead.toState && lead.fromState !== lead.toState) {
    moveContext += ' from ' + (lead.fromCity || lead.fromState) + ' to ' + (lead.toCity || lead.toState);
  } else if (lead.fromCity && lead.toCity && lead.fromCity !== lead.toCity) {
    moveContext += ' from ' + lead.fromCity + ' to ' + lead.toCity;
  }
  
  return {
    '{first_name}': WebApp_Templates_capitalize_(lead.firstName) || 'there',
    '{last_name}': WebApp_Templates_capitalize_(lead.lastName) || '',
    '{full_name}': WebApp_Templates_capitalize_((lead.firstName || '') + ' ' + (lead.lastName || '')).trim() || 'there',
    '{move_date}': moveDate || 'your upcoming move',
    '{days_until}': daysUntil !== null ? String(daysUntil) : '',
    '{move_context}': moveContext,
    '{from_city}': lead.fromCity || '',
    '{from_state}': lead.fromState || '',
    '{to_city}': lead.toCity || '',
    '{to_state}': lead.toState || '',
    '{from_location}': (lead.fromCity ? lead.fromCity + ', ' : '') + (lead.fromState || ''),
    '{to_location}': (lead.toCity ? lead.toCity + ', ' : '') + (lead.toState || ''),
    '{rep_name}': context.displayName || context.firstName || 'Your Rep',
    '{rep_first}': context.firstName || context.displayName || 'Your Rep',
    '{source}': lead.source || '',
    '{job_no}': lead.jobNo || ''
  };
}

/**
 * Replace variables in template text
 * @param {string} text - Template text
 * @param {Object} vars - Variable map
 * @returns {string} Text with variables replaced
 */
function WebApp_Templates_replaceVars_(text, vars) {
  var result = text;
  
  for (var key in vars) {
    if (vars.hasOwnProperty(key)) {
      // Use global replace
      result = result.split(key).join(vars[key] || '');
    }
  }
  
  return result;
}

/**
 * Capitalize first letter of each word
 * @param {string} str - String to capitalize
 * @returns {string} Capitalized string
 */
function WebApp_Templates_capitalize_(str) {
  if (!str) return '';
  return String(str).toLowerCase().replace(/\b\w/g, function(l) { return l.toUpperCase(); });
}

/* ════════════════════════════════════════════════════════════════
 * TEMPLATE CATEGORIES (for UI grouping)
 * ════════════════════════════════════════════════════════════════ */

function WebApp_Templates_getCategories() {
  return [
    { id: 'initial', name: 'Initial Contact', icon: '👋' },
    { id: 'followup', name: 'Follow-Up', icon: '🔄' },
    { id: 'urgent', name: 'Urgent / Time-Sensitive', icon: '⚡' },
    { id: 'closing', name: 'Closing', icon: '🏁' }
  ];
}