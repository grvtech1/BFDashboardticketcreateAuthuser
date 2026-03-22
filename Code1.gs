/* ==================================================
   🏢 ENTERPRISE CONFIGURATION v10.0
   BillFree TechSupport Ops - Production Ready
   Phase 1: Critical Security & Stability Fixes
   ================================================== */

/**
 * 🔧 SYSTEM CONFIGURATION
 * Centralized config for easy management and deployment
 */
const CONFIG = Object.freeze({
  // Application Settings
  APP_TITLE: "BillFree TechSupport Ops v10.0 PRO",
  APP_VERSION: "10.0.0",
  SHEET_NAME: "IT Tracker 26",
  AUDIT_SHEET_NAME: "Audit Log",
  CALL_LOG_SHEET_NAME: "Call Log",
  
  // Cache & Performance
  CACHE_TTL_SECONDS: 300,           // 5 minutes
  TICKET_INDEX_TTL: 600,            // 10 minutes
  MAX_BATCH_SIZE: 100,
  
  // Rate Limiting
  RATE_LIMIT_WINDOW_SECONDS: 60,
  RATE_LIMIT_MAX_REQUESTS: 30,
  
  // Business Rules
  MIN_CLOSURE_DAYS: 7,
  CRITICAL_AGE_DAYS: 15,
  WARNING_AGE_DAYS: 7,
  MIN_FOLLOWUP_DAYS: 7,
  
  // Pagination
  DEFAULT_PAGE_SIZE: 100,
  MAX_PAGE_SIZE: 500,
  
  // Lock Settings (Phase 1 - New)
  LOCK_TIMEOUT_MS: 5000,
  MAX_RETRIES: 3
});

/**
 * ⚡ P1-7: SPREADSHEET HANDLE SINGLETON
 * Each SpreadsheetApp.getActiveSpreadsheet() costs ~50-100ms.
 * This lazy singleton caches the handle for the request lifecycle.
 * GAS resets module-level state between requests, so no staleness risk.
 */
let _ssInstance = null;
function getSpreadsheet_() {
  if (!_ssInstance) _ssInstance = SpreadsheetApp.getActiveSpreadsheet();
  return _ssInstance;
}

/**
 * ⚡ P1-FIX6: SESSION EMAIL SINGLETON
 * Session.getActiveUser().getEmail() costs ~50-100ms per call.
 * This caches it for the request lifecycle (GAS resets module state between requests).
 *
 * Fallback chain:
 *   1. Session.getActiveUser().getEmail()  — works when user has authorized
 *   2. Session.getEffectiveUser().getEmail() — returns script owner for "Execute as: Me"
 *
 * GAS web apps deployed as "Execute as: Me" may return empty from getActiveUser()
 * in google.script.run calls, but getEffectiveUser() always returns the owner email.
 */
let _cachedSessionEmail = null;
let _requestUser = null;

function setRequestUser_(email) {
  _requestUser = normalizeEmail_(email);
}

function getSessionEmail_() {
  if (_requestUser) return _requestUser;

  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (_) {
    return '';
  }
}

// Legacy compatibility
const SHEET_NAME = CONFIG.SHEET_NAME;
const APP_TITLE = CONFIG.APP_TITLE;

/**
 * 🔐 ADMIN CONFIGURATION — declared early so initializeAgentPhones can use it
 * (Duplicate const below is removed; this is now the single definition)
 */
const ADMIN_EMAILS = Object.freeze([
  "gaurav.pal@billfree.in"
]);

// Legacy compatibility alias
const ADMIN_EMAIL = ADMIN_EMAILS[0];

// Pre-normalized admin list for consistent case-insensitive checks.
const ADMIN_EMAILS_NORMALIZED = Object.freeze(
  ADMIN_EMAILS
    .map(email => String(email || '').toLowerCase().trim())
    .filter(Boolean)
);

function normalizeEmail_(email) {
  return String(email || '').toLowerCase().trim();
}

function isAdminEmail_(email) {
  return ADMIN_EMAILS_NORMALIZED.includes(normalizeEmail_(email));
}

/**
 * 👥 AGENT DIRECTORY
 * Contact info (PII) now stored in PropertiesService for security
 * Use getAgentPhone() to retrieve phone numbers securely
 */
const AGENT_DIRECTORY = Object.freeze({
  "Suraj":        { email: "suraj.billfree2@gmail.com",      role: "agent" },
  "Veer Bahadur": { email: "veer.billfree@gmail.com",        role: "agent" },
  "Neeraj Kumar": { email: "neerajkumar.billfree@gmail.com", role: "agent" },
  "Manjeet":      { email: "manjeetkashyap.billfree@gmail.com", role: "agent" },
  "Admin":        { email: "gaurav.pal@billfree.in",         role: "admin" }
});

/**
 * 🔍 LOOK UP AN AGENT BY EMAIL
 * Returns the agent object { name, email, role } or null.
 * Called by getCurrentUserEmail() and doGet() to resolve agentName.
 * @param {string} email
 * @returns {{ name: string, email: string, role: string }|null}
 */
function getAgentByEmail(email) {
  if (!email) return null;
  const normalized = email.toLowerCase().trim();
  for (const [name, info] of Object.entries(AGENT_DIRECTORY)) {
    if (info.email && info.email.toLowerCase() === normalized) {
      return { name, email: info.email, role: info.role || 'agent' };
    }
  }
  return null;
}

/**
 * 📋 GET AGENT LIST (JSON)
 * Returns all agents from AGENT_DIRECTORY as a structured JSON response.
 * Called by:
 *   - frontend google.script.run.getAgentList()
 *   - doGet() to inject agents into the HTML template at render time
 */
function getAgentList() {
  try {
    const agents = Object.entries(AGENT_DIRECTORY).map(([name, info]) => ({
      name:  name,
      email: info.email || '',
      role:  info.role  || 'agent'
    }));
    return JSON.stringify({
      success: true,
      agents:  agents,
      count:   agents.length
    });
  } catch (e) {
    Logger.log('getAgentList error: ' + e.toString());
    return JSON.stringify({ success: false, agents: [], error: e.toString() });
  }
}



/**
 * 🔐 SECURE PII RETRIEVAL
 * Phone numbers stored in PropertiesService, not in source code
 * Run initializeAgentPhones() once to set up phone numbers
 */
function getAgentPhone(agentName) {
  try {
    const props = PropertiesService.getScriptProperties();
    const key = `AGENT_PHONE_${agentName.replace(/\s/g, '_').toUpperCase()}`;
    return props.getProperty(key) || null;
  } catch (e) {
    Logger.log('Error retrieving agent phone: ' + e.toString());
    return null;
  }
}

/**
 * 🔧 SECURE PHONE NUMBER INITIALIZATION (Phase 1 Security Fix)
 * 
 * ⚠️ IMPORTANT: Phone numbers are NO LONGER stored in source code.
 * Run this function from the Apps Script editor with the phoneNumbers parameter.
 * 
 * Usage:
 *   initializeAgentPhones({
 *     'MANJEET': '91XXXXXXXXXX',
 *     'SURAJ': '91XXXXXXXXXX', 
 *     'VEER_BAHADUR': '91XXXXXXXXXX'
 *   });
 * 
 * @param {Object} phoneNumbers - Map of agent name (uppercase, underscore for spaces) to phone number
 * @returns {string} Status message
 */
function initializeAgentPhones(phoneNumbers) {
  // Security check - only allow from Apps Script editor, not web app
  const userEmail = Session.getActiveUser().getEmail();
  if (!isAdminEmail_(userEmail)) {
    return 'Error: Only administrators can initialize phone numbers';
  }
  
  if (!phoneNumbers || typeof phoneNumbers !== 'object') {
    return `
╔═══════════════════════════════════════════════════════════════╗
║  🔐 SECURE PHONE NUMBER INITIALIZATION                        ║
╠═══════════════════════════════════════════════════════════════╣
║                                                               ║
║  Phone numbers are NOT stored in source code for security.    ║
║                                                               ║
║  To initialize, run from Apps Script editor:                  ║
║                                                               ║
║    initializeAgentPhones({                                    ║
║      'MANJEET': '91XXXXXXXXXX',                               ║
║      'SURAJ': '91XXXXXXXXXX',                                 ║
║      'VEER_BAHADUR': '91XXXXXXXXXX'                           ║
║    });                                                        ║
║                                                               ║
║  Replace X with actual phone digits.                          ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
    `;
  }
  
  const props = PropertiesService.getScriptProperties();
  const results = [];
  
  for (const [agentKey, phone] of Object.entries(phoneNumbers)) {
    const sanitizedKey = agentKey.toUpperCase().replace(/\s/g, '_');
    const propKey = `AGENT_PHONE_${sanitizedKey}`;
    
    // Validate phone format (basic check)
    if (!/^\d{10,15}$/.test(phone.replace(/\D/g, ''))) {
      results.push(`⚠️ ${agentKey}: Invalid phone format (skipped)`);
      continue;
    }
    
    props.setProperty(propKey, phone.replace(/\D/g, ''));
    results.push(`✅ ${agentKey}: Phone stored securely`);
  }
  
  // Log this security-sensitive action
  logAuditEvent('PHONE_NUMBERS_INITIALIZED', null, {
    agentCount: Object.keys(phoneNumbers).length,
    initializedBy: userEmail
  }, 'INFO');
  
  return '🔐 Phone Initialization Complete:\n' + results.join('\n');
}

// ✅ getAgentList() is defined at line ~108, getAgentByEmail() at line ~90.
// Do NOT re-declare them here — GAS uses the last definition, causing silent divergence bugs.

// ✅ ADMIN_EMAILS & ADMIN_EMAIL are declared above AGENT_DIRECTORY (line ~60)
// to avoid temporal dead zone when referenced inside initializeAgentPhones().


/**
 * 🔒 ROLE-BASED ACCESS CONTROL
 */
const ROLES = Object.freeze({
  ADMIN: 'admin',
  MANAGER: 'manager',
  AGENT: 'agent',
  VIEWER: 'viewer',
  SYSTEM: 'system'
});

const PERMISSIONS = Object.freeze({
  CLOSE_TICKET: [ROLES.ADMIN, ROLES.MANAGER],
  DELETE_TICKET: [ROLES.ADMIN],
  VIEW_AUDIT: [ROLES.ADMIN, ROLES.MANAGER],
  MANAGE_USERS: [ROLES.ADMIN],
  UPDATE_TICKET: [ROLES.ADMIN, ROLES.MANAGER, ROLES.AGENT],
  VIEW_ANALYTICS: [ROLES.ADMIN, ROLES.MANAGER, ROLES.AGENT],
  CREATE_TICKET: [ROLES.ADMIN, ROLES.MANAGER, ROLES.AGENT],
  CALL_LOG_EVENT: [ROLES.ADMIN, ROLES.MANAGER, ROLES.AGENT],
  EXPORT_TICKETS: [ROLES.ADMIN, ROLES.MANAGER],
  EXPORT_REPORT: [ROLES.ADMIN, ROLES.MANAGER],
  EXPORT_HISTORY: [ROLES.ADMIN, ROLES.MANAGER]
});

/**
 * 📊 STATUS CONFIGURATION
 */
const VALID_STATUSES = Object.freeze(['Not Completed', 'Completed', 'Closed', 'Pending', 'In Progress', "Can't Do"]);

const STATUS_ENUM = Object.freeze({
  NOT_COMPLETED: 'Not Completed',
  PENDING: 'Pending',
  IN_PROGRESS: 'In Progress',
  COMPLETED: 'Completed',
  CLOSED: 'Closed',
  CANT_DO: "Can't Do"
});

/**
 * 🚨 ERROR CODES
 */
const ERROR_CODES = Object.freeze({
  RATE_LIMITED: 'E001',
  UNAUTHORIZED: 'E002',
  NOT_FOUND: 'E003',
  VALIDATION_FAILED: 'E004',
  SHEET_ERROR: 'E005',
  LOCK_TIMEOUT: 'E006',
  INVALID_STATUS: 'E007',
  INSUFFICIENT_PERMISSIONS: 'E008',
  UNKNOWN_ERROR: 'E999'
});

/**
 * 💬 USER-FRIENDLY ERROR MESSAGES
 * Translates technical error codes into clear, actionable messages for users.
 * Technical codes are still logged for debugging purposes.
 */
const USER_FRIENDLY_ERRORS = Object.freeze({
  'E001': 'Too many requests. Please wait a moment and try again.',
  'E002': 'Please sign in to continue.',
  'E003': 'The item you\'re looking for doesn\'t exist or has been removed.',
  'E004': 'Please check your input and try again.',
  'E005': 'Unable to access the database. Please refresh the page.',
  'E006': 'The system is busy. Please try again in a few seconds.',
  'E007': 'Please select a valid status option.',
  'E008': 'You don\'t have permission for this action. Contact your administrator.',
  'E999': 'Something went wrong. Please try again or refresh the page.'
});

/**
 * 🔄 GET USER-FRIENDLY ERROR MESSAGE
 * @param {string} errorCode - Technical error code (E001, E002, etc.)
 * @param {string} fallbackMessage - Original message if no friendly version exists
 * @returns {string} User-friendly error message
 */
function getUserFriendlyError(errorCode, fallbackMessage) {
  // Extract error code if embedded in message like "[E001]..."
  const codeMatch = errorCode.match(/E\d{3}/);
  const code = codeMatch ? codeMatch[0] : errorCode;
  return USER_FRIENDLY_ERRORS[code] || fallbackMessage || 'An error occurred. Please try again.';
}

/* ==================================================
   🔧 PHASE 2: API RESPONSE WRAPPER & CORRELATION IDS

   ================================================== */

/**
 * 📝 Generate unique correlation ID for request tracing
 * Format: yyMMddHHmmss-random4chars
 */
function generateCorrelationId() {
  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Kolkata', 'yyMMddHHmmss');
  const random = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${timestamp}-${random}`;
}

/**
 * ✅ STANDARDIZED API SUCCESS RESPONSE
 * @param {Object} data - Response payload
 * @param {string} correlationId - Request tracking ID
 * @param {Object} meta - Optional metadata (pagination, etc.)
 */
function apiSuccess(data, correlationId = null, meta = {}) {
  const response = {
    success: true,
    data: data,
    correlationId: correlationId || generateCorrelationId(),
    timestamp: new Date().toISOString(),
    version: CONFIG.APP_VERSION,
    ...meta
  };
  return JSON.stringify(response);
}

/**
 * ❌ STANDARDIZED API ERROR RESPONSE
 * @param {string} errorCode - From ERROR_CODES enum
 * @param {string} message - Technical error message (for logging)
 * @param {string} correlationId - Request tracking ID
 * @param {Object} details - Optional error details
 */
function apiError(errorCode, message, correlationId = null, details = {}) {
  const cid = correlationId || generateCorrelationId();
  
  // Log technical error for debugging (preserves error codes)
  Logger.log(`[${cid}] API Error: [${errorCode}] ${message}`);
  if (Object.keys(details).length > 0) {
    Logger.log(`[${cid}] Error Details: ${JSON.stringify(details)}`);
  }
  
  // Return user-friendly message to frontend (hides technical codes)
  const userMessage = getUserFriendlyError(errorCode, message);
  
  const response = {
    success: false,
    error: {
      code: errorCode,                    // Keep code for frontend error handling logic
      message: userMessage                // User-friendly message shown to user
    },
    correlationId: cid,
    timestamp: new Date().toISOString(),
    version: CONFIG.APP_VERSION
  };
  return JSON.stringify(response);
}


/**
 * 🔄 WRAP EXISTING FUNCTION FOR CORRELATION TRACING
 * Adds correlation IDs and standardized error handling
 */
function withCorrelation(fn) {
  return function(...args) {
    const correlationId = generateCorrelationId();
    const startTime = Date.now();
    
    try {
      Logger.log(`[${correlationId}] Starting: ${fn.name}`);
      const result = fn.apply(this, args);
      const duration = Date.now() - startTime;
      Logger.log(`[${correlationId}] Completed: ${fn.name} in ${duration}ms`);
      return result;
    } catch (e) {
      const duration = Date.now() - startTime;
      Logger.log(`[${correlationId}] ERROR in ${fn.name} after ${duration}ms: ${e.toString()}`);
      return apiError(ERROR_CODES.UNKNOWN_ERROR, e.message, correlationId, {
        function: fn.name,
        duration: duration
      });
    }
  };
}

/**
 * 🔄 NORMALIZE STATUS WITH DEFAULT
 * Converts any status string to a canonical STATUS_ENUM value.
 * Returns NOT_COMPLETED for null/undefined/unrecognized inputs.
 * Use this when a guaranteed valid status is needed (e.g., reading sheet data).
 * @param {string|null} s - Raw status string
 * @returns {string} Canonical status from STATUS_ENUM (never null)
 */
function normalizeStatusWithDefault(s) {
  if (!s) return STATUS_ENUM.NOT_COMPLETED;
  const v = String(s).toLowerCase();
  if (v === 'completed') return STATUS_ENUM.COMPLETED;
  if (v === 'closed') return STATUS_ENUM.CLOSED;
  // ✅ FIX: removed duplicate v.includes("can't") check (smart-quote was a dup of ascii)
  if (v.includes('cant') || v.includes("can't")) return STATUS_ENUM.CANT_DO;
  if (v === 'pending') return STATUS_ENUM.PENDING;
  if (v === 'in progress') return STATUS_ENUM.IN_PROGRESS;
  return STATUS_ENUM.NOT_COMPLETED;
}

/**
 * 🔍 PARSE STATUS OR NULL
 * Strictly parses user-supplied status input to a canonical STATUS_ENUM value.
 * Returns null for unrecognized values (useful for input validation).
 * Use this when validating user-submitted status changes.
 * @param {string|null} s - Raw status input from user
 * @returns {string|null} Canonical status from STATUS_ENUM, or null if invalid
 */
function parseStatusOrNull(s) {
  if (s === null || s === undefined) return null;
  const v = String(s).trim().toLowerCase();
  if (v === 'not completed' || v === 'notcompleted') return STATUS_ENUM.NOT_COMPLETED;
  if (v === 'completed') return STATUS_ENUM.COMPLETED;
  if (v === 'closed') return STATUS_ENUM.CLOSED;
  if (v === 'pending') return STATUS_ENUM.PENDING;
  if (v === 'in progress' || v === 'inprogress') return STATUS_ENUM.IN_PROGRESS;
  if (v === "can't do" || v === 'cant do') return STATUS_ENUM.CANT_DO;
  return null;
}


/* ==================================================
   🛡️ ENTERPRISE UTILITY FUNCTIONS
   ================================================== */

/**

 * 🚦 RATE LIMITING
 * Prevents abuse and ensures fair usage
 */
function rateLimitCheck(action) {
  const userEmail = getSessionEmail_() || 'anonymous';
  const cache = CacheService.getUserCache();
  const key = `rate_${action}_${userEmail.replace(/[@.]/g, '_')}`;
  
  const currentCount = parseInt(cache.get(key) || '0');
  
  if (currentCount >= CONFIG.RATE_LIMIT_MAX_REQUESTS) {
    throw new Error(`[${ERROR_CODES.RATE_LIMITED}] Rate limit exceeded. Please wait ${CONFIG.RATE_LIMIT_WINDOW_SECONDS} seconds.`);
  }
  
  cache.put(key, String(currentCount + 1), CONFIG.RATE_LIMIT_WINDOW_SECONDS);
  return true;
}

/**
 * 🔐 CSRF PROTECTION
 * Prevents cross-site request forgery attacks
 */
function generateCSRFToken() {
  const token = Utilities.getUuid();
  const cache = CacheService.getUserCache();
  const timestamp = String(Date.now());
  cache.put('CSRF_TOKEN', token, 3600); // 1 hour TTL
  cache.put('CSRF_TOKEN_TS', timestamp, 3600);
  return token;
}

/**
 * Validate CSRF token (delegates to enhanced validator for age checks)
 * @param {string} token - CSRF token from client
 * @returns {boolean} Whether the token is valid
 */
function validateCSRFToken(token) {
  const result = validateCSRFTokenEnhanced(token);
  return result.valid;
}

function getCSRFToken() {
  try {
    const cache = CacheService.getUserCache();
    let token = cache.get('CSRF_TOKEN');
    if (!token) {
      token = generateCSRFToken();
    } else if (!cache.get('CSRF_TOKEN_TS')) {
      cache.put('CSRF_TOKEN_TS', String(Date.now()), 3600);
    }
    return JSON.stringify({ success: true, token: token });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 🛡️ CSRF TOKEN ENFORCEMENT (Phase 1 Security Fix)
 * Validates CSRF token for all mutating operations
 * @param {string} token - CSRF token from client
 * @throws {Error} If token is missing or invalid
 */
function requireCSRFToken(token) {
  if (!token) {
    logAuditEvent('CSRF_MISSING', null, { 
      message: 'CSRF token not provided' 
    }, 'WARNING');
    throw new Error(`[${ERROR_CODES.UNAUTHORIZED}] CSRF token required for this operation`);
  }
  
  const result = validateCSRFTokenEnhanced(token);
  if (!result.valid) {
    logAuditEvent('CSRF_INVALID', null, { 
      message: 'Invalid or expired CSRF token',
      reason: result.reason,
      providedToken: token.substring(0, 8) + '...' // Log partial for debugging
    }, 'WARNING');
    throw new Error(`[${ERROR_CODES.UNAUTHORIZED}] Invalid or expired CSRF token. Please refresh the page.`);
  }
  
  return true;
}

/**
 * 🔒 PERMISSION ENFORCEMENT (Critical Fix)
 * Checks if current user has required permission based on their role
 * @param {string} action - Permission key from PERMISSIONS constant
 * @throws {Error} If user lacks the required permission
 */
function requirePermission(action) {
  const userEmail = normalizeEmail_(getSessionEmail_());
  if (!userEmail) {
    throw new Error(`[${ERROR_CODES.UNAUTHORIZED}] Authentication required`);
  }
  
  // Admins always have all permissions (case-insensitive)
  if (isAdminEmail_(userEmail)) {
    return true;
  }
  
  // Determine user's role from agent directory
  const userRole = getUserRole(userEmail);
  
  // Get allowed roles for this action — FAIL CLOSED for unknown keys
  const allowedRoles = PERMISSIONS[action];
  if (!allowedRoles) {
    logAuditEvent('PERMISSION_CONFIG_ERROR', null, {
      user: userEmail,
      action: action,
      message: 'Unknown permission key requested'
    }, 'ERROR');
    throw new Error(`[${ERROR_CODES.INSUFFICIENT_PERMISSIONS}] Unknown permission: ${action}`);
  }
  
  // Check if user's role is in the allowed list
  if (!allowedRoles.includes(userRole)) {
    logAuditEvent('PERMISSION_DENIED', null, {
      user: userEmail,
      action: action,
      userRole: userRole,
      allowedRoles: allowedRoles.join(', ')
    }, 'WARNING');
    
    throw new Error(`[${ERROR_CODES.INSUFFICIENT_PERMISSIONS}] You don't have permission: ${action}`);
  }
  
  return true;
}

/**
 * 👤 GET USER ROLE
 * Determines the role for a given user email
 * @param {string} email - User's email address (optional, defaults to current user)
 * @returns {string} User's role from ROLES constant
 */
function getUserRole(email) {
  const userEmail = normalizeEmail_(email || getSessionEmail_());
  if (!userEmail) return ROLES.VIEWER;
  
  // Check if admin (case-insensitive)
  if (isAdminEmail_(userEmail)) {
    return ROLES.ADMIN;
  }
  
  // Check agent directory for specific role assignments (e.g., manager)
  for (const [name, info] of Object.entries(AGENT_DIRECTORY)) {
    if (info.email && info.email.toLowerCase() === userEmail) {
      return info.role || ROLES.AGENT;
    }
  }
  
  // Fail closed: unknown users are VIEWER until explicitly mapped.
  return ROLES.VIEWER;
}

/**
 * 👤 GET CURRENT USER IDENTITY
 * Returns current user identity for frontend initialization and forms.
 * ⚡ SECURE IDENTITY SYNC: Now accepts an optional ID token from Cloudflare/GSI.
 * If no token is provided, falls back to standard GAS Session (often restricted in iframes).
 * 
 * @param {string} [idToken] - Google ID Token from Cloudflare/GSI
 * @returns {string} JSON response with success, email, role, and isAdmin status
 */
function getCurrentUserEmail(idToken) {
  try {
    let email = '';
    
    // 1. Try to verify provided ID Token (Production-grade security)
    if (idToken && typeof idToken === 'string' && idToken.length > 32) {
      const verified = verifyGoogleToken_(idToken);
     if (verified && verified.email) {
  email = verified.email;

  // ✅ CRITICAL FIX: persist user for request
  setRequestUser_(email);
}
    }
    
    // 2. Fallback to standard Session if token verification failed or was missing
    if (!email) {
  return JSON.stringify({
    success: false,
    error: "User identity not verified"
  });
}
    
    // 3. Domain/Whitelist Validation (Security Hardening)
    const normalizedEmail = normalizeEmail_(email);
    const domain = normalizedEmail.split('@')[1];
    const allowedDomains = ['billfree.in', 'gmail.com'];
const isAllowedDomain = allowedDomains.includes(domain);
    const isWhitelisted = isAdminEmail_(normalizedEmail) || AGENT_EMAILS_NORMALIZED_.includes(normalizedEmail);
    
    // Fail-closed security for unauthorized users
    if (email && !isAllowedDomain && !isWhitelisted) {
       Logger.log(`[Security] Unauthorized access attempt: ${email}`);
       return apiError(ERROR_CODES.UNAUTHORIZED, 'Unauthorized domain or email.');
    }

    const role = getUserRole(email);
    const agent = getAgentByEmail(email);
    
    logAuditEvent('IDENTITY_SYNC_VERIFIED', null, {
      email: email,
      agentName: agent ? agent.name : 'Unknown',
      isVerified: !!idToken
    });

    return JSON.stringify({
      success: true,
      email: email,
      role: role,
      isAdmin: isAdminEmail_(email),
      agentName: agent ? agent.name : email.split('@')[0],
      isVerified: !!idToken // Indicate if identity was token-verified
    });
  } catch (e) {
    Logger.log('getCurrentUserEmail error: ' + e.toString());
    return JSON.stringify({
      success: false,
      email: '',
      agentName: null,
      error: e.toString()
    });
  }
}

/**
 * 🔐 VERIFY GOOGLE ID TOKEN
 * Validates the token against Google's tokeninfo endpoint.
 * Ensures the token was issued to our Client ID and is not expired.
 * 
 * @param {string} token
 * @returns {Object|null} Payload if valid, null otherwise
 */
function verifyGoogleToken_(token) {
  if (!token) return null;
  try {
    // ⚡ Performance: Tokeninfo endpoint call (costs ~200-400ms)
    const response = UrlFetchApp.fetch(`https://oauth2.googleapis.com/tokeninfo?id_token=${token}`, {
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      Logger.log('Token verification failed: ' + response.getContentText());
      return null;
    }
    
    const payload = JSON.parse(response.getContentText());
    
    // Security Checks:
    // 1. Check expiration (handled by Google endpoint response)
    // 2. Check audience (aud) matches our Client ID
    const EXPECTED_CLIENT_ID = '694517401702-4oidkp8p8h9fcpkst0d76o7tn10r1vqq.apps.googleusercontent.com';
    if (payload.aud !== EXPECTED_CLIENT_ID) {
      Logger.log('[Security] Token Audience mismatch!');
      return null;
    }
    
    // 3. Check email_verified
    if (!payload.email_verified) {
      Logger.log('[Security] Email not verified by Google!');
      return null;
    }
    
    return payload;
  } catch (e) {
    Logger.log('verifyGoogleToken_ error: ' + e.toString());
    return null;
  }
}

// Helper for quick whitelist check (cached)
let _cachedAgentEmails = null;
const AGENT_EMAILS_NORMALIZED_ = (function() {
  return Object.values(AGENT_DIRECTORY).map(i => normalizeEmail_(i.email)).filter(Boolean);
})();

/**
 * 🔐 CHECK IF USER HAS PERMISSION (Non-throwing version)
 * @param {string} action - Permission key from PERMISSIONS
 * @returns {boolean} True if user has permission
 */
function hasPermission(action) {
  try {
    requirePermission(action);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * 📝 AUDIT LOGGING
 * Comprehensive audit trail for compliance
 */
function logAuditEvent(action, ticketId, details, severity = 'INFO') {
  try {
    const ss = getSpreadsheet_();
    let auditSheet = ss.getSheetByName(CONFIG.AUDIT_SHEET_NAME);
    
    // Create audit sheet if it doesn't exist
    if (!auditSheet) {
      auditSheet = ss.insertSheet(CONFIG.AUDIT_SHEET_NAME);
      auditSheet.appendRow([
        'Timestamp', 'User Email', 'Action', 'Ticket ID', 
        'Details', 'Severity', 'IP/Session', 'Version'
      ]);
      auditSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#F1F5F9');
      auditSheet.setFrozenRows(1);
    }
    
    // ⚡ FIX5: Auto-rotate when audit log exceeds threshold
    rotateAuditLogIfNeeded_(auditSheet);
    
    const userEmail = getSessionEmail_() || 'system';
    const timestamp = new Date();
    const sessionId = Session.getTemporaryActiveUserKey() || 'N/A';
    
    auditSheet.appendRow([
      timestamp,
      userEmail,
      action,
      ticketId || '-',
      typeof details === 'object' ? JSON.stringify(details) : String(details || ''),
      severity,
      sessionId,
      CONFIG.APP_VERSION
    ]);
    
  } catch (e) {
    Logger.log('Audit log error (non-critical): ' + e.toString());
  }
}

/**
 * ⚡ FIX5: AUDIT LOG ROTATION
 * Archives oldest 2500 rows when the audit sheet exceeds 5000 rows.
 * Runs at most once per request lifecycle (module-level flag).
 * This prevents the audit sheet from growing unbounded and degrading
 * SpreadsheetApp performance over time.
 */
let _auditRotationChecked = false;
function rotateAuditLogIfNeeded_(auditSheet) {
  if (_auditRotationChecked) return; // Only check once per request
  _auditRotationChecked = true;
  
  try {
    const MAX_AUDIT_ROWS = 5000;
    const ROWS_TO_ARCHIVE = 2500;
    const lastRow = auditSheet.getLastRow();
    
    if (lastRow <= MAX_AUDIT_ROWS) return;
    
    Logger.log(`📜 Audit log rotation: ${lastRow} rows exceeds ${MAX_AUDIT_ROWS} limit`);
    
    // Archive oldest rows to a timestamped sheet
    const ss = getSpreadsheet_();
    const archiveName = `Audit Archive ${Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd_HHmmss')}`;
    const archiveSheet = ss.insertSheet(archiveName);
    
    // Copy header + oldest rows to archive
    const headerRow = auditSheet.getRange(1, 1, 1, 8).getValues();
    archiveSheet.getRange(1, 1, 1, 8).setValues(headerRow);
    archiveSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#F1F5F9');
    archiveSheet.setFrozenRows(1);
    
    const archiveData = auditSheet.getRange(2, 1, ROWS_TO_ARCHIVE, 8).getValues();
    archiveSheet.getRange(2, 1, ROWS_TO_ARCHIVE, 8).setValues(archiveData);
    
    // Delete archived rows from the main audit sheet (keep header)
    auditSheet.deleteRows(2, ROWS_TO_ARCHIVE);
    
    Logger.log(`✅ Audit log rotated: ${ROWS_TO_ARCHIVE} rows moved to "${archiveName}", ${lastRow - ROWS_TO_ARCHIVE} rows remaining`);
    
  } catch (e) {
    Logger.log('Audit rotation error (non-critical): ' + e.toString());
  }
}

/**
 * Shared SLA duration formatter for update history.
 * @param {number} ms
 * @returns {{formatted:string,hours:number,days:number,minutes:number,category:string}}
 */
function formatDurationSla_(ms) {
  const safeMs = Math.max(0, Number(ms) || 0);
  const minutes = Math.floor(safeMs / (1000 * 60));
  const totalHours = safeMs / (1000 * 60 * 60);
  const hours = Math.floor(totalHours);
  const days = Math.floor(totalHours / 24);

  let formatted = '';
  if (days >= 1) {
    const remainingHours = hours % 24;
    formatted = `${days}d ${remainingHours}h`;
  } else if (hours >= 1) {
    const remainingMinutes = minutes % 60;
    formatted = remainingMinutes > 0 ? `${hours}h ${remainingMinutes}m` : `${hours}h`;
  } else {
    formatted = `${minutes}m`;
  }

  let category = 'normal';
  if (totalHours < 4) category = 'fast';
  else if (totalHours < 24) category = 'normal';
  else if (totalHours < 72) category = 'slow';
  else category = 'critical';

  return {
    formatted,
    hours: Math.round(totalHours * 10) / 10,
    days: Math.round(days * 10) / 10,
    minutes,
    category
  };
}

const CALL_EVENT_TYPES = Object.freeze([
  'CALL_INITIATED',
  'CALL_DISPOSITION',
  'CALL_CONNECTED',
  'CALL_COMPLETED',
  'CALL_NO_ANSWER',
  'CALL_FAILED',
  'PROVIDER_CDR'
]);

const CALL_OUTCOME_TYPES = Object.freeze([
  '',
  'CONNECTED',
  'NO_ANSWER',
  'BUSY',
  'SWITCHED_OFF',
  'WRONG_NUMBER',
  'CALLBACK_REQUESTED',
  'FAILED',
  'OTHER'
]);

const CALL_LOG_HEADERS = Object.freeze([
  'Timestamp',
  'Event ID',
  'Ticket ID',
  'MID',
  'Business',
  'Customer Phone',
  'Agent Email',
  'Agent Name',
  'Agent Role',
  'Event Type',
  'Outcome',
  'Duration Sec',
  'Channel',
  'Provider',
  'Provider Call ID',
  'Source',
  'Notes',
  'Session Key',
  'Verified'
]);

function normalizeCallPhone_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const cleaned = raw.replace(/[^\d+]/g, '');
  return cleaned.substring(0, 20);
}

function formatPhoneDisplay_(value) {
  const phone = normalizeCallPhone_(value).replace(/^\+/, '');
  if (!phone) return '';
  if (phone.length === 10) return `${phone.slice(0, 5)} ${phone.slice(5)}`;
  if (phone.length === 12 && phone.startsWith('91')) {
    return `+91 ${phone.slice(2, 7)} ${phone.slice(7)}`;
  }
  return value ? String(value) : '';
}

// ⚡ Module-level flag — avoids re-reading the sheet header on every ticket create.
// GAS resets module state between requests, so no staleness risk.
let _phoneColumnEnsured = false;

function ensureTicketPhoneColumn_(sheet) {
  if (_phoneColumnEnsured) return;
  if (!sheet) return;
  const phoneCol = 15;
  const headerValue = String(sheet.getRange(1, phoneCol).getValue() || '').trim();
  if (!headerValue) {
    sheet.getRange(1, phoneCol).setValue('Phone');
  }
  _phoneColumnEnsured = true;
}

function generateCallEventId_(timestamp = new Date()) {
  return `CALL-${Utilities.formatDate(timestamp, 'Asia/Kolkata', 'yyyyMMddHHmmss')}-${Utilities.getUuid().substring(0, 8)}`;
}

function parseDurationSec_(value) {
  const parsed = parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

function getAgentNameByEmail_(email) {
  const agent = getAgentByEmail(email);
  return agent ? agent.name : String(email || 'Unknown');
}

function ensureCallLogSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.CALL_LOG_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(CONFIG.CALL_LOG_SHEET_NAME);
  sheet.getRange(1, 1, 1, CALL_LOG_HEADERS.length).setValues([CALL_LOG_HEADERS]);
  sheet.getRange(1, 1, 1, CALL_LOG_HEADERS.length).setFontWeight('bold').setBackground('#E2E8F0');
  sheet.setFrozenRows(1);
  return sheet;
}

function appendCallLogRow_(entry) {
  const eventType = String(entry.eventType || '').toUpperCase();
  const outcome = String(entry.outcome || '').toUpperCase();
  if (!CALL_EVENT_TYPES.includes(eventType)) {
    throw new Error(`[${ERROR_CODES.VALIDATION_FAILED}] Invalid eventType: ${eventType}`);
  }
  if (!CALL_OUTCOME_TYPES.includes(outcome)) {
    throw new Error(`[${ERROR_CODES.VALIDATION_FAILED}] Invalid outcome: ${outcome}`);
  }

  const timestamp = entry.timestamp || new Date();
  const eventId = entry.eventId || generateCallEventId_(timestamp);
  const verified = String(entry.verified || '').toUpperCase() === 'YES' ? 'YES' : 'NO';

  ensureCallLogSheet_().appendRow([
    timestamp,
    eventId,
    entry.ticketId || '-',
    entry.mid || '-',
    entry.business || '-',
    normalizeCallPhone_(entry.customerPhone),
    entry.agentEmail || 'unknown',
    entry.agentName || 'Unknown',
    entry.role || ROLES.AGENT,
    eventType,
    outcome,
    parseDurationSec_(entry.durationSec),
    entry.channel || 'WEBAPP',
    entry.provider || '',
    entry.providerCallId || '',
    entry.source || 'WEBAPP',
    entry.notes || '',
    entry.sessionKey || '',
    verified
  ]);

  return eventId;
}

function normalizeProviderOutcome_(rawOutcome, rawStatus) {
  const value = String(rawOutcome || rawStatus || '').trim().toLowerCase();
  if (!value) return '';

  if (/connected|answered|completed|success|human/.test(value)) return 'CONNECTED';
  if (/no[\s_-]?answer|unanswered|ring[\s_-]?timeout/.test(value)) return 'NO_ANSWER';
  if (/busy/.test(value)) return 'BUSY';
  if (/switched[\s_-]?off|not[\s_-]?reachable|unreachable|power[\s_-]?off/.test(value)) return 'SWITCHED_OFF';
  if (/wrong[\s_-]?number|invalid[\s_-]?number/.test(value)) return 'WRONG_NUMBER';
  if (/callback/.test(value)) return 'CALLBACK_REQUESTED';
  if (/fail|error|cancel|rejected|blocked|missed/.test(value)) return 'FAILED';
  return 'OTHER';
}

function isDuplicateProviderCdrEvent_(providerCallId, eventType, outcome, durationSec) {
  const callId = String(providerCallId || '').trim();
  if (!callId) return false;

  const sheet = ensureCallLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const scanLimit = 2000;
  const startRow = Math.max(2, lastRow - scanLimit + 1);
  const rowCount = lastRow - startRow + 1;
  const rows = sheet.getRange(startRow, 1, rowCount, CALL_LOG_HEADERS.length).getValues();
  const normalizedCallId = callId.toLowerCase();

  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    const rowCallId = String(row[14] || '').toLowerCase();
    const rowEvent = String(row[9] || '').toUpperCase();
    const rowOutcome = String(row[10] || '').toUpperCase();
    const rowDuration = parseDurationSec_(row[11]);
    const rowSource = String(row[15] || '').toUpperCase();
    if (
      rowCallId === normalizedCallId &&
      rowSource === 'PROVIDER_WEBHOOK' &&
      rowEvent === String(eventType || '').toUpperCase() &&
      rowOutcome === String(outcome || '').toUpperCase() &&
      rowDuration === parseDurationSec_(durationSec)
    ) {
      return true;
    }
  }

  return false;
}

function ingestProviderCdrEvent_(payload = {}) {
  const now = new Date();
  const provider = sanitizeInput(payload.provider || payload.vendor || payload.gateway || 'UNKNOWN', { maxLength: 50 }).toUpperCase();
  const providerCallId = sanitizeInput(
    payload.providerCallId || payload.callId || payload.call_id || payload.sid || payload.uuid,
    { maxLength: 100 }
  );
  if (!providerCallId) {
    return { success: false, error: '[E004] Missing provider call ID in webhook payload' };
  }

  const requestedEvent = sanitizeInput(payload.eventType || payload.event || 'PROVIDER_CDR', { maxLength: 50 }).toUpperCase();
  const eventType = CALL_EVENT_TYPES.includes(requestedEvent) ? requestedEvent : 'PROVIDER_CDR';
  const outcome = normalizeProviderOutcome_(payload.outcome, payload.status || payload.callStatus || payload.disposition);
  const durationSec = parseDurationSec_(
    payload.durationSec || payload.duration || payload.billsec || payload.talkTime || payload.talk_time
  );

  if (isDuplicateProviderCdrEvent_(providerCallId, eventType, outcome, durationSec)) {
    return {
      success: true,
      duplicate: true,
      providerCallId,
      message: 'Duplicate provider CDR event ignored'
    };
  }

  const ticketId = sanitizeInput(payload.ticketId || payload.ticket_id || payload.referenceId, { type: 'id', maxLength: 100 }) || '-';
  const mid = sanitizeInput(payload.mid || payload.merchantId || payload.merchant_id, { maxLength: 50 }) || '-';
  const business = sanitizeInput(payload.business || payload.businessName || payload.customerName, { maxLength: 200 }) || '-';
  const customerPhone = normalizeCallPhone_(payload.customerPhone || payload.phone || payload.to || payload.customer_number);
  const statusText = sanitizeInput(payload.status || payload.callStatus || payload.disposition, { maxLength: 100 });
  const agentEmail = sanitizeInput(
    payload.agentEmail || payload.agent_email || payload.userEmail || 'provider.webhook@system.local',
    { type: 'email', maxLength: 255 }
  ) || 'provider.webhook@system.local';
  const agentName = sanitizeInput(payload.agentName || payload.agent || 'Provider Webhook', { maxLength: 100 }) || 'Provider Webhook';
  const notes = sanitizeInput(payload.notes || payload.note || (statusText ? `provider_status=${statusText}` : ''), { maxLength: 1000 });

  const eventId = appendCallLogRow_({
    timestamp: now,
    ticketId,
    mid,
    business,
    customerPhone,
    agentEmail,
    agentName,
    role: ROLES.SYSTEM,
    eventType,
    outcome,
    durationSec,
    channel: (sanitizeInput(payload.channel || 'PROVIDER_CDR', { maxLength: 50 }) || 'PROVIDER_CDR').toUpperCase(),
    provider,
    providerCallId,
    source: 'PROVIDER_WEBHOOK',
    notes,
    sessionKey: 'WEBHOOK',
    verified: 'YES'
  });

  logAuditEvent('PROVIDER_CDR_RECEIVED', ticketId, {
    eventId,
    provider,
    providerCallId,
    eventType,
    outcome,
    durationSec
  }, 'INFO');

  return {
    success: true,
    duplicate: false,
    eventId,
    provider,
    providerCallId,
    outcome
  };
}

function logCallEvent(callData = {}, csrfToken = '') {
  const lock = LockService.getScriptLock();
  const now = new Date();
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  let lockAcquired = false;
  try {
    requireCSRFToken(csrfToken);
    requirePermission('CALL_LOG_EVENT');
    rateLimitCheck('CALL_LOG_EVENT');
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    lockAcquired = true;

    if (!callData || typeof callData !== 'object') {
      return JSON.stringify({ success: false, error: `[${ERROR_CODES.VALIDATION_FAILED}] Missing call payload` });
    }

    const ticketId = sanitizeInput(callData.ticketId, { type: 'id', maxLength: 100 }) || '-';
    const mid = sanitizeInput(callData.mid, { maxLength: 50 }) || '-';
    const business = sanitizeInput(callData.business, { maxLength: 200 }) || '-';
    const customerPhone = normalizeCallPhone_(callData.customerPhone);
    const requestedEventType = sanitizeInput(callData.eventType, { maxLength: 50 }) || 'CALL_INITIATED';
    const eventType = requestedEventType.toUpperCase();
    const outcomeRaw = sanitizeInput(callData.outcome, { maxLength: 50 }) || '';
    const outcome = outcomeRaw.toUpperCase();
    const durationSec = parseDurationSec_(callData.durationSec);
    const channel = (sanitizeInput(callData.channel, { maxLength: 50 }) || 'WEBAPP').toUpperCase();
    const provider = sanitizeInput(callData.provider, { maxLength: 50 }) || '';
    const providerCallId = sanitizeInput(callData.providerCallId, { maxLength: 100 }) || '';
    const source = (sanitizeInput(callData.source, { maxLength: 50 }) || 'WEBAPP_UI').toUpperCase();
    const notes = sanitizeInput(callData.notes, { maxLength: 1000 }) || '';

    if (!CALL_EVENT_TYPES.includes(eventType)) {
      return JSON.stringify({ success: false, error: `[${ERROR_CODES.VALIDATION_FAILED}] Invalid eventType: ${eventType}` });
    }
    if (!CALL_OUTCOME_TYPES.includes(outcome)) {
      return JSON.stringify({ success: false, error: `[${ERROR_CODES.VALIDATION_FAILED}] Invalid outcome: ${outcome}` });
    }

    const eventId = appendCallLogRow_({
      timestamp: now,
      ticketId,
      mid,
      business,
      customerPhone,
      agentEmail: userEmail,
      agentName: getAgentNameByEmail_(userEmail),
      role: getUserRole(userEmail),
      eventType,
      outcome,
      durationSec,
      channel,
      provider,
      providerCallId,
      source,
      notes,
      sessionKey: Session.getTemporaryActiveUserKey() || '',
      verified: 'NO'
    });

    logAuditEvent('CALL_EVENT_LOGGED', ticketId, {
      eventId,
      eventType,
      outcome,
      channel,
      durationSec,
      source
    }, 'INFO');

    return JSON.stringify({ success: true, eventId, timestamp: now.toISOString() });
  } catch (e) {
    Logger.log('logCallEvent error: ' + e.toString());
    return JSON.stringify({ success: false, error: e.message || e.toString() });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function getCallHistory(config = {}) {
  try {
    requirePermission('CALL_LOG_EVENT');
    const page = Math.max(1, parseInt(config.page, 10) || 1);
    const pageSize = Math.min(Math.max(1, parseInt(config.pageSize, 10) || 50), 200);
    const filters = config.filters || {};
    const currentUserEmail = normalizeEmail_(Session.getActiveUser().getEmail() || '');
    const role = getUserRole(currentUserEmail);
    const canViewAll = isAdminEmail_(currentUserEmail) || role === ROLES.ADMIN || role === ROLES.MANAGER;

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(CONFIG.CALL_LOG_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      return JSON.stringify({
        success: true,
        data: [],
        pagination: { page: 1, pageSize, totalRows: 0, totalPages: 0 },
        message: 'No call records found.'
      });
    }

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, CALL_LOG_HEADERS.length).getValues();
    let records = rows.map(row => {
      const ts = row[0];
      const timestampMs = ts instanceof Date ? ts.getTime() : (new Date(ts).getTime() || 0);
      return {
        timestamp: ts instanceof Date ? Utilities.formatDate(ts, 'Asia/Kolkata', 'dd-MMM-yyyy HH:mm:ss') : String(ts || ''),
        timestampMs: timestampMs,
        eventId: String(row[1] || ''),
        ticketId: String(row[2] || ''),
        mid: String(row[3] || ''),
        business: String(row[4] || ''),
        customerPhone: String(row[5] || ''),
        customerPhoneDisplay: formatPhoneDisplay_(row[5]),
        agentEmail: String(row[6] || ''),
        agentName: String(row[7] || ''),
        role: String(row[8] || ''),
        eventType: String(row[9] || ''),
        outcome: String(row[10] || ''),
        durationSec: parseDurationSec_(row[11]),
        channel: String(row[12] || ''),
        provider: String(row[13] || ''),
        providerCallId: String(row[14] || ''),
        source: String(row[15] || ''),
        notes: String(row[16] || ''),
        sessionKey: String(row[17] || ''),
        verified: String(row[18] || '')
      };
    });

    if (!canViewAll) {
      records = records.filter(r => r.agentEmail.toLowerCase() === currentUserEmail);
    }

    if (filters.ticketId && String(filters.ticketId).trim() !== '') {
      const q = String(filters.ticketId).toLowerCase().trim();
      records = records.filter(r => r.ticketId.toLowerCase().includes(q));
    }
    if (filters.mid && String(filters.mid).trim() !== '') {
      const q = String(filters.mid).toLowerCase().trim();
      records = records.filter(r => r.mid.toLowerCase().includes(q));
    }
    if (filters.agentEmail && String(filters.agentEmail).trim() !== '') {
      const q = String(filters.agentEmail).toLowerCase().trim();
      records = records.filter(r => r.agentEmail.toLowerCase().includes(q));
    }
    if (filters.eventType && String(filters.eventType).trim() !== '' && filters.eventType !== 'all') {
      records = records.filter(r => r.eventType === String(filters.eventType));
    }
    if (filters.outcome && String(filters.outcome).trim() !== '' && filters.outcome !== 'all') {
      records = records.filter(r => r.outcome === String(filters.outcome));
    }
    if (filters.verified && String(filters.verified).trim() !== '' && filters.verified !== 'all') {
      records = records.filter(r => String(r.verified || '').toUpperCase() === String(filters.verified || '').toUpperCase());
    }
    if (filters.startDate) {
      const start = new Date(filters.startDate);
      if (!isNaN(start.getTime())) {
        start.setHours(0, 0, 0, 0);
        records = records.filter(r => r.timestampMs >= start.getTime());
      }
    }
    if (filters.endDate) {
      const end = new Date(filters.endDate);
      if (!isNaN(end.getTime())) {
        end.setHours(23, 59, 59, 999);
        records = records.filter(r => r.timestampMs <= end.getTime());
      }
    }

    records.sort((a, b) => b.timestampMs - a.timestampMs);

    const totalRows = records.length;
    const totalPages = Math.ceil(totalRows / pageSize) || 1;
    const validPage = Math.min(Math.max(1, page), totalPages);
    const startIdx = (validPage - 1) * pageSize;
    const pageData = records.slice(startIdx, startIdx + pageSize);

    return JSON.stringify({
      success: true,
      data: pageData,
      pagination: {
        page: totalRows === 0 ? 1 : validPage,
        pageSize: pageSize,
        totalRows: totalRows,
        totalPages: totalRows === 0 ? 0 : totalPages
      }
    });
  } catch (e) {
    Logger.log('getCallHistory error: ' + e.toString());
    return JSON.stringify({
      success: false,
      error: e.toString(),
      data: [],
      pagination: { page: 1, pageSize: 50, totalRows: 0, totalPages: 0 }
    });
  }
}

function csvSafeCell_(value) {
  const raw = String(value == null ? '' : value);
  const sanitized = /^[=+\-@]/.test(raw) ? `'${raw}` : raw;
  return (sanitized.includes(',') || sanitized.includes('"') || sanitized.includes('\n'))
    ? `"${sanitized.replace(/"/g, '""')}"`
    : sanitized;
}

function csvRow_(values) {
  return values.map(csvSafeCell_).join(',');
}

function exportCallHistoryCSV(config = {}) {
  try {
    requirePermission('EXPORT_HISTORY');
    requireCSRFToken(config.csrfToken || '');
    const filters = config.filters || {};
    const historyResult = JSON.parse(getCallHistory({
      page: 1,
      pageSize: 10000,
      filters: filters
    }));

    if (!historyResult.success) {
      return JSON.stringify({ success: false, error: historyResult.error || 'Unable to export call history' });
    }

    const records = historyResult.data || [];
    const headers = [
      'Timestamp',
      'Event ID',
      'Ticket ID',
      'MID',
      'Business',
      'Customer Phone',
      'Agent Email',
      'Agent Name',
      'Role',
      'Event Type',
      'Outcome',
      'Duration Sec',
      'Channel',
      'Provider',
      'Provider Call ID',
      'Source',
      'Notes',
      'Session Key',
      'Verified'
    ];

    const csvRows = [csvRow_(headers)];
    records.forEach(record => {
      const values = [
        record.timestamp,
        record.eventId,
        record.ticketId,
        record.mid,
        record.business,
        record.customerPhone,
        record.agentEmail,
        record.agentName,
        record.role,
        record.eventType,
        record.outcome,
        record.durationSec,
        record.channel,
        record.provider,
        record.providerCallId,
        record.source,
        record.notes,
        record.sessionKey,
        record.verified
      ];
      csvRows.push(csvRow_(values));
    });

    return JSON.stringify({
      success: true,
      csv: csvRows.join('\n'),
      rowCount: records.length,
      filename: `call_history_${Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd_HHmmss')}.csv`
    });
  } catch (e) {
    Logger.log('exportCallHistoryCSV error: ' + e.toString());
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function parseWebhookPayload_(e) {
  const payload = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(key => {
      payload[key] = e.parameter[key];
    });
  }

  if (e && e.postData && e.postData.contents) {
    const contentType = String(e.postData.type || '').toLowerCase();
    const raw = e.postData.contents;
    if (contentType.includes('application/json')) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') {
          Object.assign(payload, parsed);
        }
      } catch (parseError) {
        throw new Error(`[${ERROR_CODES.VALIDATION_FAILED}] Invalid JSON payload`);
      }
    } else if (raw && raw.trim()) {
      payload.rawBody = raw;
    }
  }

  return payload;
}

function getCallWebhookSecret_() {
  return PropertiesService.getScriptProperties().getProperty('CALL_WEBHOOK_SECRET') || '';
}

function secureEquals_(leftValue, rightValue) {
  const left = String(leftValue || '');
  const right = String(rightValue || '');

  const leftDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, left, Utilities.Charset.UTF_8);
  const rightDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, right, Utilities.Charset.UTF_8);
  if (!leftDigest || !rightDigest || leftDigest.length !== rightDigest.length) {
    return false;
  }

  let diff = left.length ^ right.length;
  for (let i = 0; i < leftDigest.length; i++) {
    diff |= ((leftDigest[i] & 0xff) ^ (rightDigest[i] & 0xff));
  }
  return diff === 0;
}

function validateWebhookSecret_(payload = {}, e) {
  const configuredSecret = getCallWebhookSecret_();
  if (!configuredSecret) {
    return {
      success: false,
      error: '[E002] CALL_WEBHOOK_SECRET is not configured in Script Properties'
    };
  }

  const requestSecret = String(
    payload.secret ||
    payload.webhookSecret ||
    payload.token ||
    (e && e.parameter && e.parameter.secret) ||
    ''
  );

  if (!requestSecret || !secureEquals_(requestSecret, configuredSecret)) {
    return { success: false, error: '[E002] Invalid webhook secret' };
  }

  return { success: true };
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🧹 INPUT SANITIZATION
 * Prevents injection attacks and data corruption
 */
function sanitizeInput(input, options = {}) {
  if (input === null || input === undefined) return options.default || '';
  
  let sanitized = String(input).trim();
  
  // Remove dangerous characters
  sanitized = sanitized.replace(/[<>]/g, '');
  
  // Limit length
  if (options.maxLength && sanitized.length > options.maxLength) {
    sanitized = sanitized.substring(0, options.maxLength);
  }
  
  // Type-specific sanitization
  if (options.type === 'email') {
    sanitized = sanitized.toLowerCase().replace(/[^a-z0-9@._-]/g, '');
  } else if (options.type === 'id') {
    sanitized = sanitized.toUpperCase().replace(/[^A-Z0-9-]/g, '');
  } else if (options.type === 'number') {
    sanitized = sanitized.replace(/[^0-9.-]/g, '');
  }
  
  return sanitized;
}

/* ==================================================
   🔧 PHASE 2.3: ENHANCED VALIDATION LAYER
   ================================================== */

/**
 * 📋 VALIDATION SCHEMAS
 * Define validation rules for different data types
 */
const VALIDATION_SCHEMAS = {
  ticketId: {
    type: 'string',
    required: true,
    minLength: 3,
    maxLength: 50,
    pattern: /^[A-Z0-9-]+$/i,
    sanitize: { type: 'id', maxLength: 50 }
  },
  status: {
    type: 'string',
    required: true,
    allowedValues: Object.values(STATUS_ENUM)
  },
  reason: {
    type: 'string',
    required: false,
    minLength: 3,
    maxLength: 2000,
    sanitize: { maxLength: 2000 }
  },
  email: {
    type: 'string',
    required: true,
    pattern: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    sanitize: { type: 'email', maxLength: 255 }
  },
  mid: {
    type: 'string',
    required: true,
    minLength: 1,
    maxLength: 50,
    sanitize: { type: 'id', maxLength: 50 }
  }
};

/**
 * ✅ VALIDATE INPUT AGAINST SCHEMA
 * @param {*} value - Value to validate
 * @param {string} schemaName - Schema key from VALIDATION_SCHEMAS
 * @returns {Object} { valid: boolean, value: sanitized, errors: string[] }
 */
function validateField(value, schemaName) {
  const schema = VALIDATION_SCHEMAS[schemaName];
  if (!schema) {
    return { valid: false, value: null, errors: [`Unknown schema: ${schemaName}`] };
  }
  
  const errors = [];
  let sanitizedValue = value;
  
  // Sanitize first if schema has sanitize rules
  if (schema.sanitize) {
    sanitizedValue = sanitizeInput(value, schema.sanitize);
  }
  
  // Required check
  if (schema.required && (!sanitizedValue || sanitizedValue.toString().trim() === '')) {
    errors.push(`${schemaName} is required`);
    return { valid: false, value: null, errors };
  }
  
  // Skip further validation if empty and not required
  if (!sanitizedValue || sanitizedValue.toString().trim() === '') {
    return { valid: true, value: '', errors: [] };
  }
  
  const strValue = String(sanitizedValue);
  
  // Min length
  if (schema.minLength && strValue.length < schema.minLength) {
    errors.push(`${schemaName} must be at least ${schema.minLength} characters`);
  }
  
  // Max length
  if (schema.maxLength && strValue.length > schema.maxLength) {
    errors.push(`${schemaName} must be at most ${schema.maxLength} characters`);
  }
  
  // Pattern
  if (schema.pattern && !schema.pattern.test(strValue)) {
    errors.push(`${schemaName} format is invalid`);
  }
  
  // Allowed values
  if (schema.allowedValues && !schema.allowedValues.includes(strValue)) {
    errors.push(`${schemaName} must be one of: ${schema.allowedValues.join(', ')}`);
  }
  
  return {
    valid: errors.length === 0,
    value: sanitizedValue,
    errors: errors
  };
}

/**
 * ✅ VALIDATE MULTIPLE FIELDS
 * @param {Object} data - Object with field values
 * @param {Object} schemas - Map of field names to schema names
 * @returns {Object} { valid: boolean, values: {}, errors: {} }
 */
function validateFields(data, schemas) {
  const result = {
    valid: true,
    values: {},
    errors: {}
  };
  
  for (const [fieldName, schemaName] of Object.entries(schemas)) {
    const validation = validateField(data[fieldName], schemaName);
    result.values[fieldName] = validation.value;
    
    if (!validation.valid) {
      result.valid = false;
      result.errors[fieldName] = validation.errors;
    }
  }
  
  return result;
}

/* ==================================================
   🔧 PHASE 2.4: AUDIT LOG MONITORING
   ================================================== */

/**
 * 📊 GET AUDIT LOG STATISTICS
 * Returns metrics about the audit log for monitoring
 */
function getAuditLogStats() {
  try {
    const ss = getSpreadsheet_();
    const auditSheet = ss.getSheetByName(CONFIG.AUDIT_SHEET_NAME);
    
    if (!auditSheet) {
      return { exists: false, rowCount: 0
      };
    }
    
    const lastRow = auditSheet.getLastRow();
    const rowCount = Math.max(0, lastRow - 1); // Exclude header
    
    // Get last 100 entries for analysis
    const sampleSize = Math.min(100, rowCount);
    const data = sampleSize > 0 
      ? auditSheet.getRange(lastRow - sampleSize + 1, 1, sampleSize, 6).getValues()
      : [];
    
    // Count by severity
    const severityCounts = { INFO: 0, WARNING: 0, ERROR: 0 };
    data.forEach(row => {
      const severity = String(row[5] || 'INFO').toUpperCase();
      if (severityCounts[severity] !== undefined) {
        severityCounts[severity]++;
      }
    });
    
    return {
      exists: true,
      rowCount: rowCount,
      sampleSize: sampleSize,
      severityCounts: severityCounts,
      needsRotation: rowCount > 10000,
      lastEntry: data.length > 0 ? data[data.length - 1][0] : null
    };
  } catch (e) {
    Logger.log('getAuditLogStats error: ' + e.toString());
    return { exists: false, error: e.toString() };
  }
}

/**
 * 🔄 ARCHIVE OLD AUDIT ENTRIES
 * Moves entries older than 90 days to archive sheet
 */
function archiveOldAuditEntries() {
  try {
    const ss = getSpreadsheet_();
    const auditSheet = ss.getSheetByName(CONFIG.AUDIT_SHEET_NAME);
    
    if (!auditSheet || auditSheet.getLastRow() < 2) {
      return { archived: 0, message: 'No entries to archive' };
    }
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 90);
    
    // Get or create archive sheet
    let archiveSheet = ss.getSheetByName('Audit Archive');
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('Audit Archive');
      archiveSheet.appendRow([
        'Timestamp', 'User Email', 'Action', 'Ticket ID', 
        'Details', 'Severity', 'IP/Session', 'Version'
      ]);
      archiveSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#F1F5F9');
      archiveSheet.setFrozenRows(1);
    }
    
    // Process in batches to avoid timeout
    const allData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 8).getValues();
    const toArchive = [];
    const toKeep = [];
    
    allData.forEach(row => {
      const timestamp = new Date(row[0]);
      if (timestamp < cutoffDate) {
        toArchive.push(row);
      } else {
        toKeep.push(row);
      }
    });
    
    if (toArchive.length > 0) {
      // Append to archive
      archiveSheet.getRange(
        archiveSheet.getLastRow() + 1, 
        1, 
        toArchive.length, 
        8
      ).setValues(toArchive);
      
      // Clear main sheet and restore kept entries
      auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 8).clearContent();
      if (toKeep.length > 0) {
        auditSheet.getRange(2, 1, toKeep.length, 8).setValues(toKeep);
      }
    }
    
    return {
      archived: toArchive.length,
      kept: toKeep.length,
      message: `Archived ${toArchive.length} entries older than 90 days`
    };
  } catch (e) {
    Logger.log('archiveOldAuditEntries error: ' + e.toString());
    return { archived: 0, error: e.toString() };
  }
}

/**
 * 🔎 TICKET INDEX MAP
 * O(1) ticket lookups instead of O(n) linear search
 */

function getTicketIndex() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('TICKET_INDEX');
  
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Cache corrupted, rebuild
    }
  }
  
  return buildTicketIndex();
}

function buildTicketIndex() {
  const cache = CacheService.getScriptCache();
  // ✅ FIX: Syntax error — stray newline between getSpreadsheet_() and .getSheetByName() removed
  const sheet = getSpreadsheet_().getSheetByName(SHEET_NAME);
  
  if (!sheet) return {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const index = {};
  
  ids.forEach((row, i) => {
    const id = String(row[0]).trim();
    if (id) {
      index[id] = i + 2; // Convert to row number (1-indexed + header)
    }
  });
  
  // Cache the index
  try {
    cache.put('TICKET_INDEX', JSON.stringify(index), CONFIG.TICKET_INDEX_TTL);
  } catch (e) {
    Logger.log('Index cache error: ' + e.toString());
  }
  
  return index;
}

function invalidateTicketIndex() {
  const cache = CacheService.getScriptCache();
  cache.remove('TICKET_INDEX');
}

/* ==================================================
   ⚡ SMART TICKET CACHING SYSTEM v1.0
   High-performance caching with chunked storage
   ================================================== */

/**
 * 🔧 CACHE CONFIGURATION
 */
const TICKET_CACHE_CONFIG = {
  CACHE_KEY_PREFIX: 'TICKETS_V2_',
  CHUNK_SIZE: 80000,        // ~80KB per chunk (GAS limit is 100KB)
  MAX_CHUNKS: 10,           // Max 10 chunks = ~800KB
  TTL_SECONDS: 300,         // 5 minutes
  METADATA_KEY: 'TICKET_CACHE_META'
};

/**
 * ⚡ GET CACHED TICKETS
 * Returns tickets from cache or fetches fresh data
 * Uses chunked storage for large datasets (>100KB)
 * 
 * @param {boolean} forceRefresh - Skip cache and fetch fresh data
 * @returns {Array} Array of ticket objects
 */
function getCachedTickets(forceRefresh = false) {
  const cache = CacheService.getScriptCache();
  const startTime = Date.now();
  
  try {
    // Check if force refresh requested
    if (forceRefresh) {
      Logger.log('⚡ Cache: Force refresh requested');
      return refreshTicketCache();
    }
    
    // Try to get cached metadata
    const metaRaw = cache.get(TICKET_CACHE_CONFIG.METADATA_KEY);
    
    if (!metaRaw) {
      Logger.log('⚡ Cache: No metadata found, refreshing');
      return refreshTicketCache();
    }
    
    const meta = JSON.parse(metaRaw);
    
    // Check data version - invalidate if stale
    const props = PropertiesService.getScriptProperties();
    const currentVersion = parseInt(props.getProperty('DATA_VERSION') || '0');
    
    if (meta.version !== currentVersion) {
      Logger.log(`⚡ Cache: Version mismatch (cached: ${meta.version}, current: ${currentVersion})`);
      return refreshTicketCache();
    }
    
    // Reconstruct data from chunks
    const chunks = [];
    for (let i = 0; i < meta.chunkCount; i++) {
      const chunkKey = TICKET_CACHE_CONFIG.CACHE_KEY_PREFIX + i;
      const chunk = cache.get(chunkKey);
      
      if (!chunk) {
        Logger.log(`⚡ Cache: Missing chunk ${i}, refreshing`);
        return refreshTicketCache();
      }
      chunks.push(chunk);
    }
    
    const tickets = JSON.parse(chunks.join(''));
    const duration = Date.now() - startTime;
    Logger.log(`⚡ Cache: HIT - ${tickets.length} tickets in ${duration}ms`);
    
    return tickets;
    
  } catch (e) {
    Logger.log('⚡ Cache: Error reading cache - ' + e.toString());
    return refreshTicketCache();
  }
}

/**
 * 🔄 REFRESH TICKET CACHE
 * Fetches fresh data and stores in chunked cache
 * 
 * @returns {Array} Fresh ticket data
 */
function refreshTicketCache() {
  const cache = CacheService.getScriptCache();
  const startTime = Date.now();
  
  try {
    // Get fresh data
    const tickets = getDataObjects();
    const json = JSON.stringify(tickets);
    
    // Get current version
    const props = PropertiesService.getScriptProperties();
    const version = parseInt(props.getProperty('DATA_VERSION') || '0');
    
    // Calculate chunks needed
    const chunkCount = Math.ceil(json.length / TICKET_CACHE_CONFIG.CHUNK_SIZE);
    
    if (chunkCount > TICKET_CACHE_CONFIG.MAX_CHUNKS) {
      Logger.log(`⚡ Cache: Data too large (${chunkCount} chunks), caching skipped`);
      return tickets;
    }
    
    // Clear old chunks first
    invalidateTicketCache();
    
    // Store chunks
    const cacheData = {};
    for (let i = 0; i < chunkCount; i++) {
      const start = i * TICKET_CACHE_CONFIG.CHUNK_SIZE;
      const end = start + TICKET_CACHE_CONFIG.CHUNK_SIZE;
      const chunkKey = TICKET_CACHE_CONFIG.CACHE_KEY_PREFIX + i;
      cacheData[chunkKey] = json.substring(start, end);
    }
    
    // Store metadata
    const meta = {
      version: version,
      chunkCount: chunkCount,
      ticketCount: tickets.length,
      cachedAt: new Date().toISOString(),
      sizeBytes: json.length
    };
    cacheData[TICKET_CACHE_CONFIG.METADATA_KEY] = JSON.stringify(meta);
    
    // Batch put all chunks (more efficient)
    cache.putAll(cacheData, TICKET_CACHE_CONFIG.TTL_SECONDS);
    
    const duration = Date.now() - startTime;
    Logger.log(`⚡ Cache: REFRESHED - ${tickets.length} tickets, ${chunkCount} chunks, ${json.length} bytes in ${duration}ms`);
    
    return tickets;
    
  } catch (e) {
    Logger.log('⚡ Cache: Error refreshing - ' + e.toString());
    // Fallback to direct fetch
    return getDataObjects();
  }
}

/**
 * 🗑️ INVALIDATE TICKET CACHE
 * Clears all cached ticket data
 */
function invalidateTicketCache() {
  const cache = CacheService.getScriptCache();
  
  try {
    // Remove metadata
    cache.remove(TICKET_CACHE_CONFIG.METADATA_KEY);
    
    // Remove all possible chunks
    const keysToRemove = [];
    for (let i = 0; i < TICKET_CACHE_CONFIG.MAX_CHUNKS; i++) {
      keysToRemove.push(TICKET_CACHE_CONFIG.CACHE_KEY_PREFIX + i);
    }
    cache.removeAll(keysToRemove);
    
    Logger.log('⚡ Cache: Invalidated');
  } catch (e) {
    Logger.log('⚡ Cache: Invalidation error - ' + e.toString());
  }
}

/**
 * 📊 GET CACHE STATISTICS
 * Returns current cache status and metrics
 */
function getTicketCacheStats() {
  const cache = CacheService.getScriptCache();
  
  try {
    const metaRaw = cache.get(TICKET_CACHE_CONFIG.METADATA_KEY);
    
    if (!metaRaw) {
      return {
        status: 'EMPTY',
        message: 'No cached data'
      };
    }
    
    const meta = JSON.parse(metaRaw);
    const props = PropertiesService.getScriptProperties();
    const currentVersion = parseInt(props.getProperty('DATA_VERSION') || '0');
    
    return {
      status: meta.version === currentVersion ? 'VALID' : 'STALE',
      ticketCount: meta.ticketCount,
      chunkCount: meta.chunkCount,
      sizeKB: Math.round(meta.sizeBytes / 1024),
      cachedAt: meta.cachedAt,
      cacheVersion: meta.version,
      currentVersion: currentVersion,
      isStale: meta.version !== currentVersion
    };
    
  } catch (e) {
    return {
      status: 'ERROR',
      error: e.toString()
    };
  }
}

// [REMOVED] getSystemHealthLegacy() — superseded by getSystemHealth() (Phase 4)

/**
 * 🔄 RETRY WRAPPER
 * Automatic retry with exponential backoff
 */
function withRetry(fn, maxRetries = 3, baseDelayMs = 1000) {
  let lastError;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return fn();
    } catch (e) {
      lastError = e;
      
      // Don't retry on permission errors
      if (e.message && e.message.includes(ERROR_CODES.INSUFFICIENT_PERMISSIONS)) {
        throw e;
      }
      
      if (attempt < maxRetries) {
        const delay = baseDelayMs * Math.pow(2, attempt - 1);
        Utilities.sleep(delay);
        Logger.log(`Retry attempt ${attempt}/${maxRetries} after ${delay}ms`);
      }
    }
  }
  
  throw lastError;
}

/* ==================================================
   📝 MARKDOWN CONVERTER
   Full Markdown support for reason fields
   ================================================== */

/**
 * Converts Markdown text to HTML
 * Supports: bold, italic, headers, lists, code, links
 */
function convertMarkdownToHtml(text) {
  if (!text) return '';
  
  let html = String(text);
  
  // Escape HTML first to prevent XSS
  html = html.replace(/&/g, '&amp;')
             .replace(/</g, '&lt;')
             .replace(/>/g, '&gt;');
  
  // Headers (h1-h3)
  html = html.replace(/^### (.*$)/gm, '<h4 style="margin:8px 0;font-size:13px;font-weight:700;">$1</h4>');
  html = html.replace(/^## (.*$)/gm, '<h3 style="margin:8px 0;font-size:14px;font-weight:700;">$1</h3>');
  html = html.replace(/^# (.*$)/gm, '<h2 style="margin:8px 0;font-size:15px;font-weight:700;">$1</h2>');
  
  // Bold: **text** or __text__
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/__(.*?)__/g, '<strong>$1</strong>');
  
  // Italic: *text* or _text_
  html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');
  html = html.replace(/_(.*?)_/g, '<em>$1</em>');
  
  // Strikethrough: ~~text~~
  html = html.replace(/~~(.*?)~~/g, '<del>$1</del>');
  
  // Inline code: `code`
  html = html.replace(/`(.*?)`/g, '<code style="background:#F1F5F9;padding:2px 6px;border-radius:4px;font-family:monospace;font-size:12px;">$1</code>');
  
  // Unordered lists: - item
  html = html.replace(/^\- (.*$)/gm, '<li style="margin-left:16px;">$1</li>');
  
  // Ordered lists: 1. item
  html = html.replace(/^\d+\. (.*$)/gm, '<li style="margin-left:16px;">$1</li>');
  
  // Links: [text](url)
  html = html.replace(/\[(.*?)\]\((.*?)\)/g, '<a href="$2" target="_blank" style="color:#4F46E5;">$1</a>');
  
  // Blockquotes: > text
  html = html.replace(/^&gt; (.*$)/gm, '<blockquote style="border-left:3px solid #CBD5E1;padding-left:12px;color:#64748B;margin:8px 0;">$1</blockquote>');
  
  // Line breaks
  html = html.replace(/\n/g, '<br>');
  
  return html;
}

/* ==================================================
   📊 MONTHLY REPORT CSV EXPORT
   (Legacy report generator removed — see generateMonthlyReport in Phase 4)
   ================================================== */

// [REMOVED] generateMonthlyReportLegacy() — superseded by generateMonthlyReport() (Phase 4)

/**
 * Exports report as CSV for Excel
 */
function exportReportAsCSV(config) {
  try {
    requirePermission('EXPORT_REPORT');
    requireCSRFToken((config && config.csrfToken) || '');
    const reportResult = JSON.parse(generateMonthlyReport(config));
    if (!reportResult.success) return reportResult;
    
    const report = reportResult.report;
    
    // Build CSV
    const csvRows = [];
    csvRows.push(csvRow_([report.title]));
    csvRows.push(csvRow_([`Generated: ${report.generatedAt}`]));
    csvRows.push(csvRow_([`Period: ${report.period.startDate} to ${report.period.endDate}`]));
    csvRows.push('');

    csvRows.push(csvRow_(['SUMMARY']));
    csvRows.push(csvRow_(['Total Tickets', report.summary.totalTickets]));
    csvRows.push(csvRow_(['Completed', report.summary.completed]));
    csvRows.push(csvRow_(['Pending', report.summary.pending]));
    csvRows.push(csvRow_(['Closed', report.summary.closed]));
    csvRows.push(csvRow_([`Can't Do`, report.summary.cantDo]));
    csvRows.push(csvRow_(['Completion Rate', `${report.summary.completionRate}%`]));
    csvRows.push('');

    csvRows.push(csvRow_(['AGENT PERFORMANCE']));
    csvRows.push(csvRow_(['Agent', 'Total', 'Completed', 'Pending', 'Closed', `Can't Do`, 'Rate']));
    report.agentRankings.forEach(a => {
      csvRows.push(csvRow_([a.name, a.total, a.completed, a.pending, a.closed, a.cantDo, `${a.completionRate}%`]));
    });
    csvRows.push('');

    csvRows.push(csvRow_(['TICKET DETAILS']));
    csvRows.push(csvRow_(['ID', 'Date', 'Agent', 'Business', 'MID', 'Concern', 'Support Type', 'Status']));
    report.tickets.forEach(t => {
      csvRows.push(csvRow_([t.id, t.date, t.agent, t.business, t.mid, t.concern, t.supportType, t.status]));
    });
    
    return JSON.stringify({ success: true, csv: csvRows.join('\n'), filename: `Report_${report.period.monthName}_${report.period.year}.csv` });
    
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

// [REMOVED] getUpdateHistoryLegacy() — superseded by getUpdateHistory() (Phase 4)





/* ==================================================
   🔹 WEB APP SERVING
   ================================================== */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');

  // Use only authenticated session identity (fail-closed).
  // Fallback: getEffectiveUser() works for "Execute as: Me" deployments
  let sessionEmail = '';
  try { sessionEmail = Session.getActiveUser().getEmail() || ''; } catch (_) {}
  if (!sessionEmail) {
    try { sessionEmail = Session.getEffectiveUser().getEmail() || ''; } catch (_) {}
  }

  const agent = getAgentByEmail(sessionEmail);
  const role  = getUserRole(sessionEmail);

  template.injectedUserEmail   = sessionEmail;
  template.injectedUserName    = agent ? agent.name : '';
  template.injectedUserRole    = role;
  template.injectedUserIsAdmin = isAdminEmail_(sessionEmail || '');

  const props = PropertiesService.getScriptProperties();
  const trustedOriginsRaw = props.getProperty('TRUSTED_PARENT_ORIGINS') || '';
  const trustedOrigins = trustedOriginsRaw
    .split(',')
    .map(x => x.trim())
    .filter(Boolean);
  template.injectedTrustedParentOrigins = JSON.stringify(trustedOrigins);

  // Inject full agent list at render-time to avoid client boot race conditions.
  try {
    template.injectedAgentsJson = getAgentList(); // already a JSON string
  } catch (_) {
    template.injectedAgentsJson = JSON.stringify({ success: false, agents: [] });
  }

  const allowIframeEmbed = String(props.getProperty('ALLOW_IFRAME_EMBED') || '').toLowerCase() === 'true';
  return template.evaluate()
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const payload = parseWebhookPayload_(e);
    const action = String(
      payload.action ||
      payload.route ||
      (e && e.parameter && e.parameter.action) ||
      'provider_cdr'
    ).toLowerCase();

    if (action !== 'provider_cdr') {
      return jsonResponse_({
        success: false,
        error: `[${ERROR_CODES.VALIDATION_FAILED}] Unsupported webhook action: ${action}`
      });
    }

    const secretValidation = validateWebhookSecret_(payload, e);
    if (!secretValidation.success) {
      return jsonResponse_(secretValidation);
    }

    delete payload.secret;
    delete payload.webhookSecret;
    delete payload.token;

    const result = ingestProviderCdrEvent_(payload);
    return jsonResponse_(result);
  } catch (ePost) {
    Logger.log('doPost error: ' + ePost.toString());
    return jsonResponse_({
      success: false,
      error: ePost.message || ePost.toString()
    });
  }
}


/* ==================================================
   🔹 CORE DATA ENGINE (UPDATED WITH 7-DAY VALIDATION)
   ================================================== */
function getDataObjects() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // ⚡ OPTIMIZATION 1: Single batch read
  const readCols = Math.min(Math.max(sheet.getLastColumn(), 14), 15); // 15th column reserved for Phone
  const data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues(); // Skip header
  const now = new Date();
  const nowTime = now.getTime(); // ⚡ Pre-compute once
  
  // ⚡ OPTIMIZATION 2: Pre-allocate array (faster than dynamic growth)
  const tickets = [];
  tickets.length = data.length;
  
  let validIndex = 0;

  // ⚡ OPTIMIZATION 3: For-loop (30% faster than map/filter chain)
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    // ⚡ Quick skip empty rows
    const rawDate = row[1];
    if (!rawDate || String(rawDate).trim() === '') continue;

    // ⚡ OPTIMIZATION 4: Simplified date handling
    let sortDate = 0;
    let displayDate = '-';
    let hourIST = 0;

    if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
      // ⚡ FAST PATH: 95% of cases
      const d = rawDate;
      sortDate = d.getTime();
      displayDate = formatDateFast(d); // ⚡ Use helper below
      hourIST = d.getHours();
    } else {
      // Fallback for strings
      const d = new Date(rawDate);
      if (!isNaN(d.getTime())) {
        sortDate = d.getTime();
        displayDate = formatDateFast(d);
        hourIST = d.getHours();
      } else {
        sortDate = nowTime;
      }
    }

    // ⚡ OPTIMIZATION 5: Pre-compute age
    const ageDays = Math.floor((nowTime - sortDate) / 86400000); // Use constant

    // ⚡ OPTIMIZATION 6: Single status normalization
    const status = normalizeStatusWithDefault(row[12]);
    
    // ⚡ OPTIMIZATION 7: Defer reason quality (only if needed for display)
    const reason = row[13] ? String(row[13]).trim() : '';
    const reasonLen = reason.length;
    
    // ⚡ OPTIMIZATION 8: Inline age category (no if-else chain)
    const ageCategory = ageDays >= 15 ? 'critical' : 
                        ageDays >= 8 ? 'old' : 
                        ageDays >= 4 ? 'aging' : 'fresh';

    // ⚡ CRITICAL FIX: Remove validation call entirely
    // Frontend will validate on-demand for visible tickets only
    const phoneRaw = row[14] || '';
    const phone = normalizeCallPhone_(phoneRaw);

    const ticketObject = {
      id: String(row[0] || `TKT-${i + 1}`),
      rowIndex: i + 2,
      date: displayDate,
      sortDate: sortDate,
      hourIST: hourIST,
      ageDays: ageDays,
      ageCategory: ageCategory,
      email: row[3] || '',
      agent: getAgentNameByEmail_(row[2]),
      requestedBy: row[4] || '-',
      mid: row[5] ? String(row[5]).trim() : '-',
      business: row[6] || '-',
      pos: row[7] ? String(row[7]).trim() : '-',
      supportType: row[8] || 'Customer Support',
      concern: row[9] || 'Unspecified',
      config: row[10] || '',
      remark: row[11] || '',
      phone: phone,
      phoneDisplay: formatPhoneDisplay_(phoneRaw),
      status: status,
      reason: reason,
      reasonQuality: reasonLen >= 30 ? 'detailed' : 
                     reasonLen >= 10 ? 'brief' : 
                     reasonLen > 0 ? 'minimal' : 'none',
      // ⚡ DEFERRED: Compute these in frontend on-demand
      invalidClosed: false, // Will be computed by frontend when needed
      validationReason: '',
      validationWarnings: []
    };

    tickets[validIndex++] = ticketObject;
  }

  // ⚡ Trim unused slots
  tickets.length = validIndex;
  
  return tickets;
}

// ⚡ HELPER: Fast date formatting (2x faster than template literals)
function formatDateFast(d) {
  const day = d.getDate();
  const month = d.getMonth() + 1;
  return (day < 10 ? '0' + day : day) + '-' + 
         (month < 10 ? '0' + month : month) + '-' + 
         d.getFullYear();
}


function getTicketData() {
  try {
    requirePermission('VIEW_ANALYTICS');
    // ⚡ Use cached tickets for 10x faster response
    const tickets = getCachedTickets();
    const props = PropertiesService.getScriptProperties();
    const version = parseInt(props.getProperty('DATA_VERSION') || '0');
    
    // Include cache stats for monitoring
    const cacheStats = getTicketCacheStats();

    return JSON.stringify({ 
      success: true, 
      tickets: tickets.sort((a, b) => b.sortDate - a.sortDate),
      directory: AGENT_DIRECTORY,
      version: version,
      cacheStatus: cacheStats.status
    });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}


/* ==================================================
   🆕 UPDATE FUNCTIONS (WITH VALIDATION)
   ================================================== */

function updateTicketStatus(ticketId, newStatus, csrfToken = '') {
  const lock = LockService.getScriptLock();
  const userEmail = Session.getActiveUser().getEmail();
  let lockAcquired = false; // Only release if we actually acquired
  
  try {
    requireCSRFToken(csrfToken);
    requirePermission('UPDATE_TICKET');
    // 🚦 Phase 1: Rate limiting check
    rateLimitCheck('UPDATE_STATUS');
    
    // Validate inputs
    if (!ticketId || !newStatus) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.VALIDATION_FAILED}] Missing required parameters` 
      });
    }

    const normalizedStatus = parseStatusOrNull(newStatus);
    if (!normalizedStatus) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.INVALID_STATUS}] Invalid status value` 
      });
    }

    // 🔒 Phase 1: Acquire lock to prevent race conditions
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    lockAcquired = true;

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.SHEET_ERROR}] Sheet not found` 
      });
    }

    // 🚀 Phase 1: O(1) lookup using ticket index
    const ticketIndex = getTicketIndex();
    let targetRow = ticketIndex[ticketId];
    
    // Fallback to linear search if index miss
    if (!targetRow) {
      const lastRow = sheet.getLastRow();
      const data = sheet.getRange(1, 1, lastRow, 1).getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === ticketId) {
          targetRow = i + 1;
          break;
        }
      }
    }

    if (!targetRow) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.NOT_FOUND}] Ticket not found: ${ticketId}` 
      });
    }

    // Get previous status for audit
    const previousStatus = String(sheet.getRange(targetRow, 13).getValue()).trim();

    // Update status
    sheet.getRange(targetRow, 13).setValue(normalizedStatus);
    SpreadsheetApp.flush();
    
    incrementDataVersion();
    invalidateTicketIndex();
    
    // 📝 Audit log
    logAuditEvent('STATUS_UPDATED', ticketId, {
      previousStatus: previousStatus,
      newStatus: normalizedStatus,
      updatedBy: userEmail
    });

    return JSON.stringify({ 
      success: true, 
      message: 'Status updated successfully',
      ticketId: ticketId,
      newStatus: normalizedStatus
    });
  } catch (e) {
    Logger.log('updateTicketStatus error: ' + e.toString());
    
    logAuditEvent('UPDATE_STATUS_ERROR', ticketId, {
      error: e.toString(),
      attemptedBy: userEmail
    }, 'ERROR');
    
    return JSON.stringify({ success: false, error: e.toString() });
  } finally {
    if (lockAcquired) lock.releaseLock(); // ✅ Safe: only release if acquired
  }
}




function appendTicketReason(ticketId, newReason, csrfToken = '') {
  const lock = LockService.getScriptLock();
  const userEmail = Session.getActiveUser().getEmail();
  let lockAcquired = false; // Only release if we actually acquired
  
  try {
    requireCSRFToken(csrfToken);
    requirePermission('UPDATE_TICKET');
    // 🚦 Phase 1: Rate limiting
    rateLimitCheck('APPEND_REASON');
    
    if (!ticketId || !newReason || newReason.trim().length < 3) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.VALIDATION_FAILED}] Invalid reason (minimum 3 characters)` 
      });
    }

    // 🔒 Phase 1: Acquire lock
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    lockAcquired = true;

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.SHEET_ERROR}] Sheet not found` 
      });
    }
    // 🚀 Phase 1: O(1) lookup using ticket index
    const ticketIndex = getTicketIndex();
    let targetRow = ticketIndex[ticketId];
    let existingReason = '';
    
    // Fallback to linear search if index miss
    if (!targetRow) {
      const lastRow = sheet.getLastRow();
      const data = sheet.getRange(1, 1, lastRow, 14).getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === ticketId) {
          targetRow = i + 1;
          existingReason = String(data[i][13] || '').trim();
          break;
        }
      }
    } else {
      // Get existing reason if found via index
      existingReason = String(sheet.getRange(targetRow, 14).getValue() || '').trim();
    }

    if (!targetRow) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.NOT_FOUND}] Ticket not found: ${ticketId}` 
      });
    }

    const timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM HH:mm");
    const sanitizedReason = sanitizeInput(newReason, { maxLength: 1000 });
    const appendedReason = existingReason 
      ? `${existingReason}\n[${timestamp}] ${sanitizedReason}` 
      : `[${timestamp}] ${sanitizedReason}`;

    sheet.getRange(targetRow, 14).setValue(appendedReason);
    SpreadsheetApp.flush();
    incrementDataVersion();

    // 📝 Audit log
    logAuditEvent('REASON_APPENDED', ticketId, {
      reasonLength: sanitizedReason.length,
      updatedBy: userEmail
    });

    return JSON.stringify({ 
      success: true, 
      message: 'Reason added successfully',
      ticketId: ticketId
    });
  } catch (e) {
    Logger.log('appendTicketReason error: ' + e.toString());
    return JSON.stringify({ success: false, error: e.toString() });
  } finally {
    if (lockAcquired) lock.releaseLock(); // ✅ Safe: only release if acquired
  }
}



function updateTicketFull(ticketId, newStatus, newReason, csrfToken) {
  const lock = LockService.getScriptLock();
  const userEmail = Session.getActiveUser().getEmail();
  let lockAcquired = false; // Only release if we actually acquired
  
  try {
    // 🛡️ Phase 1: CSRF validation (Security Fix)
    requireCSRFToken(csrfToken);
    
    // 🚦 Rate limiting check
    rateLimitCheck('UPDATE_TICKET');
    
    // 🔒 Permission check
    requirePermission('UPDATE_TICKET');
    
    // ✅ FIX: Use CONFIG.LOCK_TIMEOUT_MS instead of hardcoded 5000
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    lockAcquired = true;
    
    Logger.log('=== ENTERPRISE UPDATE TICKET START ===');
    
    // 1. Input Validation & Sanitization
    const sanitizedTicketId = sanitizeInput(ticketId, { type: 'id', maxLength: 50 });
    const sanitizedReason = sanitizeInput(newReason, { maxLength: 2000 });
    
    if (!sanitizedTicketId) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.VALIDATION_FAILED}] Missing Ticket ID` 
      });
    }
    
    if (!newStatus) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.VALIDATION_FAILED}] Missing Status` 
      });
    }

    const normalizedStatus = parseStatusOrNull(newStatus);
    if (!normalizedStatus) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.INVALID_STATUS}] Invalid status: "${newStatus}". Allowed: ${VALID_STATUSES.join(', ')}` 
      });
    }

    // 🔐 Closure authority enforcement (supports multiple admins)
    if (normalizedStatus === STATUS_ENUM.CLOSED) {
      if (!isAdminEmail_(userEmail)) {
        logAuditEvent('CLOSE_ATTEMPT_DENIED', sanitizedTicketId, { 
          attemptedBy: userEmail,
          reason: 'Not in admin list'
        }, 'WARNING');
        
        return JSON.stringify({
          success: false,
          error: `[${ERROR_CODES.INSUFFICIENT_PERMISSIONS}] Only admin can close tickets`
        });
      }

      if (!sanitizedReason || sanitizedReason.trim().length < 5) {
        return JSON.stringify({
          success: false,
          error: `[${ERROR_CODES.VALIDATION_FAILED}] Closure requires follow-up reason (min 5 chars)`
        });
      }
    }
    
    // 2. Open Sheet
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.SHEET_ERROR}] Sheet not found` 
      });
    }

    // 3. 🚀 O(1) LOOKUP using Ticket Index (instead of O(n) linear search)
    const ticketIndex = getTicketIndex();
    let targetRow = ticketIndex[sanitizedTicketId];
    
    // Fallback to linear search if index miss (rare case - new ticket or stale cache)
    if (!targetRow) {
      Logger.log('Index miss for ' + sanitizedTicketId + ', falling back to linear search');
      const lastRow = sheet.getLastRow();
      const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      
      for (let i = 0; i < idColumn.length; i++) {
        if (String(idColumn[i][0]).trim() === sanitizedTicketId) {
          targetRow = i + 2;
          break;
        }
      }
    }
    
    if (!targetRow) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.NOT_FOUND}] Ticket not found: ${sanitizedTicketId}` 
      });
    }

    // 4. Get previous status for audit
    const previousStatus = String(sheet.getRange(targetRow, 13).getValue()).trim();

    // 5. Update Status (Column 13 / M)
    sheet.getRange(targetRow, 13).setValue(normalizedStatus);

    // 6. Update Reason with Timestamp (Column 14 / N)
    let existingReason = String(sheet.getRange(targetRow, 14).getValue()).trim();
    
    if (sanitizedReason && sanitizedReason.trim() !== "") {
      const timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM HH:mm");
      let entry = `[${timestamp}] ${sanitizedReason.trim()}`;
      
      let finalReason = existingReason 
        ? existingReason + "\n" + entry 
        : entry;

      sheet.getRange(targetRow, 14).setValue(finalReason);
    }
    
    // 7. Finish & Audit
    SpreadsheetApp.flush();
    _incrementDataVersionNoLock(); // Use no-lock version since we already hold lock
    invalidateTicketIndex(); // Invalidate index on any change
    
    // 📝 Log successful update
    logAuditEvent('TICKET_UPDATED', sanitizedTicketId, {
      previousStatus: previousStatus,
      newStatus: normalizedStatus,
      reasonAdded: sanitizedReason ? true : false,
      updatedBy: userEmail
    });
    
    return JSON.stringify({ 
      success: true, 
      message: 'Updated successfully',
      ticketId: sanitizedTicketId,
      newStatus: normalizedStatus
    });

  } catch (e) {
    Logger.log('Error: ' + e.toString());
    
    // Log error to audit
    logAuditEvent('UPDATE_ERROR', ticketId, {
      error: e.toString(),
      attemptedBy: userEmail
    }, 'ERROR');
    
    return JSON.stringify({ 
      success: false, 
      error: e.message || e.toString() 
    });
  } finally {
    if (lockAcquired) lock.releaseLock(); // ✅ Safe: only release if acquired
  }
}


/* ==================================================
   🎫 CREATE NEW TICKET (ENTERPRISE GRADE)
   ================================================== */

function createNewTicket(ticketData, csrfToken) {
  const lock = LockService.getScriptLock();
  const userEmail = Session.getActiveUser().getEmail();
  let lockAcquired = false;

  try {
    // 🛡️ Phase 1: CSRF validation (Security Fix)
    requireCSRFToken(csrfToken);
    
    // 🚦 Rate limiting check
    rateLimitCheck('CREATE_TICKET');
    
    // 🔒 Permission check
    requirePermission('CREATE_TICKET');
    
    // ✅ FIX: Use CONFIG.LOCK_TIMEOUT_MS instead of hardcoded 5000
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    lockAcquired = true;
    
    Logger.log('=== ENTERPRISE CREATE TICKET START ===');

    // 1. Input Validation & Sanitization
    if (!ticketData) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.VALIDATION_FAILED}] Missing ticket data` 
      });
    }
    
    const sanitized = {
      email: sanitizeInput(ticketData.email, { type: 'email', maxLength: 100 }),
      agent: sanitizeInput(ticketData.agent, { maxLength: 50 }) || 'Unassigned',
      requestedBy: sanitizeInput(ticketData.requestedBy, { maxLength: 100 }) || '-',
      mid: sanitizeInput(ticketData.mid, { maxLength: 20 }) || '-',
      business: sanitizeInput(ticketData.business, { maxLength: 200 }) || '-',
      pos: sanitizeInput(ticketData.pos, { maxLength: 50 }) || '-',
      supportType: sanitizeInput(ticketData.supportType, { maxLength: 50 }) || 'Customer Support',
      concern: sanitizeInput(ticketData.concern, { maxLength: 100 }) || 'Unspecified',
      config: sanitizeInput(ticketData.config, { maxLength: 100 }) || '',
      remark: sanitizeInput(ticketData.remark, { maxLength: 1000 }) || '',
      reason: sanitizeInput(ticketData.reason, { maxLength: 500 }) || '',
      phone: normalizeCallPhone_(ticketData.phone)
    };
    
    // Validate required fields
    if (!sanitized.agent || sanitized.agent === 'Unassigned') {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.VALIDATION_FAILED}] Agent is required` 
      });
    }
    
    if (!sanitized.concern || sanitized.concern === 'Unspecified') {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.VALIDATION_FAILED}] Concern is required` 
      });
    }

    // 2. Open Sheet
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return JSON.stringify({ 
        success: false, 
        error: `[${ERROR_CODES.SHEET_ERROR}] Sheet not found` 
      });
    }
    ensureTicketPhoneColumn_(sheet);

    // 3. Generate Ticket ID (Format: BF-TKT-YYYY-MM-XXXX)
    const nextRowIndex = sheet.getLastRow();
    const now = new Date();
    const year = Utilities.formatDate(now, "Asia/Kolkata", "yyyy");
    const month = Utilities.formatDate(now, "Asia/Kolkata", "MM");
    const sequence = String(nextRowIndex).padStart(4, '0');
    const generatedTicketId = `BF-TKT-${year}-${month}-${sequence}`;

    const createdStatus = parseStatusOrNull(ticketData.status) || STATUS_ENUM.NOT_COMPLETED;

    // 4. Prepare Row Data
    const newRow = [
      generatedTicketId,
      now,
      sanitized.agent,
      sanitized.email,
      sanitized.requestedBy,
      sanitized.mid,
      sanitized.business,
      sanitized.pos,
      sanitized.supportType,
      sanitized.concern,
      sanitized.config,
      sanitized.remark,
      createdStatus,
      sanitized.reason,
      sanitized.phone
    ];

    // 5. Save to Sheet
    sheet.appendRow(newRow);
    
    // 6. Update version and invalidate index (use no-lock since we hold lock)
    _incrementDataVersionNoLock();
    invalidateTicketIndex();

    // 📝 Audit log AFTER core write — still inside try so generatedTicketId is in scope,
    //    but logAuditEvent failure won't block the success response.
    try {
      logAuditEvent('TICKET_CREATED', generatedTicketId, {
        agent: sanitized.agent,
        concern: sanitized.concern,
        business: sanitized.business,
        createdBy: userEmail
      });
    } catch (_auditErr) {
      Logger.log('Audit log failed (non-critical): ' + _auditErr);
    }

    return JSON.stringify({ 
      success: true, 
      ticketId: generatedTicketId, 
      message: 'Ticket created successfully' 
    });

  } catch (e) {
    Logger.log('Error: ' + e.toString());
    
    try {
      logAuditEvent('CREATE_ERROR', null, {
        error: e.toString(),
        attemptedBy: userEmail
      }, 'ERROR');
    } catch (_auditErr) {
      Logger.log('Audit log failed in error handler: ' + _auditErr);
    }
    
    return JSON.stringify({ 
      success: false, 
      error: e.message || e.toString() 
    });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}


// [REMOVED] getAgentsList() — superseded by getAgentList() which returns full agent objects


/* ==================================================
   🔄 REAL-TIME SYNC LOGIC (Phase 1 Lock Fix)
   ================================================== */
function checkVersion() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('DATA_VERSION');
  const version = Number.isInteger(Number(raw)) ? Number(raw) : 0;

  return JSON.stringify({
    success: true,
    version: version
  });
}

/**
 * 🔒 INTERNAL: Increment version WITHOUT acquiring lock
 * Use this when caller already holds a lock to prevent deadlock
 * @private
 */
function _incrementDataVersionNoLock() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('DATA_VERSION');
  const current = Number.isInteger(Number(raw)) ? Number(raw) : 0;
  props.setProperty('DATA_VERSION', String(current + 1));
  
  // ⚡ Invalidate the smart ticket cache
  invalidateTicketCache();
}

/**
 * 🔄 PUBLIC: Increment version with lock acquisition
 * Use this for standalone calls (not inside another locked function)
 */
function incrementDataVersion() {
  const lock = LockService.getScriptLock();
  try {
    // ✅ FIX: Use CONFIG.LOCK_TIMEOUT_MS instead of hardcoded 5000
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    _incrementDataVersionNoLock();
  } catch (e) {
    Logger.log('Lock timeout in incrementVersion: ' + e);
  } finally {
    lock.releaseLock();
  }
}



function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    const row = e.range.getRow();
    if (row < 2) return; // Ignore header edits

    // ✅ FIX: Was deleting 'TICKET_CACHE' (wrong/legacy key).
    // Now calls invalidateTicketCache() which clears the chunked cache system.
    invalidateTicketCache();

  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}


/* ==================================================
   📧 NOTIFICATION BOTS
   ================================================== */
function botDailyPending() {
  const allTickets = getDataObjects();
  const now = new Date();
  const hour = now.getHours();

  let timeOfDay;
  if (hour < 12) timeOfDay = '🌅 Morning';
  else if (hour < 17) timeOfDay = '☀️ Afternoon';
  else timeOfDay = '🌆 Evening';

  const pending = allTickets.filter(t => {
    const s = t.status.toLowerCase();
    return s === 'not completed' || s === 'pending' || s === 'in progress';
  });

  if (pending.length === 0) return;

  const grouped = pending.reduce((acc, t) => {
    (acc[t.agent] = acc[t.agent] || []).push(t);
    return acc;
  }, {});

  for (const [agent, tickets] of Object.entries(grouped)) {
    if (!AGENT_DIRECTORY[agent]?.email) continue;

    try {
      tickets.forEach(t => {
        const hours = Math.floor((now - t.sortDate) / 36e5);
        t.age = hours > 24 ? `${Math.floor(hours/24)}d` : `${hours}h`;
      });

      MailApp.sendEmail({
        to: AGENT_DIRECTORY[agent].email,
        subject: `${timeOfDay} 📋 Pending Tickets Report (${tickets.length})`,
        htmlBody: generateDailyPendingHTML(agent, tickets, now, timeOfDay)
      });
    } catch (e) { 
      Logger.log(`Failed to email ${agent}: ${e}`); 
    }
  }
}


/**
 * 🛡️ HTML ESCAPE HELPER
 * Prevents XSS / HTML injection when interpolating user data into email bodies.
 * MUST be used on every field sourced from user input or the spreadsheet.
 * @param {*} value - Any value to escape
 * @returns {string} HTML-safe string
 */
function escapeHtml_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function generateDailyPendingHTML(agent, tickets, date, time) {
  // ✅ All user-supplied fields HTML-escaped to prevent email injection / XSS
  const rows = tickets.map(t =>
    `<tr>
      <td style="border-bottom:1px solid #ddd; padding:8px;">${escapeHtml_(t.id)}</td>
      <td style="border-bottom:1px solid #ddd;">${escapeHtml_(t.business)}</td>
      <td style="border-bottom:1px solid #ddd; color:red; font-weight:bold;">${escapeHtml_(t.age)}</td>
    </tr>`
  ).join('');

  return `
    <div style="font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5;">
      <div style="background: white; padding: 30px; border-radius: 10px; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #1a1f3a; margin-bottom: 10px;">Hi ${agent},</h2>
        <h3 style="color: #667eea;">${time} Reminder 🔔</h3>
        <p style="font-size: 16px; color: #555;">You have <strong>${tickets.length} pending tickets</strong> that need attention:</p>

        <table style="width:100%; text-align:left; border-collapse:collapse; margin: 20px 0;">
          <thead>
            <tr style="background: #f8fafc;">
              <th style="padding: 12px 8px; border-bottom: 2px solid #ddd;">Ticket ID</th>
              <th style="padding: 12px 8px; border-bottom: 2px solid #ddd;">Business</th>
              <th style="padding: 12px 8px; border-bottom: 2px solid #ddd;">Age</th>
            </tr>
          </thead>
          <tbody>
            ${rows}
          </tbody>
        </table>

        <p style="color: #888; font-size: 12px; margin-top: 30px;">
          This is an automated reminder from BillFree IT Support System.
        </p>
      </div>
    </div>
  `;
}




/* ==================================================
   📊 ADVANCED ANALYTICS FUNCTIONS
   ================================================== */

function getTopMIDsSameConcern() {
  try {
    const tickets = getDataObjects();
    const excludeMIDs = ['301', '201', '202', '302'];
    
    // ✅ Filter out completed/closed tickets FIRST
    const activeTickets = tickets.filter(t => {
      return t.status !== STATUS_ENUM.CLOSED && t.status !== STATUS_ENUM.COMPLETED;
    });

    const midConcernMap = {};

    activeTickets.forEach(t => {
      if (!t.mid || t.mid === '-' || excludeMIDs.includes(t.mid)) return;

      const key = `${t.mid}|||${t.concern}`;
      if (!midConcernMap[key]) {
        midConcernMap[key] = {
          mid: t.mid,
          concern: t.concern,
          business: t.business,
          count: 0,
          tickets: []
        };
      }
      midConcernMap[key].count++;
      midConcernMap[key].tickets.push(t.id);
    });

    const results = Object.values(midConcernMap)
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    return JSON.stringify({ success: true, data: results });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}




function getTopMIDsDifferentConcerns() {
  try {
    const tickets = getCachedTickets();
    const excludeMIDs = ['301', '201', '202', '302'];
    
    // ✅ Filter out completed/closed tickets FIRST
    const activeTickets = tickets.filter(t => {
      return t.status !== STATUS_ENUM.CLOSED && t.status !== STATUS_ENUM.COMPLETED;
    });

    const midMap = {};

    activeTickets.forEach(t => {
      if (!t.mid || t.mid === '-' || excludeMIDs.includes(t.mid)) return;

      if (!midMap[t.mid]) {
        midMap[t.mid] = {
          mid: t.mid,
          business: t.business,
          concerns: new Set(),
          totalTickets: 0
        };
      }
      midMap[t.mid].concerns.add(t.concern);
      midMap[t.mid].totalTickets++;
    });

    const results = Object.values(midMap)
      .map(m => ({
        mid: m.mid,
        business: m.business,
        concernCount: m.concerns.size,
        concerns: Array.from(m.concerns),
        totalTickets: m.totalTickets
      }))
      .filter(m => m.concernCount > 1)
      .sort((a, b) => b.concernCount - a.concernCount)
      .slice(0, 10);

    return JSON.stringify({ success: true, data: results });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}




function getTopPOS() {
  try {
    const tickets = getCachedTickets();
    const excludePOS = ['bf', 'billfree', '-', '', 'na', 'n/a', 'unknown'];

    // ⚡ Normalization Dictionary
    const normalizePOS = (raw) => {
        const v = raw.toLowerCase().replace(/[^a-z0-9]/g, '');
        
        if (v.includes('tally')) return 'Tally';
        if (v.includes('busy')) return 'Busy';
        if (v.includes('custom') || v.includes('costum')) return 'Custom';
        if (v.includes('mmi')) return 'MMI';
        if (v.includes('petpooja') || v.includes('pet')) return 'PetPooja';
        if (v.includes('margh') || v.includes('marg')) return 'Marg';
        if (v.includes('logic')) return 'Logic ERP';
        if (v.includes('cider')) return 'Cider';
        if (v.includes('wing')) return 'Wing';
        if (v.includes('gofrugal')) return 'GoFrugal';
        if (v.includes('posist')) return 'Posist';
        if (v.includes('saral')) return 'Saral';
        
        // Return Title Case for others
        return raw.charAt(0).toUpperCase() + raw.slice(1).toLowerCase();
    };

    const posMap = {};

    tickets.forEach(t => {
      let raw = String(t.pos || '').trim();
      const lower = raw.toLowerCase();
      
      // Check excludes loosely
      if (excludePOS.some(ex => lower.includes(ex) && ex.length > 2) || excludePOS.includes(lower)) return;

      const pos = normalizePOS(raw);

      if (!posMap[pos]) {
        posMap[pos] = {
          pos: pos,
          count: 0,
          businesses: new Set()
        };
      }
      posMap[pos].count++;
      posMap[pos].businesses.add(t.business);
    });

    const results = Object.values(posMap)
      .map(p => ({
        pos: p.pos,
        count: p.count,
        businessCount: p.businesses.size
      }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    return JSON.stringify({ success: true, data: results });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}


function getRepeatCustomerAnalysis() {
  try {
    const tickets = getCachedTickets();

    const businessMap = {};
    tickets.forEach(t => {
      if (!t.business || t.business === '-') return;

      if (!businessMap[t.business]) {
        businessMap[t.business] = {
          business: t.business,
          mid: t.mid,
          ticketCount: 0,
          concerns: new Set(),
          agents: new Set(),
          completedCount: 0,
          avgResolutionDays: 0,
          totalDays: 0,
          completedTickets: []
        };
      }
      businessMap[t.business].ticketCount++;
      businessMap[t.business].concerns.add(t.concern);
      businessMap[t.business].agents.add(t.agent);

      if (t.status.toLowerCase() === 'completed') {
        businessMap[t.business].completedCount++;
        businessMap[t.business].totalDays += t.ageDays;
        businessMap[t.business].completedTickets.push(t.ageDays);
      }
    });

    const repeatCustomers = Object.values(businessMap)
      .filter(b => b.ticketCount >= 3)
      .map(b => ({
        business: b.business,
        mid: b.mid,
        ticketCount: b.ticketCount,
        concernCount: b.concerns.size,
        agentCount: b.agents.size,
        completionRate: Math.round((b.completedCount / b.ticketCount) * 100),
        avgResolutionDays: b.completedCount > 0 
          ? Math.round(b.totalDays / b.completedCount) 
          : 0
      }))
      .sort((a, b) => b.ticketCount - a.ticketCount)
      .slice(0, 10);

    return JSON.stringify({ success: true, data: repeatCustomers });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}


function getConcernTrendAnalysis() {
  try {
    const tickets = getCachedTickets();
    const now = new Date();
    const last30Days = now.getTime() - (30 * 24 * 60 * 60 * 1000);
    const last60to30Days = now.getTime() - (60 * 24 * 60 * 60 * 1000);

    const concernMapCurrent = {};
    const concernMapPrevious = {};

    tickets.forEach(t => {
      if (t.sortDate >= last30Days) {
        concernMapCurrent[t.concern] = (concernMapCurrent[t.concern] || 0) + 1;
      } else if (t.sortDate >= last60to30Days && t.sortDate < last30Days) {
        concernMapPrevious[t.concern] = (concernMapPrevious[t.concern] || 0) + 1;
      }
    });

    const trends = Object.keys(concernMapCurrent).map(concern => {
      const current = concernMapCurrent[concern];
      const previous = concernMapPrevious[concern] || 0;
      const change = previous > 0 
        ? Math.round(((current - previous) / previous) * 100)
        : 100;

      return {
        concern: concern,
        currentMonth: current,
        previousMonth: previous,
        changePercent: change,
        trend: change > 10 ? 'rising' : change < -10 ? 'falling' : 'stable'
      };
    }).sort((a, b) => b.currentMonth - a.currentMonth);

    return JSON.stringify({ success: true, data: trends });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}


function getAgentSpecializationMatrix() {
  try {
    const tickets = getCachedTickets();

    const agentConcernMap = {};

    tickets.forEach(t => {
      if (!agentConcernMap[t.agent]) {
        agentConcernMap[t.agent] = {};
      }
      if (!agentConcernMap[t.agent][t.concern]) {
        agentConcernMap[t.agent][t.concern] = {
          total: 0,
          completed: 0
        };
      }
      agentConcernMap[t.agent][t.concern].total++;
      if (t.status.toLowerCase() === 'completed') {
        agentConcernMap[t.agent][t.concern].completed++;
      }
    });

    const specializations = [];
    for (const [agent, concerns] of Object.entries(agentConcernMap)) {
      for (const [concern, stats] of Object.entries(concerns)) {
        if (stats.total >= 3) {
          specializations.push({
            agent: agent,
            concern: concern,
            ticketCount: stats.total,
            completionRate: Math.round((stats.completed / stats.total) * 100),
            expertise: stats.total >= 10 ? 'Expert' : stats.total >= 5 ? 'Experienced' : 'Learning'
          });
        }
      }
    }

    specializations.sort((a, b) => b.completionRate - a.completionRate);

    return JSON.stringify({ success: true, data: specializations });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}


/* ==================================================
   ✅ 7-DAY MINIMUM VALIDATION FUNCTIONS
   Updated: January 3, 2026 - 10:59 PM IST
   ================================================== */

function isInvalidClosedCorrected(ticket, userEmail) {
  if (ticket.status.toLowerCase() !== 'closed') return { isInvalid: false };

  const result = {
    isInvalid: false,
    reason: '',
    warnings: []
  };

  const now = new Date();
  const ticketAgeDays = Math.floor((now - ticket.sortDate) / (1000 * 60 * 60 * 24));

  // ❌ RULE 1: Cannot close before 7 days
  if (ticketAgeDays < 7) {
    const minClosureDate = new Date(ticket.sortDate);
    minClosureDate.setDate(minClosureDate.getDate() + 7);

    result.isInvalid = true;
    result.reason = `Ticket is only ${ticketAgeDays} days old. Minimum 7 days required.`;
    result.warnings.push({
      severity: 'CRITICAL',
      code: 'TOO_EARLY_CLOSURE',
      message: `Cannot close before ${Utilities.formatDate(minClosureDate, "Asia/Kolkata", "dd-MMM-yyyy")}`,
      daysRemaining: 7 - ticketAgeDays,
      action: 'Keep ticket open and continue follow-ups'
    });
    return result;
  }

  // ✅ RULE 2: Manager approval required for 7-10 day old tickets
  const MANAGER_EMAIL = ADMIN_EMAIL; // ✅ Uses the Global Constant from top of file
  const isManager = userEmail === MANAGER_EMAIL;

  if (ticketAgeDays >= 7 && ticketAgeDays <= 10) {
    if (!isManager) {
      result.isInvalid = true;
      result.reason = `Ticket is ${ticketAgeDays} days old. Only manager can close (Day 8-10).`;
      result.warnings.push({
        severity: 'HIGH',
        code: 'MANAGER_APPROVAL_REQUIRED',
        message: `Tickets aged 7-10 days require manager closure`,
        currentUser: userEmail,
        requiredUser: MANAGER_EMAIL,
        action: 'Request manager (Gaurav Pal) to close this ticket'
      });
      return result;
    }
  }

  // ✅ RULE 3: Check follow-up entries exist
  if (!ticket.reason || ticket.reason.trim().length === 0) {
    result.isInvalid = true;
    result.reason = 'No follow-up entries found';
    result.warnings.push({
      severity: 'HIGH',
      code: 'NO_FOLLOWUP',
      message: 'Ticket has no follow-up documentation',
      action: 'Add follow-up entries before closing'
    });
    return result;
  }

  // ✅ RULE 4: Parse and validate follow-up timestamps
  const timestampPattern = /\[(\d{1,2})-([A-Za-z]{3})\s+(\d{1,2}):(\d{2})\]/g;
  const matches = [...ticket.reason.matchAll(timestampPattern)];

  if (matches.length === 0) {
    result.isInvalid = true;
    result.reason = 'No timestamped follow-up entries found';
    result.warnings.push({
      severity: 'HIGH',
      code: 'NO_TIMESTAMPS',
      message: 'Follow-ups must have timestamps in format [dd-MMM HH:mm]',
      action: 'Add timestamped follow-up entries'
    });
    return result;
  }

  // Parse all follow-up dates
  const followUpDates = parseFollowUpDates(matches);

  if (followUpDates.length === 0) {
    result.isInvalid = true;
    result.reason = 'Could not parse follow-up dates';
    return result;
  }

  // Get first and last follow-up
  const firstFollowUp = new Date(Math.min(...followUpDates.map(d => d.getTime())));
  const lastFollowUp = new Date(Math.max(...followUpDates.map(d => d.getTime())));

  // Calculate follow-up duration
  const followUpDurationDays = Math.floor((lastFollowUp - firstFollowUp) / (1000 * 60 * 60 * 24));

  // ❌ RULE 5: Follow-up period must span at least 7 days
  if (followUpDurationDays < 7) {
    result.isInvalid = true;
    result.reason = `Follow-up period is only ${followUpDurationDays} days. Minimum 7 days of follow-up required.`;
    result.warnings.push({
      severity: 'CRITICAL',
      code: 'INSUFFICIENT_FOLLOWUP_DURATION',
      message: `Follow-ups span ${followUpDurationDays} days (${formatDate(firstFollowUp)} to ${formatDate(lastFollowUp)})`,
      required: '7 days',
      actual: followUpDurationDays + ' days',
      action: 'Continue follow-ups until 7 days complete'
    });
    return result;
  }

  // ⚠️  RULE 6: Warn if too few follow-ups (not critical, just warning)
  const expectedMinFollowUps = Math.ceil(followUpDurationDays / 2);
  if (matches.length < expectedMinFollowUps) {
    result.warnings.push({
      severity: 'MEDIUM',
      code: 'FEW_FOLLOWUPS',
      message: `Only ${matches.length} follow-ups in ${followUpDurationDays} days (expected: ${expectedMinFollowUps})`,
      action: 'Add more regular follow-ups for better documentation'
    });
  }

  // ✅ ALL CHECKS PASSED - VALID CLOSURE
  return result;
}


function parseFollowUpDates(matches) {
  const dates = [];
  const currentYear = new Date().getFullYear();
  const monthMap = {
    'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5,
    'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11
  };

  matches.forEach(match => {
    const day = parseInt(match[1]);
    const monthStr = match[2].toLowerCase();
    const hour = parseInt(match[3]);
    const minute = parseInt(match[4]);

    if (monthMap[monthStr] !== undefined) {
      let year = currentYear;
      const currentMonth = new Date().getMonth();

      // Handle year transition (Dec to Jan)
      if (monthMap[monthStr] === 11 && currentMonth <= 1) {
        year = currentYear - 1;
      } else if (monthMap[monthStr] <= 1 && currentMonth === 11) {
        year = currentYear + 1;
      }

      const followUpDate = new Date(year, monthMap[monthStr], day, hour, minute);
      dates.push(followUpDate);
    }
  });

  return dates;
}


function formatDate(date) {
  return Utilities.formatDate(date, "Asia/Kolkata", "dd-MMM-yyyy");
}

// ═══════════════════════════════════════════════════════════════════════
// PAGINATION BACKEND FUNCTION - REQUIRED FOR FRONTEND
// ═══════════════════════════════════════════════════════════════════════

function getTicketsPaginated(config) {
  try {
    requirePermission('VIEW_ANALYTICS');
    const props = PropertiesService.getScriptProperties();
    const rawVersion = props.getProperty('DATA_VERSION');
    const dataVersion = Number.isInteger(Number(rawVersion)) ? Number(rawVersion) : 0;

    // ✅ CACHE IMPLEMENTATION
    const cache = CacheService.getScriptCache();
    const configString = JSON.stringify(config || {});
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, configString);
    const configHash = Utilities.base64Encode(digest);
    const cacheKey = `PAGE_V${dataVersion}_${configHash}`;

    // Try Filtered Cache First
    const cachedResult = cache.get(cacheKey);
    if (cachedResult) {
      return cachedResult;
    }

    // ✅ USE EXISTING getDataObjects() - This has proper date parsing!
    const allTickets = getDataObjects();
    
    if (!allTickets || allTickets.length === 0) {
      return JSON.stringify({
        success: true,
        data: [],
        pagination: { page: 1, pageSize: 100, totalRows: 0, totalPages: 0 },
        version: dataVersion
      });
    }

    const safeConfig = config || {};
    const safeFilters = safeConfig.filters || {};
    const safeSort = safeConfig.sort || {};
    const page = Number(safeConfig.page) > 0 ? Number(safeConfig.page) : 1;
    const pageSize = Math.min(
      CONFIG.MAX_PAGE_SIZE,
      Math.max(1, Number(safeConfig.pageSize) > 0 ? Number(safeConfig.pageSize) : CONFIG.DEFAULT_PAGE_SIZE)
    );
    const statusFilter = safeFilters.status !== undefined ? safeFilters.status : 'all';
    const searchFilter = safeFilters.search !== undefined ? safeFilters.search : '';

    // ✅ NEW: Date range filters (ISO date strings from frontend)
    var dateFromMs = 0;
    var dateToMs = Infinity;
    if (safeFilters.dateFrom) {
      var df2 = new Date(safeFilters.dateFrom);
      if (!isNaN(df2.getTime())) { df2.setHours(0, 0, 0, 0); dateFromMs = df2.getTime(); }
    }
    if (safeFilters.dateTo) {
      var dt2 = new Date(safeFilters.dateTo);
      if (!isNaN(dt2.getTime())) { dt2.setHours(23, 59, 59, 999); dateToMs = dt2.getTime(); }
    }
    var hasDateFilter2 = dateFromMs > 0 || dateToMs < Infinity;

    // ✅ Apply filters using proper normalized data
    let filteredTickets = allTickets.filter(ticket => {
      // Date range filter (if dateFrom/dateTo provided)
      if (hasDateFilter2) {
        if (!ticket.sortDate || ticket.sortDate < dateFromMs || ticket.sortDate > dateToMs) return false;
      }

      // Status filter
      if (statusFilter !== 'all') {
        const requestedStatus = parseStatusOrNull(statusFilter);
        if (requestedStatus && ticket.status !== requestedStatus) return false;
      }
      
      // Search filter
      if (searchFilter && searchFilter.trim() !== '') {
        const term = searchFilter.toLowerCase();
        const searchable = [
          ticket.id, ticket.business, ticket.mid, ticket.agent, 
          ticket.concern, ticket.remark, ticket.reason
        ].join(' ').toLowerCase();
        if (!searchable.includes(term)) return false;
      }
      
      return true;
    });

    // ✅ ALWAYS SORT: Descending by default (newest first)
    const sortOrder = (safeSort.order || 'desc').toLowerCase();
    filteredTickets.sort((a, b) => {
      return sortOrder === 'asc' 
        ? (a.sortDate - b.sortDate) 
        : (b.sortDate - a.sortDate);
    });

    // Calculate pagination
    const totalRows = filteredTickets.length;
    const totalPages = Math.ceil(totalRows / pageSize) || 1;
    const startIdx = (page - 1) * pageSize;
    const endIdx = startIdx + pageSize;
    const pageData = filteredTickets.slice(startIdx, endIdx);

    // ✅ Map to expected frontend format
    const mappedData = pageData.map(t => ({
      'Ticket ID': t.id,
      'Date': t.timestamp, // Native Date object
      'Timestamp': t.timestamp,
      'IT Person Email': t.email,
      'IT Person': t.agent,
      'Requested By': t.requestedBy,
      'MID': t.mid,
      'Business Name': t.business,
      'POS System': t.pos,
      'Support Type': t.supportType,
      'Concern Related to': t.concern,
      'System Configuration': t.config,
      'Remark': t.remark,
      'Status': t.status,
      'Follow-up Reason/ Remark': t.reason,
      'Invalid Closed': t.invalidClosed,
      // Also include computed fields for frontend convenience
      '_sortDate': t.sortDate,
      '_ageDays': t.ageDays,
      '_ageCategory': t.ageCategory,
      '_reasonQuality': t.reasonQuality || 'none'  // ✅ Added
    }));

    const result = {
      success: true,
      data: mappedData,
      pagination: {
        page: page,
        pageSize: pageSize,
        totalRows: totalRows,
        totalPages: totalPages
      },
      version: dataVersion,
      sort: safeSort
    };

    const jsonResult = JSON.stringify(result);

    // ✅ SAVE TO CACHE (5 mins)
    if (jsonResult.length < 100000) {
      cache.put(cacheKey, jsonResult, 300);
    }

    return jsonResult;
    
  } catch (error) {
    Logger.log('Error in getTicketsPaginated: ' + error.toString());
    return JSON.stringify({
      success: false,
      error: error.toString()
    });
  }
}


// ═══════════════════════════════════════════════════════════════════════
// 📊 SERVER-SIDE DASHBOARD STATS (Hybrid Architecture)
// Returns aggregated KPIs + agent stats without raw rows.
// Called by frontend: gas.getDashboardStats({ dateFrom, dateTo })
// ═══════════════════════════════════════════════════════════════════════

/**
 * @param {Object} config
 * @param {string} [config.dateFrom] - ISO date string for range start (inclusive)
 * @param {string} [config.dateTo]   - ISO date string for range end (inclusive)
 * @returns {string} JSON with { success, kpi, agents[], totalTickets, version }
 */
function getDashboardStats(config) {
  try {
    requirePermission('VIEW_ANALYTICS');
    const props = PropertiesService.getScriptProperties();
    const rawVersion = props.getProperty('DATA_VERSION');
    const dataVersion = Number.isInteger(Number(rawVersion)) ? Number(rawVersion) : 0;

    // ── Cache check ──────────────────────────────────────────────────
    const cache = CacheService.getScriptCache();
    const safeConfig = config || {};
    const cacheKey = 'STATS_V' + dataVersion + '_' + (safeConfig.dateFrom || 'all') + '_' + (safeConfig.dateTo || 'all');
    const cached = cache.get(cacheKey);
    if (cached) return cached;

    // ── Load all tickets (uses existing optimized reader) ────────────
    const allTickets = getDataObjects();

    // ── Date range filter ────────────────────────────────────────────
    var dateFromMs = 0;
    var dateToMs = Infinity;
    
    if (safeConfig.dateFrom) {
      var df = new Date(safeConfig.dateFrom);
      if (!isNaN(df.getTime())) {
        df.setHours(0, 0, 0, 0);
        dateFromMs = df.getTime();
      }
    }
    if (safeConfig.dateTo) {
      var dt = new Date(safeConfig.dateTo);
      if (!isNaN(dt.getTime())) {
        dt.setHours(23, 59, 59, 999);
        dateToMs = dt.getTime();
      }
    }

    var hasDateFilter = dateFromMs > 0 || dateToMs < Infinity;

    // ── Single-pass aggregation ──────────────────────────────────────
    var kpi = { total: 0, completed: 0, notCompleted: 0, closed: 0, cantDo: 0 };
    var agentMap = {};
    var totalAllTime = allTickets.length;

    for (var i = 0; i < allTickets.length; i++) {
      var t = allTickets[i];

      // Date filter
      if (hasDateFilter) {
        if (!t.sortDate || t.sortDate < dateFromMs || t.sortDate > dateToMs) continue;
      }

      kpi.total++;

      // Status counts
      if (t.status === STATUS_ENUM.COMPLETED)          kpi.completed++;
      else if (t.status === STATUS_ENUM.NOT_COMPLETED) kpi.notCompleted++;
      else if (t.status === STATUS_ENUM.CLOSED)        kpi.closed++;
      else if (t.status === STATUS_ENUM.CANT_DO)       kpi.cantDo++;

      // Agent stats
      var agentName = t.agent || 'Unassigned';
      if (!agentMap[agentName]) {
        agentMap[agentName] = {
          name: agentName,
          total: 0, completed: 0, notCompleted: 0, closed: 0, cantDo: 0,
          oldNotCompleted: 0
        };
      }
      var a = agentMap[agentName];
      a.total++;
      if (t.status === STATUS_ENUM.COMPLETED)          a.completed++;
      else if (t.status === STATUS_ENUM.NOT_COMPLETED) { a.notCompleted++; if (t.ageDays >= 7) a.oldNotCompleted++; }
      else if (t.status === STATUS_ENUM.CLOSED)        a.closed++;
      else if (t.status === STATUS_ENUM.CANT_DO)       a.cantDo++;
    }

    // ── Compute scores and rank ──────────────────────────────────────
    var agentNames = Object.keys(agentMap);
    var agents = [];
    for (var j = 0; j < agentNames.length; j++) {
      var ag = agentMap[agentNames[j]];
      ag.score = (ag.completed * 10) - (ag.cantDo * 5) - (ag.oldNotCompleted * 3);
      ag.rate = ag.total ? Math.round((ag.completed / ag.total) * 100) : 0;
      agents.push(ag);
    }
    agents.sort(function(a, b) { return b.score - a.score; });
    for (var k = 0; k < agents.length; k++) {
      agents[k].rank = k + 1;
    }

    var result = {
      success: true,
      kpi: kpi,
      agents: agents,
      totalTickets: totalAllTime,
      version: dataVersion
    };

    var jsonResult = JSON.stringify(result);

    // Cache for 5 minutes
    if (jsonResult.length < 100000) {
      cache.put(cacheKey, jsonResult, 300);
    }

    return jsonResult;

  } catch (error) {
    Logger.log('Error in getDashboardStats: ' + error.toString());
    return JSON.stringify({ success: false, error: error.toString() });
  }
}


// ═══════════════════════════════════════════════════════════════════════

// TEST FUNCTION - Run this first to verify it works
// ═══════════════════════════════════════════════════════════════════════

function testGetTicketsPaginated() {
  Logger.log('🧪 Testing pagination function...');
  
  const testConfig = {
    page: 1,
    pageSize: 10,
    filters: {
      search: '',
      status: 'all'
    },
    sort: {
      field: 'Timestamp',
      order: 'desc'
    }
  };
  
  const response = getTicketsPaginated(testConfig);
  
  try {
    const result = JSON.parse(response);  // ✅ FIXED: Parse JSON string first
    
    if (result && result.success) {
      Logger.log('✅✅✅ SUCCESS! Function works!');
      Logger.log('Data rows: ' + result.data.length);
      Logger.log('Total pages: ' + result.pagination.totalPages);
      Logger.log('First ticket: ' + JSON.stringify(result.data[0]));
    } else {
      Logger.log('❌❌❌ FAILED!');
      Logger.log('Error: ' + (result ? result.error : 'null response'));
    }
    
    return result;
  } catch (e) {
    Logger.log('❌❌❌ Parse error: ' + e.toString());
    return null;
  }
}

/* ==================================================
   🛠️ UTILITY: DATA SANITIZATION
   Run this manually once to fix "Can't Do" mismatch
   ================================================== */
function sanitizeDatabase() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log("❌ Sheet not found");
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Status is Column M (Index 13)
  const range = sheet.getRange(2, 13, lastRow - 1, 1); 
  const values = range.getValues();
  
  let updates = 0;
  
  const cleaned = values.map(row => {
    let val = String(row[0]).trim();
    const lower = val.toLowerCase();
    
    // Normalize Can't Do (Smart Quotes, Typos)
    if (lower.includes("cant") || lower.includes("can't") || lower.includes("can’t")) {
      if (val !== "Can't Do") {
        updates++;
        return ["Can't Do"];
      }
    }
    // Normalize Completed
    else if (lower === "completed" && val !== "Completed") { 
        updates++; return ["Completed"]; 
    }
    // Normalize Closed
    else if (lower === "closed" && val !== "Closed") { 
        updates++; return ["Closed"]; 
    }
    // Normalize Pending
    else if (lower === "pending" && val !== "Pending") { 
        updates++; return ["Pending"]; 
    }
    // Normalize Not Completed
    else if ((lower === "not completed" || lower === "notcompleted") && val !== "Not Completed") { 
        updates++; return ["Not Completed"]; 
    }

    return [row[0]]; // Return original if no change
  });
  
  if (updates > 0) {
    range.setValues(cleaned);
    Logger.log(`✅ Sanitized ${updates} rows. Fixed incorrect statuses.`);
  } else {
    Logger.log("✨ Database is already clean.");
  }
}

// ==========================================
// 🧹 CACHE CLEARING UTILITY
// ==========================================
/**
 * Run this function manually after deploying new code to clear cached data.
 * This forces fresh data to be fetched with the new calculations.
 */
function clearAllCache() {
  try {
    // ✅ FIX: Was passing empty array [] to removeAll — which does nothing.
    // Now calls invalidateTicketCache() (clears all chunk keys and metadata)
    // and also clears the ticket index and paginated cache.
    invalidateTicketCache();
    invalidateTicketIndex();
    
    const cache = CacheService.getScriptCache();
    // Clear paginated cache entries (keys are dynamic, so bump version invalidates them)
    cache.remove('HEALTH_CHECK');
    
    // Increment DATA_VERSION to invalidate all version-keyed caches
    const props = PropertiesService.getScriptProperties();
    const currentVersion = parseInt(props.getProperty('DATA_VERSION') || '0');
    props.setProperty('DATA_VERSION', String(currentVersion + 1));
    
    Logger.log('✅ Cache cleared and DATA_VERSION incremented to: ' + (currentVersion + 1));
    return JSON.stringify({ success: true, message: 'Cache cleared successfully' });
  } catch (e) {
    Logger.log('❌ Error clearing cache: ' + e.toString());
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/* ==================================================
   🔧 PHASE 4: ENTERPRISE FEATURES
   ================================================== */

/**
 * 📊 4.1 ENHANCED SYSTEM HEALTH CHECK
 * Comprehensive health monitoring endpoint
 */
function getSystemHealth() {
  const startTime = Date.now();
  const timestamp = new Date().toISOString();
  const health = {
    status: 'healthy',
    timestamp,
    version: CONFIG.APP_VERSION,
    checks: {}
  };
  let dataVersion = '0';
  
  try {
    // Check 1: Spreadsheet Access
    try {
      const ss = getSpreadsheet_();
      const sheet = ss.getSheetByName(SHEET_NAME);
      health.checks.spreadsheet = {
        status: sheet ? 'pass' : 'fail',
        rowCount: sheet ? sheet.getLastRow() : 0,
        message: sheet ? 'Connected' : 'Sheet not found'
      };
    } catch (e) {
      health.checks.spreadsheet = { status: 'fail', message: e.toString() };
      health.status = 'degraded';
    }
    
    // Check 2: Cache Service
    try {
      const cache = CacheService.getScriptCache();
      cache.put('HEALTH_CHECK', 'OK', 10);
      const retrieved = cache.get('HEALTH_CHECK');
      health.checks.cache = {
        status: retrieved === 'OK' ? 'pass' : 'fail',
        message: retrieved === 'OK' ? 'Working' : 'Read mismatch'
      };
    } catch (e) {
      health.checks.cache = { status: 'fail', message: e.toString() };
      health.status = 'degraded';
    }
    
    // Check 3: Properties Service
    try {
      const props = PropertiesService.getScriptProperties();
      const version = props.getProperty('DATA_VERSION');
      dataVersion = version || '0';
      health.checks.properties = {
        status: 'pass',
        dataVersion: dataVersion,
        message: 'Working'
      };
    } catch (e) {
      health.checks.properties = { status: 'fail', message: e.toString() };
      health.status = 'degraded';
    }
    
    // Check 4: Audit Log
    try {
      const auditStats = getAuditLogStats();
      health.checks.auditLog = {
        status: auditStats.exists ? 'pass' : 'warn',
        rowCount: auditStats.rowCount || 0,
        needsRotation: auditStats.needsRotation || false,
        message: auditStats.exists ? 'Operational' : 'Not initialized'
      };
    } catch (e) {
      health.checks.auditLog = { status: 'warn', message: e.toString() };
    }
    
    // Check 5: User Session
    try {
      const userEmail = Session.getActiveUser().getEmail();
      health.checks.session = {
        status: userEmail ? 'pass' : 'warn',
        message: userEmail ? 'Authenticated' : 'Anonymous'
      };
    } catch (e) {
      health.checks.session = { status: 'warn', message: 'Session unavailable' };
    }
    
    // Calculate response time
    health.responseTimeMs = Date.now() - startTime;
    
  } catch (e) {
    health.status = 'unhealthy';
    health.error = e.toString();
  }

  const spreadsheetStatus = health.checks.spreadsheet?.status === 'pass' ? 'OK' : 'ERROR';
  const cacheStatus = health.checks.cache?.status === 'pass' ? 'OK' : 'ERROR';
  const cacheTest = health.checks.cache?.status === 'pass' ? 'HIT' : 'MISS';
  const overallLegacyStatus = health.status === 'healthy'
    ? 'HEALTHY'
    : (health.status === 'degraded' ? 'DEGRADED' : 'ERROR');

  return JSON.stringify({
    success: health.status !== 'unhealthy',
    health: {
      status: overallLegacyStatus,
      version: CONFIG.APP_VERSION,
      dataVersion: parseInt(dataVersion, 10) || 0,
      sheet: {
        status: spreadsheetStatus,
        rows: health.checks.spreadsheet?.rowCount || 0
      },
      cache: {
        status: cacheStatus,
        test: cacheTest
      },
      performance: {
        responseTimeMs: health.responseTimeMs || 0
      },
      timestamp
    },
    status: health.status,
    timestamp,
    version: CONFIG.APP_VERSION,
    checks: health.checks,
    responseTimeMs: health.responseTimeMs || 0,
    error: health.error || null
  });
}

/**
 * 🎚️ 4.2 FEATURE FLAGS SYSTEM
 * Safe feature rollout with PropertiesService-based toggles
 */
const DEFAULT_FEATURE_FLAGS = {
  ENABLE_REAL_TIME_SYNC: true,
  ENABLE_AUDIT_LOGGING: true,
  ENABLE_RATE_LIMITING: true,
  ENABLE_NOTIFICATIONS: true,
  ENABLE_ADVANCED_ANALYTICS: true,
  ENABLE_TICKET_VALIDATION: true,
  MAX_EXPORT_ROWS: 5000,
  CACHE_DURATION_SECONDS: 300
};

/**
 * Get a feature flag value
 */
function getFeatureFlag(flagName) {
  try {
    const props = PropertiesService.getScriptProperties();
    const value = props.getProperty(`FF_${flagName}`);
    
    if (value === null || value === undefined) {
      return DEFAULT_FEATURE_FLAGS[flagName];
    }
    
    // Parse boolean values
    if (value === 'true') return true;
    if (value === 'false') return false;
    
    // Parse numbers
    const num = parseInt(value);
    if (!isNaN(num)) return num;
    
    return value;
  } catch (e) {
    Logger.log(`getFeatureFlag error: ${e.toString()}`);
    return DEFAULT_FEATURE_FLAGS[flagName];
  }
}

/**
 * Set a feature flag value (Admin only)
 */
function setFeatureFlag(flagName, value, csrfToken = '') {
  try {
    requireCSRFToken(csrfToken);
    requirePermission('MANAGE_USERS');
    
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`FF_${flagName}`, String(value));
    
    logAuditEvent('FEATURE_FLAG_CHANGED', null, {
      flag: flagName,
      newValue: value
    });
    
    return JSON.stringify({ 
      success: true, 
      message: `Feature flag ${flagName} set to ${value}` 
    });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * Get all feature flags
 */
function getAllFeatureFlags() {
  try {
    requirePermission('MANAGE_USERS');
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    const flags = {};
    
    // Get default flags
    for (const [key, defaultValue] of Object.entries(DEFAULT_FEATURE_FLAGS)) {
      const storedValue = allProps[`FF_${key}`];
      if (storedValue !== null && storedValue !== undefined) {
        // Parse stored value
        if (storedValue === 'true') flags[key] = true;
        else if (storedValue === 'false') flags[key] = false;
        else if (!isNaN(parseInt(storedValue))) flags[key] = parseInt(storedValue);
        else flags[key] = storedValue;
      } else {
        flags[key] = defaultValue;
      }
    }
    
    return JSON.stringify({ success: true, flags: flags });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 📥 4.3 DATA EXPORT CAPABILITY
 * Export tickets to CSV format
 */
function exportTicketsToCSV(options = {}) {
  try {
    requirePermission('EXPORT_TICKETS');
    requireCSRFToken(options.csrfToken || '');
    const startDate = options.startDate ? new Date(options.startDate) : null;
    const endDate = options.endDate ? new Date(options.endDate) : null;
    const status = options.status || null;
    const maxRows = Math.min(options.maxRows || 1000, getFeatureFlag('MAX_EXPORT_ROWS'));
    
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return JSON.stringify({ success: false, error: 'Sheet not found' });
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return JSON.stringify({ success: true, data: '', rowCount: 0, headers: [] });
    }
    const readCols = Math.min(Math.max(sheet.getLastColumn(), 14), 15);
    const data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();
    
    // Filter data
    let filtered = data;
    
    if (startDate) {
      filtered = filtered.filter(row => {
        const rowDate = new Date(row[1]);
        return rowDate >= startDate;
      });
    }
    
    if (endDate) {
      filtered = filtered.filter(row => {
        const rowDate = new Date(row[1]);
        return rowDate <= endDate;
      });
    }
    
    if (status) {
      filtered = filtered.filter(row => {
        return String(row[12]).toLowerCase() === status.toLowerCase();
      });
    }
    
    // Limit rows
    filtered = filtered.slice(0, maxRows);
    
    // Convert to CSV
    const headers = [
      'Ticket ID', 'Date', 'IT Person', 'IT Person Email', 'Requested By',
      'MID', 'Business Name', 'POS System', 'Support Type', 'Concern',
      'System Configuration', 'Remark', 'Status', 'Follow-up Reason', 'Phone'
    ];
    
    const csvRows = [csvRow_(headers)];
    
    filtered.forEach(row => {
      const normalizedRow = row.slice(0, headers.length);
      while (normalizedRow.length < headers.length) normalizedRow.push('');
      csvRows.push(csvRow_(normalizedRow));
    });
    
    const csvContent = csvRows.join('\n');
    
    logAuditEvent('DATA_EXPORTED', null, {
      rowCount: filtered.length,
      filters: { startDate, endDate, status }
    });
    
    return JSON.stringify({
      success: true,
      data: csvContent,
      rowCount: filtered.length,
      headers: headers
    });
  } catch (e) {
    Logger.log('exportTicketsToCSV error: ' + e.toString());
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 🔐 4.4 ENHANCED TOKEN MANAGEMENT
 * Refresh CSRF token with extended validation
 */
function refreshCSRFToken() {
  try {
    const cache = CacheService.getUserCache();
    
    // Generate new token
    const newToken = Utilities.getUuid();
    const timestamp = Date.now();
    
    // Store with timestamp for validation
    cache.put('CSRF_TOKEN', newToken, 3600);
    cache.put('CSRF_TOKEN_TS', String(timestamp), 3600);
    
    return JSON.stringify({ 
      success: true, 
      token: newToken,
      expiresIn: 3600
    });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * Validate CSRF token with age check
 */
function validateCSRFTokenEnhanced(token) {
  try {
    const cache = CacheService.getUserCache();
    const storedToken = cache.get('CSRF_TOKEN');
    const tokenTs = cache.get('CSRF_TOKEN_TS');
    
    if (!storedToken || !token) {
      return { valid: false, reason: 'Missing token' };
    }
    
    if (storedToken !== token) {
      return { valid: false, reason: 'Token mismatch' };
    }
    
    if (!tokenTs) {
      return { valid: false, reason: 'Token missing timestamp' };
    }

    // Check token age (strict enforcement)
    const parsedTokenTs = parseInt(tokenTs, 10);
    const age = Date.now() - parsedTokenTs;
    const maxAge = 3600 * 1000; // 1 hour
    if (!Number.isFinite(parsedTokenTs) || !Number.isFinite(age) || age < 0 || age > maxAge) {
      return { valid: false, reason: 'Token expired' };
    }
    
    return { valid: true };
  } catch (e) {
    return { valid: false, reason: e.toString() };
  }
}

/* ==================================================
   📜 PHASE 1: UPDATE HISTORY & MONTHLY REPORTS
   Missing Backend Functions Implementation
   ================================================== */

/**
 * 📜 GET UPDATE HISTORY (Paginated Audit Log)
 * Retrieves audit log entries with filtering and pagination
 * @param {Object} config - { page, pageSize, filters: { ticketId, user, action, startDate, endDate, severity } }
 */
function getUpdateHistory(config = {}) {
  try {
    requirePermission('VIEW_AUDIT');
    const page = Math.max(1, parseInt(config.page, 10) || 1);
    const pageSize = Math.min(Math.max(1, parseInt(config.pageSize, 10) || 50), 100);
    const filters = config.filters || {};
    
    const ss = getSpreadsheet_();
    const auditSheet = ss.getSheetByName(CONFIG.AUDIT_SHEET_NAME);
    
    if (!auditSheet || auditSheet.getLastRow() < 2) {
      return JSON.stringify({
        success: true,
        data: [],
        pagination: { page: 1, pageSize, totalRows: 0, totalPages: 0 },
        durationStats: {
          totalWithDuration: 0,
          avgHours: 0,
          fastCount: 0,
          normalCount: 0,
          slowCount: 0,
          criticalCount: 0
        },
        message: 'No history records found. The audit log may not be enabled or is empty.'
      });
    }
    
    const lastRow = auditSheet.getLastRow();
    const rawData = auditSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const statusChangeActions = ['TICKET_UPDATED', 'CLOSE_ATTEMPT_DENIED'];
    const finalStatuses = ['Completed', 'Closed', "Can't Do"];
    const pendingStatuses = ['Not Completed', 'Pending', 'In Progress'];
    
    function parseDetails(detailsRaw) {
      if (!detailsRaw) return {};
      try {
        return typeof detailsRaw === 'string' ? JSON.parse(detailsRaw) : detailsRaw;
      } catch (e) {
        return {};
      }
    }
    
    function toTimestampMs(value) {
      if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
      if (!value) return 0;
      const parsed = new Date(value);
      return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
    }
    
    function formatTimestamp(value) {
      if (value instanceof Date && !isNaN(value.getTime())) {
        return Utilities.formatDate(value, 'Asia/Kolkata', 'dd-MMM-yyyy HH:mm:ss');
      }
      return String(value || '-');
    }
    
    let records = rawData.map((row, idx) => {
      const timestamp = row[0];
      const detailsRaw = row[4];
      const details = parseDetails(detailsRaw);
      const previousStatus = String(details.previousStatus || '-');
      const newStatus = String(details.newStatus || details.status || '-');
      const hasReason = Boolean(
        details.reason ||
        (typeof details.reasonLength === 'number' && details.reasonLength > 0)
      );
      
      return {
        rowNum: idx + 2,
        timestamp: formatTimestamp(timestamp),
        timestampMs: toTimestampMs(timestamp),
        user: String(row[1] || 'Unknown'),
        action: String(row[2] || 'UNKNOWN'),
        ticketId: String(row[3] || '-'),
        details: detailsRaw ? String(detailsRaw) : '',
        severity: String(row[5] || 'INFO'),
        sessionId: String(row[6] || ''),
        version: String(row[7] || ''),
        previousStatus,
        newStatus,
        reasonAdded: hasReason ? 'Yes' : 'No'
      };
    });
    
    records = records.filter(r => statusChangeActions.includes(r.action));
    
    // Build chronological timelines per ticket for duration calculations.
    const ticketTimelines = {};
    records.forEach(record => {
      if (!record.ticketId || record.ticketId === '-' || !record.timestampMs) return;
      if (!ticketTimelines[record.ticketId]) ticketTimelines[record.ticketId] = [];
      ticketTimelines[record.ticketId].push({
        timestampMs: record.timestampMs,
        previousStatus: record.previousStatus,
        newStatus: record.newStatus
      });
    });
    
    Object.keys(ticketTimelines).forEach(ticketId => {
      ticketTimelines[ticketId].sort((a, b) => a.timestampMs - b.timestampMs);
    });
    
    records = records.map(record => {
      let duration = null;
      if (
        record.ticketId &&
        record.ticketId !== '-' &&
        record.timestampMs &&
        pendingStatuses.includes(record.previousStatus) &&
        finalStatuses.includes(record.newStatus)
      ) {
        const timeline = ticketTimelines[record.ticketId] || [];
        let startTime = null;
        
        for (const entry of timeline) {
          if (entry.timestampMs > record.timestampMs) break;
          if (pendingStatuses.includes(entry.newStatus)) {
            startTime = entry.timestampMs;
          }
        }
        
        if (!startTime && timeline.length > 0) {
          startTime = timeline[0].timestampMs;
        }
        
        if (startTime && record.timestampMs >= startTime) {
          duration = formatDurationSla_(record.timestampMs - startTime);
        }
      }
      
      return {
        ...record,
        duration: duration ? duration.formatted : null,
        durationHours: duration ? duration.hours : null,
        durationCategory: duration ? duration.category : null
      };
    });
    
    // Apply filters
    if (filters.ticketId && filters.ticketId.trim() !== '') {
      const searchTerm = filters.ticketId.toLowerCase().trim();
      records = records.filter(r => r.ticketId.toLowerCase().includes(searchTerm));
    }
    
    if (filters.user && filters.user.trim() !== '') {
      const searchTerm = filters.user.toLowerCase().trim();
      records = records.filter(r => r.user.toLowerCase().includes(searchTerm));
    }
    
    if (filters.action && filters.action !== 'all') {
      records = records.filter(r => r.action === filters.action);
    }
    
    if (filters.severity && filters.severity !== 'all') {
      records = records.filter(r => r.severity === filters.severity);
    }
    
    if (filters.startDate) {
      const start = new Date(filters.startDate);
      if (!isNaN(start.getTime())) {
        start.setHours(0, 0, 0, 0);
        const startMs = start.getTime();
        records = records.filter(r => r.timestampMs >= startMs);
      }
    }
    
    if (filters.endDate) {
      const end = new Date(filters.endDate);
      if (!isNaN(end.getTime())) {
        end.setHours(23, 59, 59, 999);
        const endMs = end.getTime();
        records = records.filter(r => r.timestampMs <= endMs);
      }
    }
    
    // Most recent first after all filters.
    records.sort((a, b) => b.timestampMs - a.timestampMs);
    
    const entriesWithDuration = records.filter(r => r.durationHours !== null);
    const durationStats = {
      totalWithDuration: entriesWithDuration.length,
      avgHours: entriesWithDuration.length > 0
        ? Math.round(
          (entriesWithDuration.reduce((sum, r) => sum + (Number(r.durationHours) || 0), 0) / entriesWithDuration.length) * 10
        ) / 10
        : 0,
      fastCount: entriesWithDuration.filter(r => r.durationCategory === 'fast').length,
      normalCount: entriesWithDuration.filter(r => r.durationCategory === 'normal').length,
      slowCount: entriesWithDuration.filter(r => r.durationCategory === 'slow').length,
      criticalCount: entriesWithDuration.filter(r => r.durationCategory === 'critical').length
    };
    
    // Paginate
    const totalRows = records.length;
    const totalPages = Math.ceil(totalRows / pageSize) || 1;
    const validPage = Math.min(Math.max(1, page), totalPages);
    const startIndex = (validPage - 1) * pageSize;
    const pageData = records.slice(startIndex, startIndex + pageSize);
    
    const response = {
      success: true,
      data: pageData,
      pagination: { 
        page: validPage, 
        pageSize, 
        totalRows, 
        totalPages: totalRows === 0 ? 0 : totalPages
      },
      durationStats
    };
    
    if (totalRows === 0) {
      response.message = 'No history records match the selected filters.';
    }
    
    return JSON.stringify(response);
  } catch (e) {
    Logger.log('getUpdateHistory error: ' + e.toString());
    return JSON.stringify({ 
      success: false, 
      error: e.toString(),
      data: [],
      pagination: { page: 1, pageSize: 50, totalRows: 0, totalPages: 0 },
      durationStats: {
        totalWithDuration: 0,
        avgHours: 0,
        fastCount: 0,
        normalCount: 0,
        slowCount: 0,
        criticalCount: 0
      }
    });
  }
}

/**
 * 📊 GENERATE MONTHLY REPORT
 * Creates comprehensive monthly analytics report
 * @param {Object} options - { month: 1-12, year: YYYY }
 */
function generateMonthlyReport(options = {}) {
  try {
    requirePermission('EXPORT_REPORT');
    const now = new Date();
    const month = parseInt(options.month, 10) || (now.getMonth() + 1);
    const year = parseInt(options.year, 10) || now.getFullYear();
    
    // Validate inputs
    if (month < 1 || month > 12) {
      return JSON.stringify({ success: false, error: 'Invalid month. Must be 1-12.' });
    }
    if (year < 2020 || year > 2100) {
      return JSON.stringify({ success: false, error: 'Invalid year.' });
    }
    
    // Date range for the month
    const startDate = new Date(year, month - 1, 1, 0, 0, 0);
    const endDate = new Date(year, month, 0, 23, 59, 59);
    
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return JSON.stringify({ success: false, error: 'Data sheet not found' });
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return JSON.stringify({ success: false, error: 'No ticket data found' });
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    
    // Filter by month
    const monthData = data.filter(row => {
      const dateCell = row[1];
      if (!dateCell) return false;
      const d = dateCell instanceof Date ? dateCell : new Date(dateCell);
      if (isNaN(d.getTime())) return false;
      return d >= startDate && d <= endDate;
    });
    
    if (monthData.length === 0) {
      const monthNames = ['January','February','March','April','May','June',
                          'July','August','September','October','November','December'];
      return JSON.stringify({ 
        success: false, 
        error: `No tickets found for ${monthNames[month-1]} ${year}` 
      });
    }
    
    // Initialize summary
    const summary = {
      totalTickets: monthData.length,
      completed: 0,
      pending: 0,
      closed: 0,
      cantDo: 0,
      invalidClosed: 0,
      avgAgeDays: 0
    };
    
    const agentStats = {};
    const concernStats = {};
    const supportTypeStats = {};
    const dailyStats = {};
    const weekdayStats = {};
    const hourlyStats = {};
    const monthlyTickets = [];
    const weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    let totalAge = 0;
    
    // Process each ticket
    monthData.forEach(row => {
      const status = normalizeStatusWithDefault(row[12]);
      const agent = String(row[3] || 'Unassigned').trim();
      const supportType = String(row[8] || 'Customer Support').trim();
      const concern = String(row[9] || 'Unspecified').trim();
      const dateCell = row[1];
      const d = dateCell instanceof Date ? dateCell : new Date(dateCell);
      const ticketId = String(row[0] || '');
      const mid = String(row[5] || '-');
      const business = String(row[6] || '-');
      const dayKey = d.getDate();
      const weekdayKey = weekdayNames[d.getDay()];
      const hourKey = d.getHours();
      const reason = String(row[13] || '').trim();
      
      // Calculate age
      const age = Math.floor((now.getTime() - d.getTime()) / (1000 * 60 * 60 * 24));
      totalAge += age;
      
      // Check for invalid closed
      let isInvalidClosed = false;
      if (status === STATUS_ENUM.CLOSED) {
        if (age < CONFIG.MIN_CLOSURE_DAYS) {
          isInvalidClosed = true;
        }
      }
      
      // Status counts
      if (status === STATUS_ENUM.COMPLETED) summary.completed++;
      else if (status === STATUS_ENUM.CLOSED) {
        summary.closed++;
        if (isInvalidClosed) summary.invalidClosed++;
      }
      else if (status === STATUS_ENUM.CANT_DO) summary.cantDo++;
      else summary.pending++;
      
      // Agent stats
      if (!agentStats[agent]) {
        agentStats[agent] = { 
          name: agent,
          total: 0, 
          completed: 0, 
          pending: 0, 
          closed: 0, 
          cantDo: 0,
          invalidClosed: 0,
          withReason: 0
        };
      }
      const agentStat = agentStats[agent];
      agentStat.total++;
      if (status === STATUS_ENUM.COMPLETED) agentStat.completed++;
      else if (status === STATUS_ENUM.CLOSED) {
        agentStat.closed++;
        if (isInvalidClosed) agentStat.invalidClosed++;
      }
      else if (status === STATUS_ENUM.CANT_DO) agentStat.cantDo++;
      else agentStat.pending++;
      if (reason.length >= 10) agentStat.withReason++;
      
      // Concern stats
      if (!concernStats[concern]) concernStats[concern] = 0;
      concernStats[concern]++;
      
      // Support type stats
      if (!supportTypeStats[supportType]) supportTypeStats[supportType] = 0;
      supportTypeStats[supportType]++;
      
      // Daily stats
      if (!dailyStats[dayKey]) dailyStats[dayKey] = { created: 0, completed: 0 };
      dailyStats[dayKey].created++;
      if (status === STATUS_ENUM.COMPLETED) dailyStats[dayKey].completed++;
      
      // Weekday distribution for insights panel compatibility
      if (!weekdayStats[weekdayKey]) weekdayStats[weekdayKey] = 0;
      weekdayStats[weekdayKey]++;
      
      // Hourly stats
      if (!hourlyStats[hourKey]) hourlyStats[hourKey] = 0;
      hourlyStats[hourKey]++;
      
      // Ticket list for CSV export compatibility
      monthlyTickets.push({
        id: ticketId,
        date: Utilities.formatDate(d, 'Asia/Kolkata', 'dd-MM-yyyy'),
        agent,
        business,
        mid,
        concern,
        supportType,
        status,
        reason
      });
    });
    
    // Calculate avg age
    summary.avgAgeDays = Math.round(totalAge / Math.max(1, summary.totalTickets));
    
    // Calculate rates
    summary.completionRate = summary.totalTickets > 0 ? 
      Math.round((summary.completed / summary.totalTickets) * 100) : 0;
    summary.resolutionRate = summary.totalTickets > 0 ?
      Math.round(((summary.completed + summary.closed) / summary.totalTickets) * 100) : 0;
    summary.cantDoRate = summary.totalTickets > 0 ?
      Math.round((summary.cantDo / summary.totalTickets) * 100) : 0;
    
    // Performance score calculation
    // Weight: Completion 60%, Low Can't Do 20%, Low Pending 20%
    const pendingRate = summary.totalTickets > 0 ? (summary.pending / summary.totalTickets) * 100 : 0;
    const score = Math.min(100, Math.max(0, 
      (summary.completionRate * 0.6) + 
      ((100 - summary.cantDoRate) * 0.2) +
      ((100 - pendingRate) * 0.2)
    ));
    summary.performanceScore = Math.round(score);
    
    // Performance grade
    if (score >= 90) summary.performanceGrade = 'A+';
    else if (score >= 80) summary.performanceGrade = 'A';
    else if (score >= 70) summary.performanceGrade = 'B';
    else if (score >= 60) summary.performanceGrade = 'C';
    else summary.performanceGrade = 'D';
    
    // Agent rankings with scoring
    const agentRankings = Object.values(agentStats).map(stats => {
      const completionRate = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;
      const reasonRate = stats.total > 0 ? Math.round((stats.withReason / stats.total) * 100) : 0;
      // Score: +10 completed, +5 closed, -5 can't do, -10 invalid closed, -3 pending
      const agentScore = (stats.completed * 10) + (stats.closed * 5) - 
                         (stats.cantDo * 5) - (stats.invalidClosed * 10) - (stats.pending * 3);
      
      return {
        ...stats,
        completionRate,
        reasonRate,
        score: agentScore
      };
    }).sort((a, b) => b.score - a.score);
    
    // Top concerns (top 10)
    const topConcerns = Object.entries(concernStats)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([concern, count]) => ({ 
        concern, 
        count,
        percentage: Math.round((count / summary.totalTickets) * 100)
      }));
    
    const supportTypeBreakdown = Object.entries(supportTypeStats)
      .sort((a, b) => b[1] - a[1])
      .map(([type, count]) => ({
        type,
        count,
        percentage: Math.round((count / summary.totalTickets) * 100)
      }));
    
    // Daily trend (fill all days of month)
    const daysInMonth = new Date(year, month, 0).getDate();
    const dailyTrend = [];
    for (let day = 1; day <= daysInMonth; day++) {
      dailyTrend.push({
        day,
        created: dailyStats[day]?.created || 0,
        completed: dailyStats[day]?.completed || 0
      });
    }
    
    // Hourly distribution
    const hourlyDistribution = [];
    for (let hour = 0; hour < 24; hour++) {
      hourlyDistribution.push({
        hour,
        label: `${hour.toString().padStart(2, '0')}:00`,
        count: hourlyStats[hour] || 0
      });
    }
    
    // Peak hour
    const peakHour = hourlyDistribution.reduce((max, curr) => 
      curr.count > max.count ? curr : max, { hour: 0, count: 0 });
    
    const dailyDistribution = weekdayNames.map(day => ({
      day,
      count: weekdayStats[day] || 0
    }));
    
    const busiestDay = dailyDistribution.reduce((max, curr) =>
      curr.count > max.count ? curr : max, { day: 'N/A', count: 0 });
    
    const activeDays = dailyDistribution.filter(d => d.count > 0);
    const slowestDay = activeDays.length > 0
      ? activeDays.reduce((min, curr) => curr.count < min.count ? curr : min, activeDays[0])
      : { day: 'N/A', count: 0 };
    
    const topPerformer = agentRankings[0] || { name: 'N/A', completed: 0, completionRate: 0, total: 0 };
    const highestRateAgent = agentRankings.reduce((best, current) => {
      if (current.total < 5) return best;
      if (!best || current.completionRate > best.completionRate) return current;
      return best;
    }, null) || topPerformer;
    
    const topConcern = topConcerns[0] || { concern: 'N/A', count: 0, percentage: 0 };
    
    // Generate recommendations
    const recommendations = [];
    
    if (summary.completionRate < 70) {
      recommendations.push({
        priority: 'HIGH',
        category: 'Performance',
        icon: '⚠️',
        message: `Completion rate is ${summary.completionRate}% (below 70% target). Review pending tickets and improve prioritization.`
      });
    }
    
    if (summary.cantDoRate > 10) {
      recommendations.push({
        priority: 'HIGH',
        category: 'Training',
        icon: '📚',
        message: `Can't Do rate is ${summary.cantDoRate}% (above 10% threshold). Consider additional training or better escalation paths.`
      });
    }
    
    if (summary.invalidClosed > 0) {
      recommendations.push({
        priority: 'MEDIUM',
        category: 'Process',
        icon: '🔒',
        message: `${summary.invalidClosed} ticket(s) closed before ${CONFIG.MIN_CLOSURE_DAYS}-day minimum. Enforce closure policy.`
      });
    }
    
    if (topConcerns.length > 0 && topConcerns[0].percentage > 30) {
      recommendations.push({
        priority: 'MEDIUM',
        category: 'Automation',
        icon: '🤖',
        message: `"${topConcerns[0].concern}" represents ${topConcerns[0].percentage}% of tickets. Consider automation or self-service documentation.`
      });
    }
    
    if (summary.avgAgeDays > 10) {
      recommendations.push({
        priority: 'MEDIUM',
        category: 'SLA',
        icon: '⏰',
        message: `Average ticket age is ${summary.avgAgeDays} days. Review SLA adherence and aging ticket management.`
      });
    }
    
    // Achievements
    const achievements = [];
    if (summary.completionRate >= 80) {
      achievements.push({ icon: '🏆', text: 'Excellent completion rate!' });
    }
    if (summary.cantDoRate < 5) {
      achievements.push({ icon: '💪', text: 'Low Can\'t Do rate - great capability!' });
    }
    if (agentRankings.length > 0 && agentRankings[0].completionRate >= 90) {
      achievements.push({ icon: '⭐', text: `Top performer: ${agentRankings[0].name} (${agentRankings[0].completionRate}%)` });
    }
    
    const insights = {
      busiestDay: { day: busiestDay.day, count: busiestDay.count },
      slowestDay: { day: slowestDay.day, count: slowestDay.count },
      topPerformer: {
        name: topPerformer.name,
        completed: topPerformer.completed || 0,
        rate: topPerformer.completionRate || 0
      },
      highestRateAgent: {
        name: highestRateAgent.name,
        rate: highestRateAgent.completionRate || 0,
        total: highestRateAgent.total || 0
      },
      topConcern: {
        name: topConcern.concern,
        count: topConcern.count,
        percentage: topConcern.percentage
      },
      recommendations: recommendations.map(r => ({
        priority: r.priority,
        icon: r.icon,
        message: r.message
      }))
    };

    const monthNames = ['January','February','March','April','May','June',
                        'July','August','September','October','November','December'];
    
    // Log audit event
    logAuditEvent('MONTHLY_REPORT_GENERATED', null, {
      month: monthNames[month-1],
      year: year,
      totalTickets: summary.totalTickets
    });
    
    return JSON.stringify({
      success: true,
      report: {
        title: `Monthly Operations Report - ${monthNames[month-1]} ${year}`,
        generatedAt: new Date().toISOString(),
        generatedBy: Session.getActiveUser().getEmail() || 'System',
        period: {
          month,
          year,
          monthName: monthNames[month-1],
          startDate: Utilities.formatDate(startDate, 'Asia/Kolkata', 'dd-MMM-yyyy'),
          endDate: Utilities.formatDate(endDate, 'Asia/Kolkata', 'dd-MMM-yyyy'),
          daysInMonth
        },
        summary,
        agentRankings,
        topConcerns,
        supportTypeBreakdown,
        insights,
        dailyDistribution,
        dailyTrend,
        hourlyDistribution,
        peakHour: peakHour.label,
        recommendations,
        achievements,
        tickets: monthlyTickets
      }
    });
  } catch (e) {
    Logger.log('generateMonthlyReport error: ' + e.toString());
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 📧 SEND MONTHLY REPORT VIA EMAIL
 * Sends formatted report to admin emails
 */
function sendMonthlyReportEmail(options = {}) {
  try {
    requirePermission('EXPORT_REPORT');
    requireCSRFToken(options.csrfToken || '');
    
    const reportResult = generateMonthlyReport(options);
    const parsed = JSON.parse(reportResult);
    
    if (!parsed.success) {
      return JSON.stringify({ success: false, error: 'Failed to generate report: ' + parsed.error });
    }
    
    const report = parsed.report;
    const recipients = ADMIN_EMAILS.join(',');
    const subject = `📊 ${report.title}`;
    
    // Create HTML email body
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #1E293B; }
          .header { background: linear-gradient(135deg, #10B981, #059669); color: white; padding: 24px; border-radius: 12px; }
          .header h1 { margin: 0; font-size: 24px; }
          .grade-badge { display: inline-block; background: white; color: #059669; padding: 8px 16px; border-radius: 8px; font-size: 24px; font-weight: bold; }
          .summary-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin: 24px 0; }
          .summary-card { background: #F8FAFC; padding: 16px; border-radius: 8px; text-align: center; border: 1px solid #E2E8F0; }
          .summary-value { font-size: 32px; font-weight: bold; color: #4F46E5; }
          .summary-label { font-size: 12px; color: #64748B; text-transform: uppercase; }
          table { width: 100%; border-collapse: collapse; margin: 20px 0; }
          th, td { padding: 12px; text-align: left; border-bottom: 1px solid #E2E8F0; }
          th { background: #F1F5F9; font-weight: 600; }
          .recommendation { padding: 12px 16px; margin: 8px 0; border-radius: 8px; border-left: 4px solid; }
          .rec-high { background: #FEE2E2; border-color: #DC2626; }
          .rec-medium { background: #FEF3C7; border-color: #F59E0B; }
          .rec-low { background: #D1FAE5; border-color: #10B981; }
          .footer { margin-top: 24px; padding-top: 16px; border-top: 1px solid #E2E8F0; color: #64748B; font-size: 12px; }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>📊 ${report.title}</h1>
          <p style="margin: 8px 0 0; opacity: 0.9;">
            ${report.period.startDate} to ${report.period.endDate}
          </p>
        </div>
        
        <div style="display: flex; justify-content: space-between; align-items: center; margin: 24px 0;">
          <div>
            <h2 style="margin: 0;">Performance Grade</h2>
            <p style="color: #64748B; margin: 4px 0;">Based on completion rate, response time, and quality metrics</p>
          </div>
          <div class="grade-badge">${report.summary.performanceGrade}</div>
        </div>
        
        <h3>📈 Executive Summary</h3>
        <table>
          <tr>
            <td><strong>Total Tickets</strong></td>
            <td>${report.summary.totalTickets}</td>
            <td><strong>Completion Rate</strong></td>
            <td>${report.summary.completionRate}%</td>
          </tr>
          <tr>
            <td><strong>Completed</strong></td>
            <td style="color: #10B981;">${report.summary.completed}</td>
            <td><strong>Pending</strong></td>
            <td style="color: #F59E0B;">${report.summary.pending}</td>
          </tr>
          <tr>
            <td><strong>Closed</strong></td>
            <td>${report.summary.closed}</td>
            <td><strong>Can't Do</strong></td>
            <td style="color: #EF4444;">${report.summary.cantDo}</td>
          </tr>
        </table>
        
        <h3>👥 Agent Performance</h3>
        <table>
          <thead>
            <tr>
              <th>Rank</th>
              <th>Agent</th>
              <th>Total</th>
              <th>Completed</th>
              <th>Rate</th>
              <th>Score</th>
            </tr>
          </thead>
          <tbody>
            ${report.agentRankings.map((a, i) => `
              <tr>
                <td>${i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : i + 1}</td>
                <td><strong>${a.name}</strong></td>
                <td>${a.total}</td>
                <td>${a.completed}</td>
                <td>${a.completionRate}%</td>
                <td>${a.score > 0 ? '+' : ''}${a.score}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
        
        ${report.recommendations.length > 0 ? `
          <h3>💡 Recommendations</h3>
          ${report.recommendations.map(r => `
            <div class="recommendation rec-${r.priority.toLowerCase()}">
              <strong>${r.icon} ${r.category}</strong><br>
              ${r.message}
            </div>
          `).join('')}
        ` : ''}
        
        <div class="footer">
          <p>Generated by BillFree TechSupport Ops v${CONFIG.APP_VERSION}</p>
          <p>Report generated on ${new Date().toLocaleString('en-IN')} by ${report.generatedBy}</p>
        </div>
      </body>
      </html>
    `;
    
    GmailApp.sendEmail(recipients, subject, '', { htmlBody });
    
    logAuditEvent('MONTHLY_REPORT_EMAILED', null, {
      recipients: recipients,
      month: report.period.monthName,
      year: report.period.year
    });
    
    return JSON.stringify({ 
      success: true, 
      message: `Report sent to ${recipients}` 
    });
  } catch (e) {
    Logger.log('sendMonthlyReportEmail error: ' + e.toString());
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 📥 EXPORT HISTORY TO CSV
 * Exports update history to CSV format
 */
function exportHistoryToCSV(config = {}) {
  try {
    requirePermission('EXPORT_HISTORY');
    requireCSRFToken(config.csrfToken || '');
    const filters = config.filters || {};
    const statusChangeActions = ['TICKET_UPDATED', 'CLOSE_ATTEMPT_DENIED'];
    
    const ss = getSpreadsheet_();
    const auditSheet = ss.getSheetByName(CONFIG.AUDIT_SHEET_NAME);
    
    if (!auditSheet || auditSheet.getLastRow() < 2) {
      return JSON.stringify({ success: false, error: 'No history data to export' });
    }
    
    const lastRow = auditSheet.getLastRow();
    const rawData = auditSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    let records = rawData.map(row => {
      const rawTs = row[0];
      const timestampMs = rawTs instanceof Date
        ? rawTs.getTime()
        : (new Date(rawTs).getTime() || 0);
      const timestamp = rawTs instanceof Date
        ? Utilities.formatDate(rawTs, 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss')
        : String(rawTs || '');
      
      let previousStatus = '-';
      let newStatus = '-';
      let reasonAdded = 'No';
      
      if (row[4]) {
        try {
          const details = typeof row[4] === 'string' ? JSON.parse(row[4]) : row[4];
          previousStatus = details.previousStatus || '-';
          newStatus = details.newStatus || details.status || '-';
          reasonAdded = (details.reason || (typeof details.reasonLength === 'number' && details.reasonLength > 0))
            ? 'Yes'
            : 'No';
        } catch (e) {}
      }
      
      return {
        timestampMs,
        timestamp,
        user: String(row[1] || ''),
        action: String(row[2] || ''),
        ticketId: String(row[3] || ''),
        previousStatus,
        newStatus,
        severity: String(row[5] || ''),
        reasonAdded
      };
    });
    
    // Apply filters
    records = records.filter(r => statusChangeActions.includes(r.action));
    
    if (filters.ticketId && String(filters.ticketId).trim() !== '') {
      const searchTerm = String(filters.ticketId).toLowerCase().trim();
      records = records.filter(r => r.ticketId.toLowerCase().includes(searchTerm));
    }
    
    if (filters.user && String(filters.user).trim() !== '') {
      const searchTerm = String(filters.user).toLowerCase().trim();
      records = records.filter(r => r.user.toLowerCase().includes(searchTerm));
    }
    
    if (filters.action && filters.action !== 'all') {
      records = records.filter(r => r.action === filters.action);
    }
    
    if (filters.severity && filters.severity !== 'all') {
      records = records.filter(r => r.severity === filters.severity);
    }
    
    if (filters.startDate) {
      const start = new Date(filters.startDate);
      if (!isNaN(start.getTime())) {
        start.setHours(0, 0, 0, 0);
        records = records.filter(r => r.timestampMs >= start.getTime());
      }
    }
    if (filters.endDate) {
      const end = new Date(filters.endDate);
      if (!isNaN(end.getTime())) {
        end.setHours(23, 59, 59, 999);
        records = records.filter(r => r.timestampMs <= end.getTime());
      }
    }
    
    records.sort((a, b) => b.timestampMs - a.timestampMs);
    
    // Generate CSV
    const headers = ['Timestamp', 'User', 'Action', 'Ticket ID', 'Previous Status', 'New Status', 'Severity', 'Reason Added'];
    const csvRows = [csvRow_(headers)];
    
    records.forEach(r => {
      const row = [
        r.timestamp,
        r.user,
        r.action,
        r.ticketId,
        r.previousStatus,
        r.newStatus,
        r.severity,
        r.reasonAdded
      ];
      csvRows.push(csvRow_(row));
    });
    
    logAuditEvent('HISTORY_EXPORTED', null, { rowCount: records.length });
    
    return JSON.stringify({
      success: true,
      csv: csvRows.join('\n'),
      data: csvRows.join('\n'),
      rowCount: records.length,
      filename: `update_history_${Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd_HHmmss')}.csv`
    });
  } catch (e) {
    Logger.log('exportHistoryToCSV error: ' + e.toString());
    return JSON.stringify({ success: false, error: e.toString() });
  }
}
