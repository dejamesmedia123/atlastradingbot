/**
 * ATLASTRADEAI — Backend (Google Apps Script)
 * Data store: Google Sheets. This script is now a pure JSON API — the frontend
 * (index.html = public landing page, app.html = Telegram mini app, admin.html =
 * admin mini app) is static and hosted on GitHub Pages, not served by this script.
 * The static pages call this web app's /exec URL over fetch() (see RPC_FUNCTIONS
 * and doPost below) instead of google.script.run.
 *
 * Nothing needs to be set in the Apps Script editor to run this. Every setting —
 * including the Telegram bot token and Flutterwave keys — is entered from the
 * Admin mini app's Settings screen (super admin only). Admin access itself has
 * no password: it's granted purely by a Telegram ID being listed in the Admins
 * sheet, checked via Telegram's own WebApp identity. The Admins sheet starts
 * empty — the very first person to open the app is auto-granted super admin
 * (see bootstrapFirstAdmin_); after that, only an existing super admin can add more.
 */

// ---------- CONFIG ----------
const SHEET_NAMES = {
  USERS: 'Users', BOTS: 'Bots', TRADES: 'Trades', DEPOSITS: 'Deposits',
  WITHDRAWALS: 'Withdrawals', AFFILIATES: 'Affiliates', REFERRALS: 'Referrals',
  ADMINS: 'Admins', WAITLIST: 'Waitlist', FLAGS: 'Flags', AUDIT_LOG: 'AuditLog', OTP_CODES: 'OtpCodes',
  SETTINGS: 'Settings', FX_RATES: 'FxRates'
};

const SCHEMAS = {
  Users: ['id', 'telegramId', 'name', 'username', 'country', 'referredBy', 'tosAcceptedAt', 'createdAt'],
  Bots: [
    'id', 'name', 'description', 'tvSymbol', 'riskLevel', 'logoUrl', 'status',
    'minDeposit', 'maxDeposit', 'currencies', 'minTopUp', 'withdrawalLockDays', 'minWithdrawal',
    'profitSharePct', 'managementFeePct', 'payoutFrequency',
    'assetClass', 'targetMarkets', 'avgTradeDuration', 'maxDrawdownPct',
    'isPublic', 'countryRestriction', 'maxSubscribers', 'waitlistEnabled',
    'autoPauseDrawdownPct', 'maxAllocationPct',
    'affiliateCommissionOverridePct', 'affiliateEligible',
    'createdAt', 'updatedAt'
  ],
  Trades: ['id', 'botId', 'symbol', 'side', 'entry', 'exit', 'profitPct', 'profitUsd', 'openedAt', 'closedAt', 'createdAt'],
  Deposits: ['id', 'userId', 'telegramId', 'botId', 'amount', 'currency', 'localAmount', 'fxRateUsed', 'method', 'proofUrl', 'txRef', 'status', 'createdAt', 'decidedAt'],
  Withdrawals: ['id', 'userId', 'telegramId', 'botId', 'amount', 'status', 'createdAt', 'decidedAt'],
  Affiliates: [
    'id', 'userId', 'telegramId', 'refCode', 'tier', 'totalReferrals', 'pendingEarnings', 'paidEarnings',
    'bankCode', 'accountNumber', 'accountName', 'bankCountry', 'createdAt'
  ],
  Referrals: ['id', 'referrerTelegramId', 'referredTelegramId', 'depositId', 'commissionAmount', 'status', 'createdAt'],
  Admins: ['telegramId', 'name', 'role', 'addedAt'],
  Waitlist: ['id', 'telegramId', 'botId', 'createdAt'],
  Flags: ['id', 'type', 'telegramId', 'relatedTelegramId', 'details', 'createdAt', 'reviewed'],
  AuditLog: ['id', 'adminTelegramId', 'action', 'targetType', 'targetId', 'details', 'createdAt'],
  OtpCodes: ['id', 'telegramId', 'code', 'purpose', 'expiresAt', 'used', 'createdAt'],
  Settings: ['key', 'value'],
  FxRates: ['currency', 'rate', 'source', 'fetchedAt']
};

// Fallback values used only until an admin saves real ones from the Settings screen.
const DEFAULT_SETTINGS = {
  botUsername: '',
  flutterwavePublicKey: '',
  flutterwaveCountries: 'NG,GH,KE,ZA,UG,TZ,RW,ZM,CI,SN,CM',
  cryptoDepositInstructions: 'Not configured yet — contact support before sending a crypto deposit.',
  directCommissionPct: 15,
  uplineCommissionPct: 3,
  payoutThreshold: 20,
  silverReferralThreshold: 5,
  silverBonusPct: 2,
  goldReferralThreshold: 20,
  goldBonusPct: 5,
  referralVelocityWindowMin: 60,
  referralVelocityThreshold: 5,
  fxMarkupPct: 2,
  fxCacheHours: 24
};
const SECRET_KEYS = ['TELEGRAM_BOT_TOKEN', 'FLW_SECRET_KEY', 'FLW_WEBHOOK_HASH'];

// ---------- SETUP ----------
// This script is standalone (NOT bound to any Google Sheet). The first time anything
// touches the database — opening the web app, calling any function, or manually running
// setupSheets() — getSs_() below creates a fresh Spreadsheet automatically, remembers its
// ID in Script Properties, and builds every tab. Nothing needs to be created by hand first.
const SPREADSHEET_ID_PROP = 'SPREADSHEET_ID';

function getSs_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(SPREADSHEET_ID_PROP);
  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (err) {
      // Stored ID no longer resolves (file trashed/deleted) — fall through and make a new one.
    }
  }
  const ss = SpreadsheetApp.create('AtlasTradeAI DB');
  props.setProperty(SPREADSHEET_ID_PROP, ss.getId());
  ensureSheets_(ss);
  return ss;
}

// Idempotent: safe to call on a brand-new Spreadsheet or an existing one. Creates any
// missing tab/header, never deletes or overwrites existing data rows.
function ensureSheets_(ss) {
  Object.keys(SCHEMAS).forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const headers = SCHEMAS[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  });
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  const existingKeys = settingsSheet.getDataRange().getValues().slice(1).map(function (r) { return r[0]; });
  Object.keys(DEFAULT_SETTINGS).forEach(function (key) {
    if (existingKeys.indexOf(key) === -1) settingsSheet.appendRow([key, DEFAULT_SETTINGS[key]]);
  });
  SpreadsheetApp.flush();
}

// Manual/legacy entry point — no longer required for first-time setup (getSs_() does this
// lazily on first use), but still handy to run once from the editor if you ever want to
// force-create any tab added by a later code update, or just to get the Sheet's URL back.
function setupSheets() {
  const ss = getSs_();
  ensureSheets_(ss);
  return 'Setup complete. Database: ' + ss.getUrl() + '. The first person to open the app (or the /admin URL) is automatically made super admin — no manual sheet editing needed.';
}

// ---------- WEB APP ENTRY ----------
// The frontend is fully static now (GitHub Pages). All RPC calls (reads AND writes) go
// through doGet, as ?fn=functionName&args=<JSON array, URI-encoded>. This is deliberate:
// Apps Script's /exec URL replies with a redirect to actually execute the request, and
// browsers silently convert a POST into a GET (dropping the body) when following that
// redirect — so a POST-based RPC call never actually reaches doPost with its data intact.
// GET isn't affected by that conversion, so it's the reliable choice here.
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.fn) return handleRpc_(params.fn, params.args);
  return jsonOut_({ ok: true, service: 'AtlasTradeAI API', usage: 'GET ?fn=functionName&args=<JSON array> on this same /exec URL.' });
}

function handleRpc_(fn, argsJson) {
  try {
    const args = argsJson ? JSON.parse(argsJson) : [];
    if (!Object.prototype.hasOwnProperty.call(RPC_FUNCTIONS, fn)) {
      return jsonOut_({ ok: false, error: 'Unknown function: ' + fn });
    }
    const result = RPC_FUNCTIONS[fn].apply(null, args);
    return jsonOut_({ ok: true, result: (result === undefined ? null : result) });
  } catch (err) {
    return jsonOut_({ ok: false, error: String((err && err.message) || err) });
  }
}

// Every RPC-callable backend function goes here, keyed by name. This is a whitelist —
// only functions listed here can be invoked from the static frontend; everything else
// (private helpers ending in "_", setupSheets, etc.) stays server-side only.
const RPC_FUNCTIONS = {
  // Auth / users / admins
  isAdmin, getAdminRole, adminAddAdmin, adminListAdmins,
  getOrCreateUser, getUser, updateUserCountry, acceptTerms, getRecommendedPaymentMethod,
  // Bots / trades
  getBots, getBot, adminSaveBot, adminDeleteBot,
  getTrades, adminSaveTrade, adminDeleteTrade, adminBulkImportTrades,
  // Deposits
  createDeposit, getUserDeposits, getPendingDeposits, adminApproveDeposit, adminRejectDeposit,
  confirmFlutterwaveDeposit,
  // Withdrawals
  requestWithdrawalOtp, createWithdrawal, getPendingWithdrawals, adminDecideWithdrawal,
  // Affiliates
  getOrCreateRefCode, getAffiliateStats, saveAffiliatePayoutDetails,
  requestAffiliatePayout, adminGetAffiliates, adminSettlePayout, adminSettlePayoutAutomated,
  // Waitlist / flags / dashboard / stats
  adminGetWaitlist, adminGetFlags, adminMarkFlagReviewed,
  getUserDashboard, getUserStatementCsv, adminGetStats,
  // Settings / FX
  getPublicSettings, adminGetSettings, adminSaveSettings, adminSaveSecrets,
  adminSetFxFallbackRate, getConvertedDepositAmount, adminGetFxRates, adminRefreshFxRates,
  // Audit
  adminGetAuditLog
};

// Flutterwave webhook only — the frontend never calls doPost (see doGet above).
function doPost(e) {
  const source = e && e.parameter && e.parameter.source;
  if (source === 'flw_webhook') return handleFlutterwaveWebhook_(e);
  return jsonOut_({ ok: false, error: 'unknown source' });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- SHEET HELPERS ----------
function getSheet_(name) {
  const ss = getSs_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    // Tab missing (very first call ever, or a tab added by a code update) — build it
    // in place rather than erroring out, then try again.
    ensureSheets_(ss);
    sheet = ss.getSheetByName(name);
  }
  if (!sheet) throw new Error('Sheet not found: ' + name + ' — and auto-setup could not create it.');
  return sheet;
}

function sheetToObjects_(name) {
  const sheet = getSheet_(name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(function (row) { return row.join('') !== ''; })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function appendRow_(name, obj) {
  const sheet = getSheet_(name);
  const headers = SCHEMAS[name];
  sheet.appendRow(headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; }));
  return obj;
}

function findRowIndexById_(name, id) {
  const sheet = getSheet_(name);
  const values = sheet.getDataRange().getValues();
  const idCol = SCHEMAS[name].indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) return i + 1;
  }
  return -1;
}

function updateRowById_(name, id, patch) {
  const sheet = getSheet_(name);
  const headers = SCHEMAS[name];
  const rowIndex = findRowIndexById_(name, id);
  if (rowIndex === -1) throw new Error('Record not found: ' + id);
  const current = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const currentObj = {};
  headers.forEach(function (h, i) { currentObj[h] = current[i]; });
  const merged = Object.assign(currentObj, patch);
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([headers.map(function (h) { return merged[h] !== undefined ? merged[h] : ''; })]);
  return merged;
}

function deleteRowById_(name, id) {
  const sheet = getSheet_(name);
  const rowIndex = findRowIndexById_(name, id);
  if (rowIndex === -1) return false;
  sheet.deleteRow(rowIndex);
  return true;
}

function newId_() { return Utilities.getUuid(); }
function nowIso_() { return new Date().toISOString(); }

// ---------- RATE LIMITING ----------
function checkRateLimit_(key, maxPerMinute) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'rl_' + key;
  const count = Number(cache.get(cacheKey) || 0);
  if (count >= maxPerMinute) throw new Error('Too many requests — please wait a moment and try again.');
  cache.put(cacheKey, String(count + 1), 60);
}

