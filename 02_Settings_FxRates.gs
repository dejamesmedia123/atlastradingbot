// Part of the AtlasTradeAI Apps Script project (split from a single Code.gs into 5 files for readability).
// All .gs files in one Apps Script project share a single global scope, so functions/consts
// in this file freely reference ones in the other 4 files -- file order in the sidebar doesn't matter.

// ---------- AUDIT LOG ----------
function logAudit_(telegramId, action, targetType, targetId, details) {
  appendRow_(SHEET_NAMES.AUDIT_LOG, {
    id: newId_(), adminTelegramId: telegramId, action: action, targetType: targetType || '',
    targetId: targetId || '', details: details || '', createdAt: nowIso_()
  });
}

function adminGetAuditLog(telegramId) {
  requireAdmin_(telegramId);
  return sheetToObjects_(SHEET_NAMES.AUDIT_LOG)
    .sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); })
    .slice(0, 200);
}

// ---------- SETTINGS ----------
function getSettings_() {
  const rows = sheetToObjects_(SHEET_NAMES.SETTINGS);
  const map = Object.assign({}, DEFAULT_SETTINGS);
  rows.forEach(function (r) { if (r.key) map[r.key] = r.value; });
  return {
    botUsername: String(map.botUsername || ''),
    flutterwavePublicKey: String(map.flutterwavePublicKey || ''),
    flutterwaveCountries: String(map.flutterwaveCountries || '').split(',').map(function (c) { return c.trim().toUpperCase(); }).filter(Boolean),
    directCommissionPct: Number(map.directCommissionPct),
    uplineCommissionPct: Number(map.uplineCommissionPct),
    payoutThreshold: Number(map.payoutThreshold),
    silverReferralThreshold: Number(map.silverReferralThreshold),
    silverBonusPct: Number(map.silverBonusPct),
    goldReferralThreshold: Number(map.goldReferralThreshold),
    goldBonusPct: Number(map.goldBonusPct),
    referralVelocityWindowMin: Number(map.referralVelocityWindowMin),
    referralVelocityThreshold: Number(map.referralVelocityThreshold),
    fxMarkupPct: Number(map.fxMarkupPct),
    fxCacheHours: Number(map.fxCacheHours)
  };
}

function getTierRules_(settings) {
  return [
    { min: settings.goldReferralThreshold, name: 'Gold', bonusPct: settings.goldBonusPct },
    { min: settings.silverReferralThreshold, name: 'Silver', bonusPct: settings.silverBonusPct },
    { min: 0, name: 'Bronze', bonusPct: 0 }
  ];
}

// Public — no auth. Used by the user mini app to build referral links and pick a payment method.
function getPublicSettings() {
  const s = getSettings_();
  return { botUsername: s.botUsername, flutterwaveCountries: s.flutterwaveCountries, flutterwavePublicKey: s.flutterwavePublicKey };
}

function adminGetSettings(telegramId) {
  requireSuperAdmin_(telegramId);
  const settings = getSettings_();
  const props = PropertiesService.getScriptProperties();
  const secretsStatus = {};
  SECRET_KEYS.forEach(function (k) { secretsStatus[k] = !!props.getProperty(k); });
  return { settings: settings, secretsStatus: secretsStatus, spreadsheetUrl: getSs_().getUrl() };
}

function adminSaveSettings(telegramId, settingsObj) {
  requireSuperAdmin_(telegramId);
  const sheet = getSheet_(SHEET_NAMES.SETTINGS);
  const values = sheet.getDataRange().getValues();
  Object.keys(settingsObj).forEach(function (key) {
    let found = false;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === key) { sheet.getRange(i + 1, 2).setValue(settingsObj[key]); found = true; break; }
    }
    if (!found) sheet.appendRow([key, settingsObj[key]]);
  });
  logAudit_(telegramId, 'update_settings', 'settings', '', Object.keys(settingsObj).join(', '));
  return getSettings_();
}

// Secrets are stored in Script Properties (not the Sheet) even though they're entered
// from the admin UI, so they're never visible to anyone with view access to the Spreadsheet.
function adminSaveSecrets(telegramId, secretsObj) {
  requireSuperAdmin_(telegramId);
  const props = PropertiesService.getScriptProperties();
  const updatedKeys = [];
  SECRET_KEYS.forEach(function (key) {
    if (secretsObj[key]) { props.setProperty(key, secretsObj[key]); updatedKeys.push(key); }
  });
  logAudit_(telegramId, 'update_secrets', 'settings', '', updatedKeys.join(', ') + ' updated');
  return true;
}

// ---------- FX RATES (TradingView scanner, cached ~24h, with admin fallback) ----------
// NOTE: scanner.tradingview.com is TradingView's own internal, undocumented endpoint —
// not an official public API. It works without a key, but can change or block requests
// without notice, and doesn't reliably cover every exotic currency. That's why every
// call here falls back to an admin-set rate if it fails. Admin-set fallback rates live
// in the SAME sheet as the live TradingView cache (FxRates), just tagged with
// source = 'admin_fallback' instead of 'tradingview' — see adminSetFxFallbackRate.
function fetchTradingViewRate_(currencyCode) {
  try {
    const ticker = 'FX_IDC:USD' + currencyCode;
    const url = 'https://scanner.tradingview.com/symbol?symbol=' + encodeURIComponent(ticker) + '&fields=lp';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (res.getResponseCode() !== 200) return null;
    const data = JSON.parse(res.getContentText());
    const rate = Number(data.lp || data.close || data.last_price);
    return (rate && rate > 0) ? rate : null;
  } catch (err) {
    return null;
  }
}

function getFxRate_(currencyCode) {
  const currency = String(currencyCode).toUpperCase();
  if (currency === 'USD') return { rate: 1, source: 'base', fetchedAt: nowIso_() };

  const settings = getSettings_();
  const cached = sheetToObjects_(SHEET_NAMES.FX_RATES).find(function (r) { return String(r.currency).toUpperCase() === currency; });
  const cacheAgeHours = cached ? (new Date() - new Date(cached.fetchedAt)) / 3600000 : Infinity;

  if (cached && cacheAgeHours < settings.fxCacheHours) {
    return { rate: Number(cached.rate), source: cached.source, fetchedAt: cached.fetchedAt };
  }

  const liveRate = fetchTradingViewRate_(currency);
  if (liveRate) {
    setFxCacheRow_(currency, liveRate, 'tradingview');
    return { rate: liveRate, source: 'tradingview', fetchedAt: nowIso_() };
  }

  // Live fetch failed — use whatever's in the FxRates sheet, however old, be it a stale
  // TradingView read or an admin-entered fallback (both live in that same sheet/row).
  if (cached) {
    const source = cached.source === 'admin_fallback' ? cached.source : cached.source + '_stale';
    return { rate: Number(cached.rate), source: source, fetchedAt: cached.fetchedAt };
  }

  throw new Error('No exchange rate available for ' + currency + ' — ask an admin to set a fallback rate in Settings → Live FX rate cache.');
}

// Admin sets/updates a manual fallback rate. Writes into the SAME FxRates sheet/row the
// TradingView scanner uses (via setFxCacheRow_), tagged source = 'admin_fallback', so
// there's one single place — one row per currency — holding whichever rate is current.
// A later successful TradingView fetch will overwrite this row with a 'tradingview' rate;
// if TradingView then fails again, getFxRate_ falls back to whatever this row last held.
function adminSetFxFallbackRate(telegramId, currency, rate) {
  requireSuperAdmin_(telegramId);
  const cur = String(currency).toUpperCase();
  const num = Number(rate);
  if (!cur || !num || num <= 0) throw new Error('Provide a currency code and a positive rate.');
  setFxCacheRow_(cur, num, 'admin_fallback');
  logAudit_(telegramId, 'set_fx_fallback', 'fx_rate', cur, String(num));
  return sheetToObjects_(SHEET_NAMES.FX_RATES);
}

function setFxCacheRow_(currency, rate, source) {
  const sheet = getSheet_(SHEET_NAMES.FX_RATES);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toUpperCase() === currency) {
      sheet.getRange(i + 1, 1, 1, 4).setValues([[currency, rate, source, nowIso_()]]);
      return;
    }
  }
  sheet.appendRow([currency, rate, source, nowIso_()]);
}

// Public — no auth. Client calls this before opening Flutterwave checkout to know what to charge.
function getConvertedDepositAmount(usdAmount, currency) {
  const cur = String(currency).toUpperCase();
  if (cur === 'USD') return { localAmount: Number(usdAmount), rate: 1, markupPct: 0, source: 'base' };
  const settings = getSettings_();
  const fx = getFxRate_(cur);
  const effectiveRate = fx.rate * (1 + settings.fxMarkupPct / 100);
  const localAmount = Math.round(Number(usdAmount) * effectiveRate * 100) / 100;
  return { localAmount: localAmount, rate: effectiveRate, markupPct: settings.fxMarkupPct, source: fx.source };
}

function adminGetFxRates(telegramId) {
  requireAdmin_(telegramId);
  return sheetToObjects_(SHEET_NAMES.FX_RATES);
}

function adminRefreshFxRates(telegramId) {
  requireSuperAdmin_(telegramId);
  const results = [];
  const currencies = Array.from(new Set(
    sheetToObjects_(SHEET_NAMES.BOTS).flatMap(function (b) { return String(b.currencies || '').split(',').map(function (c) { return c.trim().toUpperCase(); }); })
      .filter(function (c) { return c && c !== 'USD'; })
  ));
  currencies.forEach(function (cur) {
    const liveRate = fetchTradingViewRate_(cur);
    if (liveRate) { setFxCacheRow_(cur, liveRate, 'tradingview'); results.push({ currency: cur, rate: liveRate, ok: true }); }
    else results.push({ currency: cur, ok: false });
  });
  logAudit_(telegramId, 'refresh_fx_rates', 'fx', '', currencies.join(', '));
  return results;
}

