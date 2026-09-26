// Part of the AtlasTradeAI Apps Script project (split from a single Code.gs into 5 files for readability).
// All .gs files in one Apps Script project share a single global scope, so functions/consts
// in this file freely reference ones in the other 4 files -- file order in the sidebar doesn't matter.

// ---------- AUTH / ROLES ----------
// Bootstrap: if the Admins sheet is completely empty, whoever triggers this next —
// signing up in the user app, or just opening the /admin URL — is auto-granted super
// admin. Only fires once (the moment a single admin row exists, it never fires again),
// so there's no need to hand-edit the sheet for the very first admin anymore.
function bootstrapFirstAdmin_(telegramId, name) {
  if (!telegramId) return false;
  const admins = sheetToObjects_(SHEET_NAMES.ADMINS);
  if (admins.length > 0) return false;
  appendRow_(SHEET_NAMES.ADMINS, { telegramId: telegramId, name: name || '', role: 'super', addedAt: nowIso_() });
  logAudit_(telegramId, 'bootstrap_first_admin', 'admin', telegramId, 'Auto-granted super admin as the first admin ever recorded.');
  return true;
}

function isAdmin(telegramId) {
  const admins = sheetToObjects_(SHEET_NAMES.ADMINS);
  if (admins.length === 0) return bootstrapFirstAdmin_(telegramId);
  return admins.some(function (a) { return String(a.telegramId) === String(telegramId); });
}

function getAdminRole(telegramId) {
  let admins = sheetToObjects_(SHEET_NAMES.ADMINS);
  if (admins.length === 0 && bootstrapFirstAdmin_(telegramId)) admins = sheetToObjects_(SHEET_NAMES.ADMINS);
  const admin = admins.find(function (a) { return String(a.telegramId) === String(telegramId); });
  return admin ? (admin.role || 'support') : null;
}

function requireAdmin_(telegramId) {
  if (!isAdmin(telegramId)) throw new Error('Not authorized');
}

function requireSuperAdmin_(telegramId) {
  if (getAdminRole(telegramId) !== 'super') throw new Error('Requires super admin role');
}

function adminAddAdmin(telegramId, targetTelegramId, name, role) {
  requireSuperAdmin_(telegramId);
  appendRow_(SHEET_NAMES.ADMINS, { telegramId: targetTelegramId, name: name || '', role: role || 'support', addedAt: nowIso_() });
  logAudit_(telegramId, 'add_admin', 'admin', targetTelegramId, role);
  return true;
}

function adminListAdmins(telegramId) {
  requireSuperAdmin_(telegramId);
  return sheetToObjects_(SHEET_NAMES.ADMINS);
}

// ---------- USERS ----------
function getOrCreateUser(profile) {
  const users = sheetToObjects_(SHEET_NAMES.USERS);
  const existing = users.find(function (u) { return String(u.telegramId) === String(profile.telegramId); });
  if (existing) return existing;

  let referredBy = profile.referredBy || '';
  if (referredBy && String(referredBy) === String(profile.telegramId)) {
    flagSuspicious_('self_referral', profile.telegramId, profile.telegramId, 'User attempted to refer themself');
    referredBy = '';
  }

  const user = {
    id: newId_(), telegramId: profile.telegramId, name: profile.name || '', username: profile.username || '',
    country: profile.country || '', referredBy: referredBy, tosAcceptedAt: '', createdAt: nowIso_()
  };
  appendRow_(SHEET_NAMES.USERS, user);
  bootstrapFirstAdmin_(profile.telegramId, profile.name);
  getOrCreateRefCode(profile.telegramId, user.id);
  return user;
}

function getUser(telegramId) {
  return sheetToObjects_(SHEET_NAMES.USERS).find(function (u) { return String(u.telegramId) === String(telegramId); }) || null;
}

function updateUserCountry(telegramId, country) {
  const user = getUser(telegramId);
  if (!user) throw new Error('User not found');
  return updateRowById_(SHEET_NAMES.USERS, user.id, { country: country });
}

function acceptTerms(telegramId) {
  const user = getUser(telegramId);
  if (!user) throw new Error('User not found');
  return updateRowById_(SHEET_NAMES.USERS, user.id, { tosAcceptedAt: nowIso_() });
}

function getRecommendedPaymentMethod(country) {
  const settings = getSettings_();
  return settings.flutterwaveCountries.indexOf(String(country || '').toUpperCase()) !== -1 ? 'flutterwave' : 'manual';
}

// ---------- BOTS ----------
function getBots() {
  return sheetToObjects_(SHEET_NAMES.BOTS).filter(function (b) { return b.status !== 'Deleted'; });
}

function getBot(id) {
  return getBots().find(function (b) { return String(b.id) === String(id); }) || null;
}

function adminSaveBot(telegramId, bot) {
  requireSuperAdmin_(telegramId);
  let saved;
  if (bot.id) {
    bot.updatedAt = nowIso_();
    saved = updateRowById_(SHEET_NAMES.BOTS, bot.id, bot);
    logAudit_(telegramId, 'update_bot', 'bot', bot.id, bot.name);
  } else {
    bot.id = newId_(); bot.createdAt = nowIso_(); bot.updatedAt = nowIso_(); bot.status = bot.status || 'Active';
    saved = appendRow_(SHEET_NAMES.BOTS, bot);
    logAudit_(telegramId, 'create_bot', 'bot', bot.id, bot.name);
  }
  return saved;
}

function adminDeleteBot(telegramId, id) {
  requireSuperAdmin_(telegramId);
  const result = updateRowById_(SHEET_NAMES.BOTS, id, { status: 'Deleted', updatedAt: nowIso_() });
  logAudit_(telegramId, 'delete_bot', 'bot', id, '');
  return result;
}

function getBotSubscriberCount_(botId) {
  const deposits = sheetToObjects_(SHEET_NAMES.DEPOSITS).filter(function (d) { return String(d.botId) === String(botId) && d.status === 'confirmed'; });
  const uniqueUsers = {};
  deposits.forEach(function (d) { uniqueUsers[d.telegramId] = true; });
  return Object.keys(uniqueUsers).length;
}

// ---------- TRADES ----------
function getTrades(botId) {
  return sheetToObjects_(SHEET_NAMES.TRADES)
    .filter(function (t) { return String(t.botId) === String(botId); })
    .sort(function (a, b) { return new Date(b.openedAt) - new Date(a.openedAt); });
}

function adminSaveTrade(telegramId, trade) {
  requireAdmin_(telegramId);
  let saved;
  const isNew = !trade.id;
  if (trade.id) {
    saved = updateRowById_(SHEET_NAMES.TRADES, trade.id, trade);
    logAudit_(telegramId, 'update_trade', 'trade', trade.id, trade.symbol);
  } else {
    trade.id = newId_(); trade.createdAt = nowIso_();
    saved = appendRow_(SHEET_NAMES.TRADES, trade);
    logAudit_(telegramId, 'create_trade', 'trade', trade.id, trade.symbol);
  }
  if (isNew) notifyBotSubscribersNewTrade_(trade);
  return saved;
}

function adminDeleteTrade(telegramId, id) {
  requireAdmin_(telegramId);
  const result = deleteRowById_(SHEET_NAMES.TRADES, id);
  logAudit_(telegramId, 'delete_trade', 'trade', id, '');
  return result;
}

function adminBulkImportTrades(telegramId, botId, csvText) {
  requireAdmin_(telegramId);
  const lines = csvText.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
  let imported = 0;
  lines.forEach(function (line) {
    const parts = line.split(',').map(function (p) { return p.trim(); });
    if (parts.length < 2 || parts[0].toLowerCase() === 'symbol') return;
    const trade = {
      id: newId_(), botId: botId, symbol: parts[0], side: parts[1] || 'Long',
      entry: parts[2] || '', exit: parts[3] || '', profitPct: parts[4] || 0, profitUsd: parts[5] || 0,
      openedAt: parts[6] || nowIso_(), closedAt: parts[7] || '', createdAt: nowIso_()
    };
    appendRow_(SHEET_NAMES.TRADES, trade);
    imported++;
  });
  logAudit_(telegramId, 'bulk_import_trades', 'bot', botId, imported + ' trades imported');
  return { imported: imported };
}

// ---------- DEPOSITS ----------
function createDeposit(payload) {
  checkRateLimit_('deposit_' + payload.telegramId, 10);
  const user = getUser(payload.telegramId);
  if (!user) throw new Error('User not found');
  if (!user.tosAcceptedAt) throw new Error('Please accept the terms before depositing.');

  const bot = getBot(payload.botId);
  if (!bot) throw new Error('Bot not found');

  const amount = Number(payload.amount);
  if (amount < Number(bot.minDeposit) || (bot.maxDeposit && amount > Number(bot.maxDeposit))) {
    throw new Error('Amount must be between ' + bot.minDeposit + ' and ' + bot.maxDeposit);
  }

  if (bot.maxSubscribers) {
    const existingDepositor = sheetToObjects_(SHEET_NAMES.DEPOSITS).some(function (d) {
      return String(d.telegramId) === String(payload.telegramId) && String(d.botId) === String(bot.id) && d.status === 'confirmed';
    });
    if (!existingDepositor && getBotSubscriberCount_(bot.id) >= Number(bot.maxSubscribers)) {
      if (bot.waitlistEnabled) {
        appendRow_(SHEET_NAMES.WAITLIST, { id: newId_(), telegramId: payload.telegramId, botId: bot.id, createdAt: nowIso_() });
        throw new Error('This bot is at capacity — you have been added to the waitlist.');
      }
      throw new Error('This bot is currently at full capacity.');
    }
  }

  const deposit = {
    id: newId_(), userId: user.id, telegramId: payload.telegramId, botId: payload.botId, amount: amount,
    currency: payload.currency || 'USD', localAmount: payload.localAmount != null ? Number(payload.localAmount) : amount,
    fxRateUsed: payload.fxRateUsed != null ? Number(payload.fxRateUsed) : 1,
    method: payload.method, proofUrl: payload.proofUrl || '',
    txRef: payload.txRef || '', status: payload.method === 'flutterwave' ? 'confirmed' : 'pending',
    createdAt: nowIso_(), decidedAt: payload.method === 'flutterwave' ? nowIso_() : ''
  };
  appendRow_(SHEET_NAMES.DEPOSITS, deposit);

  if (deposit.status === 'confirmed') {
    recordReferralOnDeposit_(deposit);
    notifyUser_(deposit.telegramId, 'Deposit confirmed: $' + amount + ' into ' + bot.name + '.');
  }
  return deposit;
}

function getUserDeposits(telegramId) {
  return sheetToObjects_(SHEET_NAMES.DEPOSITS)
    .filter(function (d) { return String(d.telegramId) === String(telegramId); })
    .sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); });
}

function getPendingDeposits(telegramId) {
  requireAdmin_(telegramId);
  return sheetToObjects_(SHEET_NAMES.DEPOSITS).filter(function (d) { return d.status === 'pending'; });
}

function adminApproveDeposit(telegramId, depositId) {
  requireAdmin_(telegramId);
  const updated = updateRowById_(SHEET_NAMES.DEPOSITS, depositId, { status: 'confirmed', decidedAt: nowIso_() });
  recordReferralOnDeposit_(updated);
  notifyUser_(updated.telegramId, 'Your deposit of $' + updated.amount + ' has been confirmed.');
  logAudit_(telegramId, 'approve_deposit', 'deposit', depositId, '');
  return updated;
}

function adminRejectDeposit(telegramId, depositId) {
  requireAdmin_(telegramId);
  const updated = updateRowById_(SHEET_NAMES.DEPOSITS, depositId, { status: 'rejected', decidedAt: nowIso_() });
  notifyUser_(updated.telegramId, 'Your deposit submission could not be verified. Please contact support.');
  logAudit_(telegramId, 'reject_deposit', 'deposit', depositId, '');
  return updated;
}

