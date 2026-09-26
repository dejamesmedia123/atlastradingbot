// Part of the AtlasTradeAI Apps Script project (split from a single Code.gs into 5 files for readability).
// All .gs files in one Apps Script project share a single global scope, so functions/consts
// in this file freely reference ones in the other 4 files -- file order in the sidebar doesn't matter.

// ---------- FLUTTERWAVE ----------
function verifyFlutterwaveTransaction_(txRef, expectedAmount) {
  const secret = PropertiesService.getScriptProperties().getProperty('FLW_SECRET_KEY');
  if (!secret) throw new Error('Flutterwave not configured (FLW_SECRET_KEY missing).');
  const url = 'https://api.flutterwave.com/v3/transactions/verify_by_reference?tx_ref=' + encodeURIComponent(txRef);
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + secret }, muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  if (data.status !== 'success' || !data.data) return { verified: false, reason: 'Transaction not found' };
  if (data.data.status !== 'successful') return { verified: false, reason: 'Transaction not successful' };
  if (expectedAmount && Number(data.data.amount) < Number(expectedAmount)) return { verified: false, reason: 'Amount mismatch' };
  return { verified: true, data: data.data };
}

// Call this from the client after Flutterwave's inline checkout callback fires,
// instead of trusting the client-side "successful" flag directly.
// payload: { telegramId, botId, amount (USD, canonical ledger amount), currency, localAmount (what was actually charged), txRef }
function confirmFlutterwaveDeposit(payload) {
  const chargedAmount = payload.currency && payload.currency !== 'USD' ? payload.localAmount : payload.amount;
  const check = verifyFlutterwaveTransaction_(payload.txRef, chargedAmount);
  if (!check.verified) throw new Error('Payment verification failed: ' + check.reason);
  const fxRateUsed = payload.localAmount && payload.amount ? (Number(payload.localAmount) / Number(payload.amount)) : 1;
  return createDeposit(Object.assign({}, payload, { method: 'flutterwave', fxRateUsed: fxRateUsed }));
}

function handleFlutterwaveWebhook_(e) {
  const receivedHash = e.parameter['verif-hash'] || (e.headers && e.headers['verif-hash']);
  const expectedHash = PropertiesService.getScriptProperties().getProperty('FLW_WEBHOOK_HASH');
  if (!expectedHash || receivedHash !== expectedHash) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'invalid signature' })).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const body = JSON.parse(e.postData.contents);
    const txRef = body.data && body.data.tx_ref;
    const deposits = sheetToObjects_(SHEET_NAMES.DEPOSITS);
    const match = deposits.find(function (d) { return d.txRef === txRef && d.status === 'pending'; });
    if (match && body.data.status === 'successful') {
      const updated = updateRowById_(SHEET_NAMES.DEPOSITS, match.id, { status: 'confirmed', decidedAt: nowIso_() });
      recordReferralOnDeposit_(updated);
      notifyUser_(updated.telegramId, 'Deposit confirmed: $' + updated.amount + '.');
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- WITHDRAWALS (with OTP two-factor) ----------
function requestWithdrawalOtp(telegramId) {
  checkRateLimit_('otp_' + telegramId, 5);
  const code = String(Math.floor(100000 + Math.random() * 900000));
  appendRow_(SHEET_NAMES.OTP_CODES, {
    id: newId_(), telegramId: telegramId, code: code, purpose: 'withdrawal',
    expiresAt: new Date(Date.now() + 5 * 60000).toISOString(), used: false, createdAt: nowIso_()
  });
  notifyUser_(telegramId, 'Your AtlasTradeAI withdrawal confirmation code is ' + code + '. It expires in 5 minutes.');
  return { sent: true };
}

function verifyOtp_(telegramId, purpose, code) {
  const codes = sheetToObjects_(SHEET_NAMES.OTP_CODES).filter(function (o) {
    return String(o.telegramId) === String(telegramId) && o.purpose === purpose && String(o.code) === String(code) && o.used !== true && o.used !== 'TRUE';
  });
  const valid = codes.find(function (o) { return new Date(o.expiresAt) > new Date(); });
  if (!valid) return false;
  updateRowById_(SHEET_NAMES.OTP_CODES, valid.id, { used: true });
  return true;
}

function createWithdrawal(payload) {
  checkRateLimit_('withdraw_' + payload.telegramId, 5);
  const user = getUser(payload.telegramId);
  if (!user) throw new Error('User not found');
  const bot = getBot(payload.botId);
  if (!bot) throw new Error('Bot not found');
  if (Number(payload.amount) < Number(bot.minWithdrawal)) throw new Error('Minimum withdrawal is ' + bot.minWithdrawal);

  if (!payload.otpCode || !verifyOtp_(payload.telegramId, 'withdrawal', payload.otpCode)) {
    throw new Error('Invalid or expired confirmation code. Request a new one.');
  }

  const dashboard = getUserDashboard(payload.telegramId);
  const position = dashboard.positions.find(function (p) { return String(p.botId) === String(payload.botId); });
  if (!position || position.daysSinceDeposit < Number(bot.withdrawalLockDays || 0)) {
    throw new Error('Withdrawal is locked for ' + bot.withdrawalLockDays + ' days from deposit');
  }
  if (Number(payload.amount) > position.balance) throw new Error('Amount exceeds your available balance.');

  const withdrawal = {
    id: newId_(), userId: user.id, telegramId: payload.telegramId, botId: payload.botId,
    amount: Number(payload.amount), status: 'pending', createdAt: nowIso_(), decidedAt: ''
  };
  appendRow_(SHEET_NAMES.WITHDRAWALS, withdrawal);
  return withdrawal;
}

function getPendingWithdrawals(telegramId) {
  requireAdmin_(telegramId);
  return sheetToObjects_(SHEET_NAMES.WITHDRAWALS).filter(function (w) { return w.status === 'pending'; });
}

function adminDecideWithdrawal(telegramId, withdrawalId, approve) {
  requireAdmin_(telegramId);
  const updated = updateRowById_(SHEET_NAMES.WITHDRAWALS, withdrawalId, { status: approve ? 'paid' : 'rejected', decidedAt: nowIso_() });
  notifyUser_(updated.telegramId, approve ? ('Your withdrawal of $' + updated.amount + ' has been paid.') : 'Your withdrawal request was rejected. Contact support for details.');
  logAudit_(telegramId, approve ? 'approve_withdrawal' : 'reject_withdrawal', 'withdrawal', withdrawalId, '');
  return updated;
}

// ---------- AFFILIATES ----------
function getOrCreateRefCode(telegramId, userId) {
  const affiliates = sheetToObjects_(SHEET_NAMES.AFFILIATES);
  const existing = affiliates.find(function (a) { return String(a.telegramId) === String(telegramId); });
  if (existing) return existing;
  const refCode = 'ref_' + Utilities.getUuid().split('-')[0].toUpperCase();
  const affiliate = {
    id: newId_(), userId: userId || '', telegramId: telegramId, refCode: refCode, tier: 'Bronze',
    totalReferrals: 0, pendingEarnings: 0, paidEarnings: 0,
    bankCode: '', accountNumber: '', accountName: '', bankCountry: '', createdAt: nowIso_()
  };
  appendRow_(SHEET_NAMES.AFFILIATES, affiliate);
  return affiliate;
}

function getAffiliateStats(telegramId) {
  const affiliate = getOrCreateRefCode(telegramId);
  const referrals = sheetToObjects_(SHEET_NAMES.REFERRALS)
    .filter(function (r) { return String(r.referrerTelegramId) === String(telegramId); })
    .sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); });
  return { affiliate: affiliate, referrals: referrals };
}

function saveAffiliatePayoutDetails(telegramId, details) {
  const affiliate = getOrCreateRefCode(telegramId);
  return updateRowById_(SHEET_NAMES.AFFILIATES, affiliate.id, {
    bankCode: details.bankCode || '', accountNumber: details.accountNumber || '',
    accountName: details.accountName || '', bankCountry: details.bankCountry || ''
  });
}

function computeTier_(totalReferrals, settings) {
  const rules = getTierRules_(settings || getSettings_());
  const rule = rules.find(function (r) { return totalReferrals >= r.min; });
  return rule || rules[rules.length - 1];
}

function flagSuspicious_(type, telegramId, relatedTelegramId, details) {
  appendRow_(SHEET_NAMES.FLAGS, {
    id: newId_(), type: type, telegramId: telegramId, relatedTelegramId: relatedTelegramId || '',
    details: details || '', createdAt: nowIso_(), reviewed: false
  });
}

function recordReferralOnDeposit_(deposit) {
  const user = sheetToObjects_(SHEET_NAMES.USERS).find(function (u) { return String(u.telegramId) === String(deposit.telegramId); });
  if (!user || !user.referredBy) return;

  const alreadyRecorded = sheetToObjects_(SHEET_NAMES.REFERRALS).some(function (r) { return String(r.depositId) === String(deposit.id); });
  if (alreadyRecorded) return;

  const bot = getBot(deposit.botId);
  const affiliateEligible = !bot || bot.affiliateEligible === true || bot.affiliateEligible === 'TRUE' || bot.affiliateEligible === '';
  if (!affiliateEligible) return;

  const settings = getSettings_();

  const recentReferrals = sheetToObjects_(SHEET_NAMES.REFERRALS).filter(function (r) {
    return String(r.referrerTelegramId) === String(user.referredBy) &&
      (new Date() - new Date(r.createdAt)) / 60000 < settings.referralVelocityWindowMin;
  });
  if (recentReferrals.length >= settings.referralVelocityThreshold) {
    flagSuspicious_('referral_velocity', user.referredBy, deposit.telegramId, recentReferrals.length + ' referrals in ' + settings.referralVelocityWindowMin + ' min');
  }

  const directAffiliate = getOrCreateRefCode(user.referredBy);
  const tierInfo = computeTier_(Number(directAffiliate.totalReferrals || 0), settings);
  const overridePct = bot && bot.affiliateCommissionOverridePct ? Number(bot.affiliateCommissionOverridePct) : settings.directCommissionPct;
  const effectivePct = overridePct + tierInfo.bonusPct;
  const directCommission = Number(deposit.amount) * (effectivePct / 100);

  appendRow_(SHEET_NAMES.REFERRALS, {
    id: newId_(), referrerTelegramId: user.referredBy, referredTelegramId: deposit.telegramId,
    depositId: deposit.id, commissionAmount: directCommission, status: 'pending', createdAt: nowIso_()
  });
  const newTotal = Number(directAffiliate.totalReferrals || 0) + 1;
  updateRowById_(SHEET_NAMES.AFFILIATES, directAffiliate.id, {
    totalReferrals: newTotal, tier: computeTier_(newTotal, settings).name,
    pendingEarnings: Number(directAffiliate.pendingEarnings || 0) + directCommission
  });
  notifyUser_(user.referredBy, 'You earned $' + directCommission.toFixed(2) + ' commission from a referral deposit.');

  const uplineUser = sheetToObjects_(SHEET_NAMES.USERS).find(function (u) { return String(u.telegramId) === String(user.referredBy); });
  if (uplineUser && uplineUser.referredBy) {
    const uplineCommission = Number(deposit.amount) * (settings.uplineCommissionPct / 100);
    appendRow_(SHEET_NAMES.REFERRALS, {
      id: newId_(), referrerTelegramId: uplineUser.referredBy, referredTelegramId: deposit.telegramId,
      depositId: deposit.id, commissionAmount: uplineCommission, status: 'pending', createdAt: nowIso_()
    });
    const uplineAffiliate = getOrCreateRefCode(uplineUser.referredBy);
    updateRowById_(SHEET_NAMES.AFFILIATES, uplineAffiliate.id, { pendingEarnings: Number(uplineAffiliate.pendingEarnings || 0) + uplineCommission });
    notifyUser_(uplineUser.referredBy, 'You earned $' + uplineCommission.toFixed(2) + ' upline commission.');
  }
}
