// Part of the AtlasTradeAI Apps Script project (split from a single Code.gs into 5 files for readability).
// All .gs files in one Apps Script project share a single global scope, so functions/consts
// in this file freely reference ones in the other 4 files -- file order in the sidebar doesn't matter.


function requestAffiliatePayout(telegramId) {
  const affiliate = getOrCreateRefCode(telegramId);
  const threshold = getSettings_().payoutThreshold;
  if (Number(affiliate.pendingEarnings) < threshold) throw new Error('Minimum payout threshold is $' + threshold);
  notifyUser_(telegramId, 'Payout of $' + Number(affiliate.pendingEarnings).toFixed(2) + ' requested — an admin will settle it shortly.');
  return true;
}

function adminGetAffiliates(telegramId) {
  requireAdmin_(telegramId);
  return sheetToObjects_(SHEET_NAMES.AFFILIATES);
}

function adminSettlePayout(telegramId, affiliateId, amount) {
  requireAdmin_(telegramId);
  const affiliate = sheetToObjects_(SHEET_NAMES.AFFILIATES).find(function (a) { return String(a.id) === String(affiliateId); });
  if (!affiliate) throw new Error('Affiliate not found');
  const settleAmount = Number(amount);
  const updated = updateRowById_(SHEET_NAMES.AFFILIATES, affiliateId, {
    pendingEarnings: Math.max(0, Number(affiliate.pendingEarnings) - settleAmount),
    paidEarnings: Number(affiliate.paidEarnings || 0) + settleAmount
  });
  notifyUser_(affiliate.telegramId, 'Your affiliate payout of $' + settleAmount.toFixed(2) + ' has been settled.');
  logAudit_(telegramId, 'settle_payout_manual', 'affiliate', affiliateId, String(settleAmount));
  return updated;
}

function adminSettlePayoutAutomated(telegramId, affiliateId, amount) {
  requireSuperAdmin_(telegramId);
  const affiliate = sheetToObjects_(SHEET_NAMES.AFFILIATES).find(function (a) { return String(a.id) === String(affiliateId); });
  if (!affiliate) throw new Error('Affiliate not found');
  if (!affiliate.accountNumber || !affiliate.bankCode) throw new Error('Affiliate has not saved payout bank details.');

  const secret = PropertiesService.getScriptProperties().getProperty('FLW_SECRET_KEY');
  if (!secret) throw new Error('Flutterwave not configured (FLW_SECRET_KEY missing).');
  const res = UrlFetchApp.fetch('https://api.flutterwave.com/v3/transfers', {
    method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + secret }, muteHttpExceptions: true,
    payload: JSON.stringify({
      account_bank: affiliate.bankCode, account_number: affiliate.accountNumber,
      amount: Number(amount), currency: 'NGN', narration: 'AtlasTradeAI affiliate payout',
      reference: 'atlas_payout_' + newId_()
    })
  });
  const data = JSON.parse(res.getContentText());
  if (data.status !== 'success') throw new Error('Transfer failed: ' + (data.message || 'unknown error'));

  const updated = adminSettlePayout(telegramId, affiliateId, amount);
  logAudit_(telegramId, 'settle_payout_automated', 'affiliate', affiliateId, String(amount));
  return updated;
}

// ---------- WAITLIST ----------
function adminGetWaitlist(telegramId, botId) {
  requireAdmin_(telegramId);
  return sheetToObjects_(SHEET_NAMES.WAITLIST).filter(function (w) { return String(w.botId) === String(botId); });
}

// ---------- FRAUD FLAGS ----------
function adminGetFlags(telegramId) {
  requireAdmin_(telegramId);
  return sheetToObjects_(SHEET_NAMES.FLAGS)
    .sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); });
}

function adminMarkFlagReviewed(telegramId, flagId) {
  requireAdmin_(telegramId);
  return updateRowById_(SHEET_NAMES.FLAGS, flagId, { reviewed: true });
}

// ---------- TELEGRAM NOTIFICATIONS ----------
function sendTelegramMessage_(chatId, text) {
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) return;
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: chatId, text: text })
    });
  } catch (err) {
    // notification failures shouldn't break core flows
  }
}

function notifyUser_(telegramId, text) {
  if (String(telegramId).indexOf('demo_') === 0) return;
  sendTelegramMessage_(telegramId, text);
}

function notifyBotSubscribersNewTrade_(trade) {
  const bot = getBot(trade.botId);
  if (!bot) return;
  const subscribers = {};
  sheetToObjects_(SHEET_NAMES.DEPOSITS).forEach(function (d) {
    if (String(d.botId) === String(trade.botId) && d.status === 'confirmed') subscribers[d.telegramId] = true;
  });
  const profit = Number(trade.profitPct || 0);
  const text = bot.name + ': new ' + trade.side + ' trade on ' + trade.symbol + ' (' + (profit >= 0 ? '+' : '') + profit + '%)';
  Object.keys(subscribers).forEach(function (tid) { notifyUser_(tid, text); });
}

// ---------- DASHBOARD ----------
function getUserDashboard(telegramId) {
  const deposits = getUserDeposits(telegramId).filter(function (d) { return d.status === 'confirmed'; });
  const withdrawals = sheetToObjects_(SHEET_NAMES.WITHDRAWALS).filter(function (w) { return String(w.telegramId) === String(telegramId) && w.status === 'paid'; });

  const byBot = {};
  deposits.forEach(function (d) {
    if (!byBot[d.botId]) byBot[d.botId] = { botId: d.botId, deposited: 0, withdrawn: 0, firstDepositAt: d.createdAt };
    byBot[d.botId].deposited += Number(d.amount);
    if (new Date(d.createdAt) < new Date(byBot[d.botId].firstDepositAt)) byBot[d.botId].firstDepositAt = d.createdAt;
  });
  withdrawals.forEach(function (w) { if (byBot[w.botId]) byBot[w.botId].withdrawn += Number(w.amount); });

  const positions = Object.keys(byBot).map(function (botId) {
    const p = byBot[botId];
    const bot = getBot(botId);
    const trades = getTrades(botId);
    const profitPct = trades.reduce(function (sum, t) { return sum + Number(t.profitPct || 0); }, 0);
    const daysSinceDeposit = Math.floor((new Date() - new Date(p.firstDepositAt)) / 86400000);
    return {
      botId: botId, botName: bot ? bot.name : 'Unknown', deposited: p.deposited, withdrawn: p.withdrawn,
      balance: p.deposited - p.withdrawn, estProfitPct: profitPct, daysSinceDeposit: daysSinceDeposit,
      withdrawalUnlocked: bot ? daysSinceDeposit >= Number(bot.withdrawalLockDays || 0) : true
    };
  });

  return { positions: positions };
}

function getUserStatementCsv(telegramId) {
  const deposits = getUserDeposits(telegramId);
  const withdrawals = sheetToObjects_(SHEET_NAMES.WITHDRAWALS).filter(function (w) { return String(w.telegramId) === String(telegramId); });
  let csv = 'Type,Bot,Amount,Status,Date\n';
  deposits.forEach(function (d) { csv += 'Deposit,' + d.botId + ',' + d.amount + ',' + d.status + ',' + d.createdAt + '\n'; });
  withdrawals.forEach(function (w) { csv += 'Withdrawal,' + w.botId + ',' + w.amount + ',' + w.status + ',' + w.createdAt + '\n'; });
  return csv;
}

// ---------- ADMIN STATS ----------
function adminGetStats(telegramId) {
  requireAdmin_(telegramId);
  const deposits = sheetToObjects_(SHEET_NAMES.DEPOSITS);
  const withdrawals = sheetToObjects_(SHEET_NAMES.WITHDRAWALS);
  const confirmedDeposits = deposits.filter(function (d) { return d.status === 'confirmed'; }).reduce(function (s, d) { return s + Number(d.amount); }, 0);
  const paidWithdrawals = withdrawals.filter(function (w) { return w.status === 'paid'; }).reduce(function (s, w) { return s + Number(w.amount); }, 0);
  const affiliates = sheetToObjects_(SHEET_NAMES.AFFILIATES);
  return {
    totalAUM: confirmedDeposits - paidWithdrawals,
    activeUsers: sheetToObjects_(SHEET_NAMES.USERS).length,
    pendingDeposits: deposits.filter(function (d) { return d.status === 'pending'; }).length,
    pendingWithdrawals: withdrawals.filter(function (w) { return w.status === 'pending'; }).length,
    pendingPayouts: affiliates.filter(function (a) { return Number(a.pendingEarnings || 0) >= getSettings_().payoutThreshold; }).length,
    openFlags: sheetToObjects_(SHEET_NAMES.FLAGS).filter(function (f) { return f.reviewed !== true && f.reviewed !== 'TRUE'; }).length
  };
}
