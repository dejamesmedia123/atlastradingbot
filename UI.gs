/**
 * ATLASTRADEAI — UI.gs
 * A one-time content-seeding tool: adds a custom menu to the Google Sheet ("AtlasTradeAI")
 * so you can populate ready-made bots without typing them into the Admin mini app one by
 * one. Most bots' name/description are modeled on a renowned forex/macro trader's real,
 * publicly documented strategy or methodology — the same way "Fibonacci retracement" or
 * "Elliott Wave" are industry-standard terms. This is a style/strategy homage, not a claim
 * that any of these people created, endorse, or are affiliated with AtlasTradeAI or these
 * bots. A small extra group at the end is fully original — AtlasTradeAI's own strategy
 * concepts with no trader name attached.
 *
 * HOW TO RUN:
 *  - Open the Google Sheet (Admin → Settings shows its URL, or run setupSheets() once to
 *    get it), then reload the page. A new "AtlasTradeAI" menu appears in the menu bar.
 *  - Click AtlasTradeAI → "Seed legendary bots (skip duplicates)". That's it — no
 *    telegramId/admin auth needed here, since this only runs from inside the Sheet/editor
 *    itself, the same trust level as setupSheets().
 *  - Re-running it is safe: the "skip duplicates" option won't create a second copy of a
 *    bot whose name already exists in the Bots sheet.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AtlasTradeAI')
    .addItem('Seed legendary bots (skip duplicates)', 'seedLegendaryBotsSafely')
    .addItem('Seed legendary bots (force — may duplicate)', 'seedLegendaryBots')
    .addItem('Seed random trade history for every bot', 'seedRandomTradesForAllBots')
    .addToUi();
}

// Currencies list matches the "full world list, USD-anchored" decision — every bot accepts
// the same broad set; whether a user actually gets Flutterwave checkout or crypto-only
// depends on their country, handled entirely on the frontend/Settings side, not per-bot.
const LEGENDARY_CURRENCIES = 'USD,NGN,GHS,KES,ZAR,UGX,TZS,RWF,ZMW,XOF,XAF';

// ---------- GROUP 1: modeled on renowned traders' publicly documented strategies ----------
const TRADER_HOMAGE_BOTS = [
  {
    name: 'Wyckoff AMD',
    description: "Modeled on Richard Wyckoff's century-old Accumulation/Distribution framework, updated with the modern AMD (Accumulation–Manipulation–Distribution) reading of intraday structure. Looks for a quiet accumulation range, a manipulation sweep of liquidity, then rides the resulting distribution move on major FX pairs.",
    tvSymbol: 'OANDA:EURUSD', riskLevel: 'Medium', assetClass: 'Forex',
    targetMarkets: 'EUR/USD, GBP/USD, USD/JPY', avgTradeDuration: '1-3 days',
    minDeposit: 100, maxDeposit: 10000, minTopUp: 50, withdrawalLockDays: 14, minWithdrawal: 20,
    profitSharePct: 20, managementFeePct: 1.5, payoutFrequency: 'Weekly', maxDrawdownPct: 10
  },
  {
    name: 'The Alchemist',
    description: "A macro, reflexivity-driven bot in the spirit of George Soros's 'Alchemy of Finance' — it looks past pure chart patterns to interest-rate differentials, risk sentiment, and self-reinforcing trends, then takes concentrated, high-conviction positions in majors and gold when a macro narrative is clearly driving price.",
    tvSymbol: 'OANDA:XAUUSD', riskLevel: 'High', assetClass: 'Macro/Forex',
    targetMarkets: 'XAU/USD, EUR/USD, USD/JPY', avgTradeDuration: '1-4 weeks',
    minDeposit: 200, maxDeposit: 20000, minTopUp: 100, withdrawalLockDays: 30, minWithdrawal: 30,
    profitSharePct: 25, managementFeePct: 2, payoutFrequency: 'Monthly', maxDrawdownPct: 18
  },
  {
    name: 'Livermore Momentum',
    description: "Named for Jesse Livermore's 'trade with the trend, not against it' philosophy — waits for a confirmed breakout on strong volume, then pyramids into the position as the move extends, and cuts losers fast the way Livermore's own rules demanded.",
    tvSymbol: 'OANDA:SPX500USD', riskLevel: 'High', assetClass: 'Indices',
    targetMarkets: 'S&P 500, Nasdaq 100', avgTradeDuration: '3-10 days',
    minDeposit: 150, maxDeposit: 15000, minTopUp: 75, withdrawalLockDays: 21, minWithdrawal: 25,
    profitSharePct: 22, managementFeePct: 2, payoutFrequency: 'Monthly', maxDrawdownPct: 15
  },
  {
    name: 'Turtle Trend',
    description: "Built on the original Turtle Traders' rules (Richard Dennis & William Eckhardt) — a purely systematic breakout-and-trend-following system with fixed position sizing and hard stop rules, spread across a diversified basket of currency pairs to avoid over-reliance on any single trend.",
    tvSymbol: 'OANDA:GBPUSD', riskLevel: 'Medium', assetClass: 'Forex',
    targetMarkets: 'GBP/USD, AUD/USD, USD/CAD, NZD/USD', avgTradeDuration: '1-2 weeks',
    minDeposit: 100, maxDeposit: 10000, minTopUp: 50, withdrawalLockDays: 14, minWithdrawal: 20,
    profitSharePct: 18, managementFeePct: 1.5, payoutFrequency: 'Weekly', maxDrawdownPct: 12
  },
  {
    name: 'Gann Time & Price',
    description: "Inspired by W.D. Gann's geometric approach to markets — combining time cycles with price angles to anticipate where gold is likely to turn, rather than only reacting after the fact.",
    tvSymbol: 'OANDA:XAUUSD', riskLevel: 'Medium', assetClass: 'Commodities',
    targetMarkets: 'XAU/USD', avgTradeDuration: '1-3 weeks',
    minDeposit: 200, maxDeposit: 15000, minTopUp: 100, withdrawalLockDays: 14, minWithdrawal: 30,
    profitSharePct: 20, managementFeePct: 1.5, payoutFrequency: 'Monthly', maxDrawdownPct: 12
  },
  {
    name: 'Bollinger Squeeze',
    description: "Uses John Bollinger's volatility bands the way they were designed — waiting for a tight 'squeeze' (low volatility) and trading the breakout that typically follows, with the bands themselves used again to manage the exit as volatility expands.",
    tvSymbol: 'OANDA:EURUSD', riskLevel: 'Low', assetClass: 'Forex',
    targetMarkets: 'EUR/USD, EUR/GBP', avgTradeDuration: '2-5 days',
    minDeposit: 50, maxDeposit: 5000, minTopUp: 25, withdrawalLockDays: 7, minWithdrawal: 10,
    profitSharePct: 15, managementFeePct: 1, payoutFrequency: 'Weekly', maxDrawdownPct: 6
  },
  {
    name: 'Seykota Systematic',
    description: "In the disciplined, rules-first mould of Ed Seykota's trend-following ('the trend is your friend until it ends') — a fully systematic model with strict position sizing and no discretionary overrides, applied here to the major crypto pairs.",
    tvSymbol: 'BINANCE:ETHUSDT', riskLevel: 'High', assetClass: 'Crypto',
    targetMarkets: 'BTC/USDT, ETH/USDT', avgTradeDuration: '1-3 weeks',
    minDeposit: 100, maxDeposit: 20000, minTopUp: 50, withdrawalLockDays: 30, minWithdrawal: 25,
    profitSharePct: 25, managementFeePct: 2, payoutFrequency: 'Monthly', maxDrawdownPct: 22
  },
  {
    name: 'Elliott Wave Navigator',
    description: "Applies Ralph Nelson Elliott's wave-count theory — mapping the five-wave impulse and three-wave correction pattern markets tend to repeat — to swing-trade major equity indices at each new wave's turning point.",
    tvSymbol: 'OANDA:NAS100USD', riskLevel: 'Medium', assetClass: 'Indices',
    targetMarkets: 'Nasdaq 100, Dow Jones', avgTradeDuration: '1-2 weeks',
    minDeposit: 150, maxDeposit: 15000, minTopUp: 75, withdrawalLockDays: 21, minWithdrawal: 25,
    profitSharePct: 20, managementFeePct: 1.5, payoutFrequency: 'Monthly', maxDrawdownPct: 14
  },
  {
    name: 'Druckenmiller Conviction',
    description: "Modeled on Stanley Druckenmiller's concentrated, top-down macro style — 'put all your eggs in one basket and watch the basket very carefully.' Waits for a small number of high-conviction setups a year across currencies and indices rather than trading constantly, sizing up hard when the risk/reward is clearly asymmetric.",
    tvSymbol: 'OANDA:US30USD', riskLevel: 'High', assetClass: 'Macro/Indices',
    targetMarkets: 'Dow Jones, EUR/USD, XAU/USD', avgTradeDuration: '2-6 weeks',
    minDeposit: 250, maxDeposit: 25000, minTopUp: 100, withdrawalLockDays: 30, minWithdrawal: 40,
    profitSharePct: 25, managementFeePct: 2, payoutFrequency: 'Monthly', maxDrawdownPct: 20
  },
  {
    name: 'Dalio All Weather',
    description: "Inspired by Ray Dalio's risk-parity 'All Weather' philosophy — spreads exposure across asset classes that tend to zig when others zag (metals, indices, majors) so the portfolio is built to hold up across inflation, growth, and risk-off regimes rather than betting on one macro outcome.",
    tvSymbol: 'OANDA:XAUUSD', riskLevel: 'Low', assetClass: 'Multi-Asset',
    targetMarkets: 'XAU/USD, S&P 500, EUR/USD, US 10Y proxy', avgTradeDuration: '2-8 weeks',
    minDeposit: 100, maxDeposit: 15000, minTopUp: 50, withdrawalLockDays: 21, minWithdrawal: 20,
    profitSharePct: 15, managementFeePct: 1.5, payoutFrequency: 'Monthly', maxDrawdownPct: 8
  },
  {
    name: 'The Sultan\'s Ledger',
    description: "Named for Bill Lipschutz, the so-called 'Sultan of Currencies' who ran Salomon Brothers' FX desk — the hallmark of his style wasn't the size of his positions but the discipline behind them: strict predefined risk per trade, scaling into a thesis gradually, and treating a stop-loss as non-negotiable. This bot applies that same size-with-discipline approach to the majors.",
    tvSymbol: 'OANDA:EURUSD', riskLevel: 'High', assetClass: 'Forex',
    targetMarkets: 'EUR/USD, USD/JPY, GBP/USD', avgTradeDuration: '3-10 days',
    minDeposit: 200, maxDeposit: 20000, minTopUp: 100, withdrawalLockDays: 21, minWithdrawal: 30,
    profitSharePct: 22, managementFeePct: 2, payoutFrequency: 'Monthly', maxDrawdownPct: 16
  },
  {
    name: 'Tudor Risk Guard',
    description: "In the spirit of Paul Tudor Jones's 'defense first' rule — 'the most important thing is protecting what you have' — this bot leads with hard risk limits and only takes a trade when the reward clearly outweighs the risk, and is quick to flatten a position the moment the setup that justified it breaks down.",
    tvSymbol: 'OANDA:SPX500USD', riskLevel: 'Medium', assetClass: 'Indices',
    targetMarkets: 'S&P 500, EUR/USD', avgTradeDuration: '2-5 days',
    minDeposit: 100, maxDeposit: 12000, minTopUp: 50, withdrawalLockDays: 14, minWithdrawal: 20,
    profitSharePct: 18, managementFeePct: 1.5, payoutFrequency: 'Weekly', maxDrawdownPct: 10
  },
  {
    name: 'Kovner Macro Edge',
    description: "Modeled on Bruce Kovner of Caxton Associates — pairing fundamental macro research (rate paths, growth data) with technical timing for the actual entry, so a position is only opened once both the 'why' and the 'when' line up on the same pair.",
    tvSymbol: 'OANDA:USDCAD', riskLevel: 'Medium', assetClass: 'Macro/Forex',
    targetMarkets: 'USD/CAD, EUR/USD, XAU/USD', avgTradeDuration: '1-3 weeks',
    minDeposit: 150, maxDeposit: 18000, minTopUp: 75, withdrawalLockDays: 21, minWithdrawal: 25,
    profitSharePct: 20, managementFeePct: 1.5, payoutFrequency: 'Monthly', maxDrawdownPct: 13
  },
  {
    name: 'Williams %R Sprint',
    description: "Built around Larry Williams's short-term, indicator-driven style (the %R oscillator he popularized) — hunts for sharp, fast overbought/oversold snaps on the majors and takes quick profits rather than holding for a bigger trend, echoing the rapid-fire trading that made Williams's own real-money track record famous.",
    tvSymbol: 'OANDA:GBPUSD', riskLevel: 'High', assetClass: 'Forex',
    targetMarkets: 'GBP/USD, EUR/USD', avgTradeDuration: 'Intraday-2 days',
    minDeposit: 75, maxDeposit: 8000, minTopUp: 30, withdrawalLockDays: 10, minWithdrawal: 15,
    profitSharePct: 20, managementFeePct: 1.5, payoutFrequency: 'Weekly', maxDrawdownPct: 14
  },
  {
    name: 'Raschke Swing Grail',
    description: "Inspired by Linda Raschke's swing-trading setups (as detailed in 'Street Smarts' with Larry Connors) — looks for a market to pull back against its dominant short-term trend and enters as momentum resumes, aiming to catch the next leg of an already-established move rather than the very start of a new one.",
    tvSymbol: 'OANDA:AUDUSD', riskLevel: 'Medium', assetClass: 'Forex',
    targetMarkets: 'AUD/USD, NZD/USD, EUR/USD', avgTradeDuration: '2-6 days',
    minDeposit: 100, maxDeposit: 10000, minTopUp: 50, withdrawalLockDays: 14, minWithdrawal: 20,
    profitSharePct: 18, managementFeePct: 1.5, payoutFrequency: 'Weekly', maxDrawdownPct: 11
  },
  {
    name: 'Schwartz Pit Bull',
    description: "Named for Marty 'Pit Bull' Schwartz, the champion futures trader known for relentless tape-reading and momentum entries after years spent unlearning a purely fundamental approach — this bot reacts to what price and volume are doing right now on fast-moving pairs, and exits the instant momentum stalls rather than hoping for more.",
    tvSymbol: 'OANDA:USDJPY', riskLevel: 'High', assetClass: 'Forex',
    targetMarkets: 'USD/JPY, GBP/USD', avgTradeDuration: 'Intraday-3 days',
    minDeposit: 100, maxDeposit: 10000, minTopUp: 50, withdrawalLockDays: 14, minWithdrawal: 20,
    profitSharePct: 22, managementFeePct: 1.5, payoutFrequency: 'Weekly', maxDrawdownPct: 15
  }
];

// ---------- GROUP 2: a small set of original AtlasTradeAI concepts, no trader homage ----------
const ORIGINAL_STRATEGY_BOTS = [
  {
    name: 'Atlas Steady FX',
    description: "House-built low-volatility strategy trading the deepest, most liquid major forex pairs with tight stop-losses and small, frequent gains. Built for users who want the slowest, most predictable ride in the lineup rather than chasing big single trades.",
    tvSymbol: 'FX:EURUSD', riskLevel: 'Low', assetClass: 'Forex',
    targetMarkets: 'EUR/USD, GBP/USD', avgTradeDuration: 'Intraday-2 days',
    minDeposit: 25, maxDeposit: 3000, minTopUp: 10, withdrawalLockDays: 3, minWithdrawal: 5,
    profitSharePct: 12, managementFeePct: 0.5, payoutFrequency: 'Daily', maxDrawdownPct: 5
  },
  {
    name: 'Session Overlap Sniper',
    description: "Trades only the London/New York session overlap, when major-pair volume and volatility peak — sitting out the rest of the day entirely. A narrow, disciplined time window rather than a chart-pattern edge.",
    tvSymbol: 'OANDA:GBPUSD', riskLevel: 'Medium', assetClass: 'Forex',
    targetMarkets: 'GBP/USD, EUR/USD, USD/JPY', avgTradeDuration: 'Intraday',
    minDeposit: 75, maxDeposit: 8000, minTopUp: 30, withdrawalLockDays: 10, minWithdrawal: 15,
    profitSharePct: 18, managementFeePct: 1, payoutFrequency: 'Weekly', maxDrawdownPct: 9
  },
  {
    name: 'Central Bank Drift',
    description: "Rules-based event strategy around major central bank rate decisions and statements — positions ahead of a scheduled announcement based on the prevailing rate-differential trend, then manages the post-announcement volatility spike systematically rather than reacting emotionally to the headline.",
    tvSymbol: 'OANDA:USDJPY', riskLevel: 'High', assetClass: 'Forex',
    targetMarkets: 'USD/JPY, EUR/USD, GBP/USD', avgTradeDuration: '1-5 days',
    minDeposit: 150, maxDeposit: 12000, minTopUp: 75, withdrawalLockDays: 14, minWithdrawal: 25,
    profitSharePct: 22, managementFeePct: 2, payoutFrequency: 'Monthly', maxDrawdownPct: 16
  },
  {
    name: 'Correlation Break Arb',
    description: "Watches historically tight-correlated currency pairs for the relationship to temporarily snap out of line, then trades the reversion back to the normal spread rather than either pair's outright direction.",
    tvSymbol: 'OANDA:AUDUSD', riskLevel: 'Medium', assetClass: 'Forex',
    targetMarkets: 'AUD/USD, NZD/USD, USD/CAD', avgTradeDuration: '2-7 days',
    minDeposit: 100, maxDeposit: 10000, minTopUp: 50, withdrawalLockDays: 14, minWithdrawal: 20,
    profitSharePct: 18, managementFeePct: 1.5, payoutFrequency: 'Weekly', maxDrawdownPct: 11
  }
];

const LEGENDARY_BOTS = TRADER_HOMAGE_BOTS.concat(ORIGINAL_STRATEGY_BOTS);

function buildBotRow_(spec) {
  const now = nowIso_();
  return Object.assign({
    id: newId_(),
    logoUrl: '', status: 'Active',
    currencies: LEGENDARY_CURRENCIES,
    isPublic: true, countryRestriction: '', maxSubscribers: '', waitlistEnabled: false,
    autoPauseDrawdownPct: '', maxAllocationPct: '',
    affiliateCommissionOverridePct: '', affiliateEligible: true,
    createdAt: now, updatedAt: now
  }, spec);
}

// Adds all bots, even if some/all already exist (will create duplicate rows if run twice).
function seedLegendaryBots() {
  LEGENDARY_BOTS.forEach(function (spec) { appendRow_(SHEET_NAMES.BOTS, buildBotRow_(spec)); });
  const msg = 'Added ' + LEGENDARY_BOTS.length + ' bots to the Bots sheet.';
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) { Logger.log(msg); }
  return msg;
}

// Safe to click more than once — skips any name already present in the Bots sheet.
function seedLegendaryBotsSafely() {
  const existingNames = sheetToObjects_(SHEET_NAMES.BOTS).map(function (b) { return String(b.name || '').toLowerCase().trim(); });
  const toAdd = LEGENDARY_BOTS.filter(function (spec) { return existingNames.indexOf(spec.name.toLowerCase()) === -1; });
  toAdd.forEach(function (spec) { appendRow_(SHEET_NAMES.BOTS, buildBotRow_(spec)); });
  const skipped = LEGENDARY_BOTS.length - toAdd.length;
  const msg = 'Added ' + toAdd.length + ' new bot(s).' + (skipped ? ' Skipped ' + skipped + ' already present.' : '');
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) { Logger.log(msg); }
  return msg;
}

// ---------- RANDOM TRADE HISTORY ----------
// Gives every bot a believable-looking track record (used by the user app's "Recent
// trades" list and the dashboard's est. profit %) instead of showing an empty chart the
// first time someone opens a freshly-seeded bot. Numbers are randomized within ranges tied
// to the bot's own riskLevel, so a "Low" risk bot gets small, mostly-positive moves and a
// "High" risk bot gets a wider, noisier spread — not real trading results.
const RISK_PROFIT_RANGES = {
  Low: { min: -1.2, max: 2.5, winRate: 0.72 },
  Medium: { min: -3, max: 5, winRate: 0.62 },
  High: { min: -6, max: 9, winRate: 0.56 }
};

function randomBetween_(min, max) { return min + Math.random() * (max - min); }
function pickRandom_(arr) { return arr[Math.floor(Math.random() * arr.length)]; }

// Builds `count` trades for one bot, spread over the last `daysBack` days, oldest first.
function generateRandomTradesForBot_(bot, count, daysBack) {
  const symbols = String(bot.targetMarkets || bot.tvSymbol || 'N/A')
    .split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const symbolPool = symbols.length ? symbols : [bot.tvSymbol || 'N/A'];
  const range = RISK_PROFIT_RANGES[bot.riskLevel] || RISK_PROFIT_RANGES.Medium;

  const trades = [];
  for (let i = 0; i < count; i++) {
    const daysAgo = randomBetween_(0, daysBack);
    const openedAt = new Date(Date.now() - daysAgo * 86400000);
    const holdHours = randomBetween_(4, 96);
    const closedAt = new Date(openedAt.getTime() + holdHours * 3600000);
    const isWin = Math.random() < range.winRate;
    const profitPct = isWin
      ? randomBetween_(0.2, range.max)
      : randomBetween_(range.min, -0.1);
    const side = Math.random() < 0.5 ? 'Long' : 'Short';
    const entry = Math.round(randomBetween_(1, 2000) * 100) / 100;
    const exit = Math.round(entry * (1 + profitPct / 100) * 100) / 100;
    trades.push({
      id: newId_(), botId: bot.id, symbol: pickRandom_(symbolPool), side: side,
      entry: entry, exit: exit,
      profitPct: Math.round(profitPct * 100) / 100,
      profitUsd: Math.round(profitPct * 10) / 10, // illustrative only — real $ depends on each user's position size
      openedAt: openedAt.toISOString(), closedAt: closedAt.toISOString(), createdAt: nowIso_()
    });
  }
  // Oldest first, so the sheet reads chronologically.
  return trades.sort(function (a, b) { return new Date(a.openedAt) - new Date(b.openedAt); });
}

// Adds 8–14 random trades to every active bot that doesn't already have any trades.
// Safe to re-run: bots that already have at least one trade are left untouched, so this
// never piles up duplicate history on top of real trades an admin has since entered.
function seedRandomTradesForAllBots() {
  const bots = sheetToObjects_(SHEET_NAMES.BOTS).filter(function (b) { return b.status !== 'Deleted'; });
  const existingTrades = sheetToObjects_(SHEET_NAMES.TRADES);
  const botsWithTrades = {};
  existingTrades.forEach(function (t) { botsWithTrades[t.botId] = true; });

  let botsSeeded = 0, tradesAdded = 0;
  bots.forEach(function (bot) {
    if (botsWithTrades[bot.id]) return; // already has history — skip
    const count = Math.floor(randomBetween_(8, 15));
    const trades = generateRandomTradesForBot_(bot, count, 60);
    trades.forEach(function (t) { appendRow_(SHEET_NAMES.TRADES, t); });
    botsSeeded++;
    tradesAdded += trades.length;
  });

  const msg = botsSeeded
    ? 'Added ' + tradesAdded + ' trades across ' + botsSeeded + ' bot(s).'
    : 'Every bot already has trade history — nothing added. Delete rows from the Trades sheet first if you want to re-seed a bot.';
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) { Logger.log(msg); }
  return msg;
}
