/**
 * ATLASTRADEAI — UI.gs
 * A one-time content-seeding tool: adds a custom menu to the Google Sheet ("AtlasTradeAI")
 * so you can populate 8 ready-made bots without typing them into the Admin mini app one by
 * one. Each bot's name/description is modeled on a well-known trader's real, publicly
 * documented strategy or methodology — the same way "Fibonacci retracement" or "Elliott
 * Wave" are industry-standard terms. This is a style/strategy homage, not a claim that any
 * of these people created, endorse, or are affiliated with AtlasTradeAI or these bots.
 *
 * HOW TO RUN:
 *  - Open the Google Sheet (Admin → Settings shows its URL, or run setupSheets() once to
 *    get it), then reload the page. A new "AtlasTradeAI" menu appears in the menu bar.
 *  - Click AtlasTradeAI → "Seed 8 legendary bots (skip duplicates)". That's it — no
 *    telegramId/admin auth needed here, since this only runs from inside the Sheet/editor
 *    itself, the same trust level as setupSheets().
 *  - Re-running it is safe: the "skip duplicates" option won't create a second copy of a
 *    bot whose name already exists in the Bots sheet.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AtlasTradeAI')
    .addItem('Seed 8 legendary bots (skip duplicates)', 'seedLegendaryBotsSafely')
    .addItem('Seed 8 legendary bots (force — may duplicate)', 'seedLegendaryBots')
    .addToUi();
}

// Currencies list matches the "full world list, USD-anchored" decision — every bot accepts
// the same broad set; whether a user actually gets Flutterwave checkout or crypto-only
// depends on their country, handled entirely on the frontend/Settings side, not per-bot.
const LEGENDARY_CURRENCIES = 'USD,NGN,GHS,KES,ZAR,UGX,TZS,RWF,ZMW,XOF,XAF';

const LEGENDARY_BOTS = [
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
  }
];

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

// Adds all 8, even if some/all already exist (will create duplicate rows if run twice).
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
