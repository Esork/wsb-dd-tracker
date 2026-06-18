const SYNC_OPTIONS = { DRY_RUN: false, DO_SORT: true };
const INSERT_GUARDS = {
  REQUIRE_MARKET_CLOSED: true,
  MARKET_TZ: 'America/New_York',
  MARKET_CLOSE_HOUR: 16,
  MARKET_CLOSE_MIN: 0,
  REQUIRE_FULL_OHLCV: true,
};
const AV_HISTORY_LOADER_NAME = ''; // leave blank; we auto-detect and fallback if missing

const COLS = {
  ticker:        'ticker',
  last_refreshed:'last_refreshed',
  date:          'date',
  open:          'open',
  high:          'high',
  low:           'low',
  close:         'close',
  volume:        'volume'
};

function sync_ticker_cache() { _run_sync_core_(); }

/* ===================== Core ===================== */

function _run_sync_core_() {
  const feSymbols = getFeFixedSymbols_();
  if (!feSymbols.length) { console.log('[SYNC] fe_fixed empty'); return; }

  const cacheSyms = getCacheSymbols_();
  const cacheSet  = new Set(cacheSyms.map(upr));
  const newSymbols      = feSymbols.filter(s => !cacheSet.has(upr(s)));
  const existingSymbols = feSymbols.filter(s =>  cacheSet.has(upr(s)));

  if (existingSymbols.length) {
    const inserted = insertDailyRowsFromFmp_(existingSymbols);
    console.log(`[SYNC] FMP inserted rows: ${inserted}${SYNC_OPTIONS.DRY_RUN ? ' (dry-run)' : ''}`);
  } else {
    console.log('[SYNC] No existing tickers to insert via FMP.');
  }

  if (newSymbols.length && !SYNC_OPTIONS.DRY_RUN) {
    importNewTickersHistoryWithAV_(newSymbols);
  } else if (newSymbols.length) {
    console.log(`[SYNC] DRY_RUN: would call AV history for: ${newSymbols.join(', ')}`);
  }

  if (SYNC_OPTIONS.DO_SORT && !SYNC_OPTIONS.DRY_RUN) sortTickerCache_();
}

/* ===================== FMP (quotes → daily row) ===================== */

function getFmpApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('FMP_KEY');
  if (!key) throw new Error('FMP_KEY not set in Script Properties.');
  return key;
}

function fmpFetchBatchRaw_(symbols) {
  if (!symbols || !symbols.length) return [];
  const key = getFmpApiKey_();
  const out = [];
  const chunks = chunk_(symbols, 100);
  for (const group of chunks) {
    const url = 'https://financialmodelingprep.com/api/v3/quote/' +
                encodeURIComponent(group.join(',')) +
                `?apikey=${encodeURIComponent(key)}`;
    let code = 0, payload = [];
    try {
      const res = (function(){try{var __resp=UrlFetchApp.fetch(url, { muteHttpExceptions: true });return __resp;}catch(e){Utils.Log.append('error','UrlFetch fetch failed',''+e);throw e;}})();
      code = res.getResponseCode();
      payload = JSON.parse(res.getContentText() || '[]');
    } catch (e) {
      console.log(`[FMP] Request error for [${group.join(', ')}]: ${e && e.message ? e.message : e}`);
      payload = [];
    }
    out.push({ group, code, payload });
  }
  return out;
}

function fmpFetchBatchDailyOhlcv_(symbols) {
  const raw = fmpFetchBatchRaw_(symbols);
  const out = [];
  for (const { code, payload } of raw) {
    if (code !== 200 || !Array.isArray(payload)) continue;
    payload.forEach(item => {
      if (!item || !item.symbol) return;
      const open   = toNum_(item.open);
      const high   = toNum_(item.dayHigh);
      const low    = toNum_(item.dayLow);
      const close  = toNum_(item.price);
      const volume = toNum_(item.volume);
      if (INSERT_GUARDS.REQUIRE_FULL_OHLCV &&
          !(isFinite(open) && isFinite(high) && isFinite(low) && isFinite(close) && isFinite(volume))) {
        console.log(`[FMP] Skip ${item.symbol}: incomplete OHLCV`);
        return;
      }
      const isoTs = item.timestamp ? new Date(item.timestamp * 1000).toISOString() : new Date().toISOString();
      out.push({ symbol: String(item.symbol).trim(), open, high, low, close, volume, ts: isoTs });
    });
  }
  return out;
}

function insertDailyRowsFromFmp_(symbols) {
  const quotes = fmpFetchBatchDailyOhlcv_(symbols);
  if (!quotes.length) { console.log('[FMP] No quotes parsed for insert.'); return 0; }

  if (INSERT_GUARDS.REQUIRE_MARKET_CLOSED) {
    const now = new Date();
    const h = parseInt(Utilities.formatDate(now, INSERT_GUARDS.MARKET_TZ, 'HH'), 10);
    const m = parseInt(Utilities.formatDate(now, INSERT_GUARDS.MARKET_TZ, 'mm'), 10);
    const closed = (h > INSERT_GUARDS.MARKET_CLOSE_HOUR) ||
                   (h === INSERT_GUARDS.MARKET_CLOSE_HOUR && m >= INSERT_GUARDS.MARKET_CLOSE_MIN);
    if (!closed) { console.log(`[INSERT] Market not closed yet in ${INSERT_GUARDS.MARKET_TZ}; skip.`); return 0; }
  }

  const sh   = getSheet_(Config.SHEETS.TICKER_CACHE);
  const hdrs = readHeaderMap_(sh);

  const colTicker = mustCol_(hdrs, COLS.ticker);
  const colLR     = mustCol_(hdrs, COLS.last_refreshed);
  const colDate   = mustCol_(hdrs, COLS.date);
  const colOpen   = mustCol_(hdrs, COLS.open);
  const colHigh   = mustCol_(hdrs, COLS.high);
  const colLow    = mustCol_(hdrs, COLS.low);
  const colClose  = mustCol_(hdrs, COLS.close);
  const colVol    = mustCol_(hdrs, COLS.volume);

  const lastCol  = sh.getLastColumn();
  const lastRow0 = sh.getLastRow();
  const existingDatesBySym = buildExistingDatesIndex_(sh, colTicker, colDate);

  const rows = [];
  let skippedDup = 0;

  for (const q of quotes) {
    const ts = new Date(q.ts);
    const dateStr = Utilities.formatDate(ts, INSERT_GUARDS.MARKET_TZ, 'yyyy-MM-dd');
    const key = upr(q.symbol);
    const seen = existingDatesBySym.get(key);
    if (seen && seen.has(dateStr)) { skippedDup++; continue; }

    const row = new Array(lastCol).fill('');
    row[colTicker - 1] = q.symbol;
    row[colLR - 1]     = q.ts;
    row[colDate - 1]   = dateStr;
    row[colOpen - 1]   = q.open;
    row[colHigh - 1]   = q.high;
    row[colLow - 1]    = q.low;
    row[colClose - 1]  = q.close;
    row[colVol - 1]    = q.volume;
    rows.push(row);
  }

  if (skippedDup) console.log(`[INSERT] Skipped ${skippedDup} duplicate (ticker,date) rows.`);
  if (!rows.length) { console.log('[INSERT] Nothing to insert.'); return 0; }

  if (SYNC_OPTIONS.DRY_RUN) { console.log(`[INSERT] DRY_RUN: would insert ${rows.length} rows after R=${lastRow0}`); return rows.length; }

  sh.insertRowsAfter(lastRow0, rows.length);
  sh.getRange(lastRow0 + 1, 1, rows.length, lastCol).setValues(rows);
  console.log(`[INSERT] Inserted ${rows.length} rows.`);
  return rows.length;
}

/* ===================== New tickers (AV history) ===================== */
function getAvApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('ALPHA_VANTAGE_KEY');
  if (!key) throw new Error('ALPHA_VANTAGE_KEY not set in Script Properties.');
  return key;
}

function importNewTickersHistoryWithAV_(symbols) {
  const loader = resolveAvHistoryLoader_() || avFallbackHistoryLoader_;
  if (!loader) { console.log('[AV] No AV history loader available.'); return; }
  if (loader === importNewTickersHistoryWithAV_) { throw new Error('[AV] Misconfigured: loader resolved to wrapper importNewTickersHistoryWithAV_.'); }
  console.log(`[AV] Using loader: ${loader.name}`);
  for (const s of symbols) { try { loader(s); } catch (e) { console.log(`[AV] Failed ${s}: ${e && e.message ? e.message : e}`); } }
}

function resolveAvHistoryLoader_() {
  const fromProp = PropertiesService.getScriptProperties().getProperty('AV_HISTORY_FN');
  const candidates = [
    AV_HISTORY_LOADER_NAME && AV_HISTORY_LOADER_NAME.trim(),
    fromProp && fromProp.trim(),
    'avFetchHistoryForNewTicker',
    'fetchAlphaVantageHistoryForNewTicker',
    'avFetchHistory',
    'fetchAVHistory',
    'alphaVantageHistory',
    'alphaVantageImport',
  ].filter(Boolean);

  const blacklist = new Set(['importNewTickersHistoryWithAV_', 'sync_ticker_cache', '_run_sync_core_']);
  for (const name of candidates) {
    if (blacklist.has(name)) continue;
    const fn = (typeof globalThis !== 'undefined' ? globalThis[name] : this[name]);
    if (typeof fn === 'function') return fn;
  }
  return null;
}

/* ------- Built-in AV fallback (used only if your own loader isn’t found) ------- */
function fetchAvDailySeries_(symbol, adjusted, outputsize) {
  const key = getAvApiKey_();
  const fn  = adjusted ? 'TIME_SERIES_DAILY_ADJUSTED' : 'TIME_SERIES_DAILY';
  const os  = (outputsize || 'full') === 'full' ? 'full' : 'compact';
  const url = `https://www.alphavantage.co/query?function=${fn}&symbol=${encodeURIComponent(symbol)}&outputsize=${os}&apikey=${encodeURIComponent(key)}`;

  let text = '';
  try {
    const res = (function(){try{var __resp=UrlFetchApp.fetch(url, { muteHttpExceptions: true });return __resp;}catch(e){Utils.Log.append('error','UrlFetch fetch failed',''+e);throw e;}})();
    text = res.getContentText() || '';
  } catch (e) {
    return { ok:false, why:`fetch error: ${e && e.message ? e.message : e}` };
  }

  let json;
  try { json = JSON.parse(text); } catch (e) {
    return { ok:false, why:`parse error: ${e && e.message ? e.message : e}` };
  }

  if (json.Note)            return { ok:false, why:`Note: ${json.Note}` };
  if (json.Information)     return { ok:false, why:`Information: ${json.Information}` };
  if (json['Error Message'])return { ok:false, why:`Error: ${json['Error Message']}` };

  const series = json['Time Series (Daily)'];
  if (!series || typeof series !== 'object') return { ok:false, why:'missing daily series' };

  return { ok:true, series };
}

// Loads full daily history for a *new* ticker and inserts any (ticker,date) rows not present
function avFallbackHistoryLoader_(symbol) {
  symbol = String(symbol || '').trim();
  if (!symbol) return;

  // Try ADJUSTED first; if AV says premium/blocked/etc., fall back to DAILY.
  let r = fetchAvDailySeries_(symbol, true, 'full');
  if (!r.ok) {
    console.log(`[AV fallback] ${symbol}: ADJUSTED failed -> ${r.why}`);
    r = fetchAvDailySeries_(symbol, false, 'full');
    if (!r.ok) throw new Error(`[AV fallback] ${symbol}: DAILY failed -> ${r.why}`);
    console.log(`[AV fallback] ${symbol}: using DAILY (non-adjusted).`);
  } else {
    console.log(`[AV fallback] ${symbol}: using DAILY_ADJUSTED.`);
  }
  const series = r.series;

  const sh   = getSheet_(Config.SHEETS.TICKER_CACHE);
  const hdrs = readHeaderMap_(sh);
  const colTicker = mustCol_(hdrs, 'ticker');
  const colLR     = mustCol_(hdrs, 'last_refreshed');
  const colDate   = mustCol_(hdrs, 'date');
  const colOpen   = mustCol_(hdrs, 'open');
  const colHigh   = mustCol_(hdrs, 'high');
  const colLow    = mustCol_(hdrs, 'low');
  const colClose  = mustCol_(hdrs, 'close');
  const colVol    = mustCol_(hdrs, 'volume');

  const lastCol  = sh.getLastColumn();
  const lastRow0 = sh.getLastRow();

  // avoid dupes: use existing (ticker,date) index
  const existing = buildExistingDatesIndex_(sh, colTicker, colDate);
  const seen = existing.get(upr(symbol)) || new Set();

  const rows = [];
  const nowIso = new Date().toISOString();

  Object.keys(series).sort().forEach(dateStr => {
    if (seen.has(dateStr)) return;
    const rowObj = series[dateStr] || {};
    const o = toNum_(rowObj['1. open']);
    const h = toNum_(rowObj['2. high']);
    const l = toNum_(rowObj['3. low']);
    const c = toNum_(rowObj['4. close']);
    const v = toNum_(rowObj['6. volume'] ?? rowObj['5. volume']); // ADJUSTED uses 6., DAILY uses 5.
    if (!(isFinite(o) && isFinite(h) && isFinite(l) && isFinite(c) && isFinite(v))) return;

    const row = new Array(lastCol).fill('');
    row[colTicker - 1] = symbol;
    row[colLR - 1]     = nowIso;
    row[colDate - 1]   = dateStr;
    row[colOpen - 1]   = o;
    row[colHigh - 1]   = h;
    row[colLow - 1]    = l;
    row[colClose - 1]  = c;
    row[colVol - 1]    = v;
    rows.push(row);
  });

  if (!rows.length) { console.log(`[AV fallback] ${symbol}: nothing to insert`); return; }
  sh.insertRowsAfter(lastRow0, rows.length);
  sh.getRange(lastRow0 + 1, 1, rows.length, lastCol).setValues(rows);
  console.log(`[AV fallback] ${symbol}: inserted ${rows.length} history rows`);
}

/* ===================== Sorting ===================== */

function sortTickerCache_() {
  const sh = getSheet_(Config.SHEETS.TICKER_CACHE);
  const hdrs = readHeaderMap_(sh);
  const tickerCol = hdrs[COLS.ticker];
  const dateCol   = hdrs[COLS.date];
  if (!tickerCol || !dateCol) return;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1) return;
  sh.getRange(2, 1, lastRow - 1, lastCol).sort([
    { column: tickerCol, ascending: true  },
    { column: dateCol,   ascending: false }
  ]);
  console.log('[SORT] ticker_cache sorted: ticker ASC, date DESC');
}

/* ===================== Sheet helpers ===================== */
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}
function readHeaderMap_(sheet) {
  const row = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  const map = {};
  row.forEach((h, i) => { const key = String(h || '').trim(); if (key) map[key] = i + 1; });
  return map;
}
function mustCol_(hdrs, name) {
  const c = hdrs[name];
  if (!c) throw new Error(`Header missing in ticker_cache: "${name}"`);
  return c;
}
function getCacheSymbols_() {
  const cache = getSheet_(Config.SHEETS.TICKER_CACHE);
  const hdrs = readHeaderMap_(cache);
  const symCol = mustCol_(hdrs, COLS.ticker);
  const count = Math.max(0, cache.getLastRow() - 1);
  if (!count) return [];
  const vals = cache.getRange(2, symCol, count, 1).getValues();
  return vals.map(r => String(r[0] || '').trim()).filter(Boolean);
}
function getFeFixedSymbols_() {
  const fe = getSheet_(Config.SHEETS.FRONTEND_FIXED);
  const hdrs = readHeaderMap_(fe);
  const symCol = hdrs['ticker'] || hdrs['symbol'] || 1;
  const count = Math.max(0, fe.getLastRow() - 1);
  if (!count) return [];
  const vals = fe.getRange(2, symCol, count, 1).getValues();
  return vals.map(r => String(r[0] || '').trim()).filter(Boolean);
}

/* ===================== Index & utils ===================== */
function buildExistingDatesIndex_(sheet, colTicker, colDate) {
  const lastRow = sheet.getLastRow();
  const rowCount = Math.max(0, lastRow - 1);
  const width = Math.max(colTicker, colDate);
  const map = new Map();
  if (!rowCount) return map;
  const vals = sheet.getRange(2, 1, rowCount, width).getDisplayValues();
  for (let i = 0; i < rowCount; i++) {
    const sym  = String(vals[i][colTicker - 1] || '').trim().toUpperCase();
    const dStr = String(vals[i][colDate   - 1] || '').trim();
    if (!sym || !dStr) continue;
    const norm = normalizeDateStr_(dStr);
    if (!norm) continue;
    if (!map.has(sym)) map.set(sym, new Set());
    map.get(sym).add(norm);
  }
  return map;
}
function normalizeDateStr_(s) {
  let m = s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (m) { const y = m[1], mo = m[2].padStart(2,'0'), d = m[3].padStart(2,'0'); return `${y}-${mo}-${d}`; }
  m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (m) { const mo = m[1].padStart(2,'0'), d = m[2].padStart(2,'0'), y = m[3]; return `${y}-${mo}-${d}`; }
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  return null;
}
function upr(s) { return String(s || '').trim().toUpperCase(); }
function toNum_(v) { if (v === null || v === undefined || v === '') return NaN; const n = Number(v); return isFinite(n) ? n : NaN; }
function chunk_(arr, size) { const out = []; for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size)); return out; }

/* ===================== Deduper ===================== */
const DEDUPE_OPTIONS = { DRY_RUN: false, MARKET_TZ: 'America/New_York' };

function dedupe_today_only() {
  const todayEt = Utilities.formatDate(new Date(), DEDUPE_OPTIONS.MARKET_TZ, 'yyyy-MM-dd');
  _dedupe_by_date_(todayEt);
}
function dedupe_specific_date_yyyy_mm_dd() { const target = '2025-08-21'; _dedupe_by_date_(target); }
function dedupe_all_dates() { _dedupe_by_date_(null); }

function _dedupe_by_date_(yyyy_mm_dd) {
  const sh = getSheet_(Config.SHEETS.TICKER_CACHE);
  const hdrs = readHeaderMap_(sh);
  const colTicker = mustCol_(hdrs, 'ticker');
  const colLR     = mustCol_(hdrs, 'last_refreshed');
  const colDate   = mustCol_(hdrs, 'date');

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) { console.log('[DEDUP] No rows.'); return; }

  const tickers = sh.getRange(2, colTicker, lastRow - 1, 1).getDisplayValues();
  const datesD  = sh.getRange(2, colDate,   lastRow - 1, 1).getDisplayValues();
  const lrefRaw = sh.getRange(2, colLR,     lastRow - 1, 1).getValues();

  const groups = new Map();
  for (let i = 0; i < tickers.length; i++) {
    const row = i + 2;
    const t = String(tickers[i][0] || '').trim().toUpperCase();
    const dNorm = normalizeDateStr_(String(datesD[i][0] || '').trim());
    if (!t || !dNorm) continue;
    if (yyyy_mm_dd && dNorm !== yyyy_mm_dd) continue;
    const lrVal = lrefRaw[i][0];
    const lrMs = typeof lrVal === 'string' ? Date.parse(lrVal) : (lrVal instanceof Date ? lrVal.getTime() : NaN);
    const key = `${t}|${dNorm}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ row, lrMs: isFinite(lrMs) ? lrMs : -Infinity });
  }

  const toDelete = [];
  groups.forEach(arr => {
    if (arr.length <= 1) return;
    arr.sort((a,b)=> b.lrMs - a.lrMs);
    const dels = arr.slice(1).map(x => x.row);
    toDelete.push(...dels);
  });

  if (!toDelete.length) { console.log('[DEDUP] Nothing to delete.'); return; }
  toDelete.sort((a,b)=> b - a);

  if (DEDUPE_OPTIONS.DRY_RUN) { console.log(`[DEDUP] DRY RUN: would delete ${toDelete.length} rows: ${toDelete.join(',')}`); return; }

  toDelete.forEach(r => sh.deleteRow(r));
  console.log(`[DEDUP] Deleted ${toDelete.length} duplicate rows.`);
  if (SYNC_OPTIONS && SYNC_OPTIONS.DO_SORT) sortTickerCache_();
}
