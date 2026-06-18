const FEFixed = (() => {
  const SRC_SHEET = Config.SHEETS.POSTS;
  const DST_SHEET = Config.SHEETS.FRONTEND_FIXED;
  const DST_HDR   = Config.HEADERS.FRONTEND_FIXED;

  function ensureFE_(ss) {
    if (Utils && Utils.Sheets && typeof Utils.Sheets.ensureSheet === 'function') {
      return Utils.Sheets.ensureSheet(ss, DST_SHEET, DST_HDR);
    }
    const sh = ss.getSheetByName(DST_SHEET) || ss.insertSheet(DST_SHEET);
    Config.ensureHeaderIfEmpty(sh, DST_HDR);
    return sh;
  }

  /* ==================== 1) Append new rows from Posts ==================== */
  function syncFromPosts() {
    const ss  = SpreadsheetApp.getActive();
    const src = ss.getSheetByName(SRC_SHEET);
    if (!src) throw new Error('Missing sheet: ' + SRC_SHEET);

    const dst = ensureFE_(ss);

    // source header + index
    const srcHdr = safeHeader_(src);
    const S = indexFor_(srcHdr);

    // read all source rows
    const srcLastRow = src.getLastRow();
    const srcLastCol = src.getLastColumn();
    const srcRows = (srcLastRow > 1 && srcLastCol > 0)
      ? src.getRange(2, 1, srcLastRow - 1, srcLastCol).getValues()
      : [];

    // dest existing ids
    const dstHdr = safeHeader_(dst);
    const D = indexFor_(dstHdr);
    const idColDst = D['id'];
    const existing = new Set();
    const dstLastRow = dst.getLastRow();
    if (dstLastRow > 1 && idColDst != null) {
      const vals = dst.getRange(2, 1, dstLastRow - 1, dst.getLastColumn()).getValues();
      for (let i = 0; i < vals.length; i++) {
        const v = (vals[i][idColDst] || '').toString().trim();
        if (v) existing.add(v);
      }
    }

    // helper
    const get = (row, key) => {
      const i = S[(key || '').toLowerCase()];
      return (i == null ? '' : row[i]);
    };

    const out = [];
    const idxDst = indexFor_(DST_HDR);
    for (const r of srcRows) {
      const id = (get(r, 'id') || '').toString().trim();
      if (!id || existing.has(id)) continue;

      const createdUtc = Number(get(r, 'created_utc')) || 0;
      const createdObj = get(r, 'created');
      const created_at = createdObj instanceof Date
        ? createdObj
        : (createdUtc > 0 ? new Date(createdUtc * 1000) : '');

      const row = new Array(DST_HDR.length).fill('');
      if (idxDst.id           != null) row[idxDst.id]           = id;
      if (idxDst.author       != null) row[idxDst.author]       = get(r, 'author') || '';
      if (idxDst.title        != null) row[idxDst.title]        = get(r, 'title') || '';
      if (idxDst.post_content != null) row[idxDst.post_content] = get(r, 'selftext') || '';
      if (idxDst.created_at   != null) row[idxDst.created_at]   = created_at;

      out.push(row);
      existing.add(id);
    }

    if (out.length) {
      const start = dst.getLastRow() + 1;
      dst.getRange(start, 1, out.length, DST_HDR.length).setValues(out);
    }

    // format created_at + sort newest
    const H = indexFor_(safeHeader_(dst));
    if (H['created_at'] != null) {
      const col = H['created_at'] + 1;
      const rows = Math.max(0, dst.getLastRow() - 1);
      if (rows > 0) dst.getRange(2, col, rows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
      if (Utils && Utils.Sheets && typeof Utils.Sheets.sortByHeader === 'function') {
        Utils.Sheets.sortByHeader(dst, 'created_at', { ascending: false });
      }
    }
    fixRowFormat(dst, 21);
    Utils.Log.append('info', `FE Fixed: appended ${out.length} new row(s)`, JSON.stringify({added: out.length}));
  }

  /* ==================== 2) Fill E:direction (uses utils_direction_classifier) ==================== */
  function fillDirections(options) {
    options = options || {};
    const force = !!options.force;

    const ss = SpreadsheetApp.getActive();
    const sh = ensureFE_(ss);

    const hdr = Utils.Sheets.header(sh);
    const H   = Utils.Sheets.index(hdr);
    const colTitle = (H['title'] != null ? H['title'] + 1 : 3);
    const colBody  = (H['post_content'] != null ? H['post_content'] + 1 : 4);
    const colDir   = (H['direction'] != null ? H['direction'] + 1 : 5);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const width = Math.max(colTitle, colBody, colDir);
    const vals = sh.getRange(2, 1, lastRow - 1, width).getValues();

    const out = new Array(vals.length);
    for (let i = 0; i < vals.length; i++) {
      const existing = (vals[i][colDir - 1] || '').toString().trim();
      if (existing && !force) { out[i] = [existing]; continue; }

      const title = (vals[i][colTitle - 1] || '').toString();
      const body  = (vals[i][colBody  - 1] || '').toString();

      let label = '';
      try {
        if (typeof Direction !== 'undefined' && typeof Direction.classify === 'function') {
          label = Direction.classify(title, body, {});
        } else if (typeof detectDirection === 'function') {
          label = detectDirection(title, body, {});
        } else if (typeof classifyDirection === 'function') {
          label = classifyDirection(title, body, {});
        } else if (typeof DirectionClassifier !== 'undefined' && typeof DirectionClassifier.classify === 'function') {
          label = DirectionClassifier.classify(title, body, {});
        } else if (typeof UtilsDirection !== 'undefined' && typeof UtilsDirection.classify === 'function') {
          label = UtilsDirection.classify(title, body, {});
        }
      } catch (_) {}

      const norm = normalizeDirection_(label);
      out[i] = [norm];
    }

    sh.getRange(2, colDir, out.length, 1).setValues(out);
    fixRowFormat(sh, 21);
    Utils.Log.append('info', 'FE Fixed: FE Fixed: directions updated', '');
  }

  function normalizeDirection_(s) {
    const k = (s || '').toString().trim().toLowerCase();
    if (!k) return '';
    if (/bull/.test(k)) return 'bullish';
    if (/bear/.test(k)) return 'bearish';
    if (/neut|flat|side/.test(k)) return 'neutral';
    return k;
  }

  /* ==================== 3) Fill F:ticker (uses utils_ticker_extractor) ==================== */
  function fillTickers(options) {
    options = options || {};
    const force = !!options.force;

    const ss  = SpreadsheetApp.getActive();
    const sh  = ensureFE_(ss);

    const hdr = Utils.Sheets.header(sh);
    const H   = Utils.Sheets.index(hdr);

    const colTitle   = H['title']        != null ? H['title']        + 1 : 3; // C
    const colContent = H['post_content'] != null ? H['post_content'] + 1 : 4; // D
    const colTicker  = H['ticker']       != null ? H['ticker']       + 1 : 6; // F

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    // Build Symbols index once (required by some extractors)
    const symSheet = ss.getSheetByName(Config.SHEETS.SYMBOLS);
    const symHdr   = symSheet ? Utils.Sheets.header(symSheet) : [];
    const symRows  = (symSheet && symSheet.getLastRow() > 1 && symSheet.getLastColumn() > 0)
      ? symSheet.getRange(2, 1, symSheet.getLastRow() - 1, symSheet.getLastColumn()).getValues()
      : [];
    const symTable = { header: symHdr, rows: symRows };
    const symIdx   = (typeof SymbolsIndex !== 'undefined' && typeof SymbolsIndex.build === 'function')
      ? SymbolsIndex.build(symTable)
      : null;

    const extractor = resolveTickerExtractor_(symIdx);

    const neededCols = Math.max(colTitle, colContent, colTicker);
    const values = sh.getRange(2, 1, lastRow - 1, neededCols).getValues();

    const out = new Array(values.length);
    for (let i = 0; i < values.length; i++) {
      const existing = (values[i][colTicker - 1] || '').toString().trim();
      if (existing && !force) { out[i] = [existing]; continue; }

      const title = (values[i][colTitle - 1] || '').toString();
      const body  = (values[i][colContent - 1]  || '').toString();

      const res = extractor(title, body);
      const picked = pickTickerFromResult_(res);
      out[i] = [normalizeTicker_(picked)];
    }

    sh.getRange(2, colTicker, out.length, 1).setValues(out);
    fixRowFormat(sh, 21);
    Utils.Log.append('info', 'FE Fixed: FE Fixed: tickers updated', '');
  }

  function resolveTickerExtractor_(symIdx) {
    if (typeof extractTicker === 'function') {
      return (t, b) => (extractTicker.length >= 2 ? extractTicker(t, b) : extractTicker((t||'')+'\n\n'+(b||'')));
    }
    if (typeof extractTickers === 'function') {
      return (t, b) => (extractTickers.length >= 2 ? extractTickers(t, b) : extractTickers((t||'')+'\n\n'+(b||'')));
    }
    if (typeof TickerExtractor !== 'undefined') {
      if (typeof TickerExtractor.extract === 'function') return (t,b) => TickerExtractor.extract(t,b);
      if (typeof TickerExtractor.extractOne === 'function') return (t,b) => TickerExtractor.extractOne(t,b);
      if (typeof TickerExtractor.extractAll === 'function') return (t,b) => TickerExtractor.extractAll(t,b);
    }
    if (typeof UtilsTicker !== 'undefined' && typeof UtilsTicker.extract === 'function') {
      return (t,b) => UtilsTicker.extract(t,b);
    }
    if (typeof Ticker !== 'undefined') {
      if (typeof Ticker.extract === 'function') {
        return (t,b) => {
          let x = Ticker.extract(t, symIdx);
          if (!x && typeof Ticker.extractFallback === 'function') x = Ticker.extractFallback(t, b, symIdx);
          return x || '';
        };
      }
    }
    throw new Error('Ticker extractor not found. Expected one of: extractTicker / extractTickers / TickerExtractor.* / UtilsTicker.extract / Ticker.extract');
  }

  function pickTickerFromResult_(res) {
    if (res == null) return '';
    if (Array.isArray(res))       return res[0] || '';
    if (typeof res === 'string')  return res;
    if (res instanceof Set)       { var it = res.values(); var n = it.next(); return n && !n.done ? n.value : ''; }
    if (typeof res === 'object')  return res.ticker || res.symbol || (Array.isArray(res.tickers) ? res.tickers[0] : '') || '';
    return '';
  }

  function normalizeTicker_(t) {
    if (!t) return '';
    var s = String(t).trim();
    s = s.replace(/^\$/, '');
    s = s.toUpperCase().replace(/[^A-Z.]/g, '');
    return s;
  }

  /* ==================== 4) Fill H:price_at_post (previous trading day) ==================== */
  function fillPricesAtPost(options) {
    options = options || {};
    const force = !!options.force;

    const ss = SpreadsheetApp.getActive();
    const fe = ensureFE_(ss);

    const hdr = Utils.Sheets.header(fe);
    const H   = Utils.Sheets.index(hdr);

    const colTicker = H['ticker']         != null ? H['ticker']         + 1 : 6; // F
    const colCreated= H['created_at']     != null ? H['created_at']     + 1 : 7; // G
    const colPrice  = H['price_at_post']  != null ? H['price_at_post']  + 1 : 8; // H

    const lastRow = fe.getLastRow();
    if (lastRow < 2) return;

    // Build cache map: "TICKER|YYYY-MM-DD" -> close
    const cache = ss.getSheetByName(Config.SHEETS.TICKER_CACHE);
    if (!cache) throw new Error('Missing sheet: ' + Config.SHEETS.TICKER_CACHE);
    const cHdr = Utils.Sheets.header(cache);
    const C    = Utils.Sheets.index(cHdr);
    const cLastR = cache.getLastRow(), cLastC = cache.getLastColumn();
    const priceMap = new Map();
    if (cLastR > 1 && cLastC > 0) {
      const cVals = cache.getRange(2, 1, cLastR - 1, cLastC).getValues();
      for (let i = 0; i < cVals.length; i++) {
        const t = (cVals[i][C['ticker']] || '').toString().trim().toUpperCase();
        const d = cVals[i][C['date']];
        const ymd = (d instanceof Date) ? ymdUTC_(d) : (d ? String(d).slice(0,10) : '');
        const close = cVals[i][C['close']];
        if (t && ymd) priceMap.set(t + '|' + ymd, close);
      }
    }

    // Read needed FE cols
    const width = Math.max(colTicker, colCreated, colPrice);
    const vals  = fe.getRange(2, 1, lastRow - 1, width).getValues();
    const out   = new Array(vals.length);

    for (let i = 0; i < vals.length; i++) {
      const existing = (vals[i][colPrice - 1] || '').toString().trim();
      if (existing && !force) { out[i] = [existing]; continue; }

      const ticker  = (vals[i][colTicker - 1] || '').toString().trim().toUpperCase();
      const created = vals[i][colCreated - 1];
      if (!ticker || !(created instanceof Date)) { out[i] = ['']; continue; }

      const key = prevTradingKey_(ticker, created, priceMap, 10); // look back up to 10 calendar days
      const price = key ? priceMap.get(key) : '';
      out[i] = [price];
    }

    // Write + format
    fe.getRange(2, colPrice, out.length, 1).setValues(out);
    fe.getRange(2, colPrice, Math.max(0, lastRow - 1), 1).setNumberFormat('0.00');
    fixRowFormat(fe, 21);
    Utils.Log.append('info', 'FE Fixed: FE Fixed: price_at_post updated', '');
  }

  // Return "TICKER|YYYY-MM-DD" for the most recent cached trading day strictly BEFORE 'created'
  function prevTradingKey_(ticker, created, priceMap, maxLookback) {
    maxLookback = (typeof maxLookback === 'number' && maxLookback > 0) ? maxLookback : 10;
    let probe = new Date(Date.UTC(
      created.getUTCFullYear(),
      created.getUTCMonth(),
      created.getUTCDate() - 1
    ));
    for (let i = 0; i < maxLookback; i++) {
      const key = ticker + '|' + ymdUTC_(probe);
      if (priceMap.has(key)) return key;
      probe = new Date(Date.UTC(
        probe.getUTCFullYear(),
        probe.getUTCMonth(),
        probe.getUTCDate() - 1
      ));
    }
    return null;
  }

  // helper: UTC YYYY-MM-DD
  function ymdUTC_(d) {
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
  }

  /* ---------- small local utils ---------- */
  function indexFor_(hdr) {
    const m = {};
    for (let i = 0; i < hdr.length; i++) {
      const k = (hdr[i] || '').toString().trim().toLowerCase();
      if (k) m[k] = i;
    }
    return m;
  }

  function safeHeader_(sh) {
    if (!sh) return [];
    const lc = sh.getLastColumn();
    if (!lc) return [];
    const vals = sh.getRange(1, 1, 1, lc).getValues()[0] || [];
    return vals.map(x => (x || '').toString().trim());
  }

  return { syncFromPosts, fillDirections, fillTickers, fillPricesAtPost };
})();

/* -------- wrappers (for menu/trigger) -------- */
function FEFixed_syncFromPosts()     { return FEFixed.syncFromPosts(); }
function FEFixed_fillDirections()     { return FEFixed.fillDirections({ force:false }); }
function FEFixed_fillTickers()        { return FEFixed.fillTickers({ force:false }); }
function FEFixed_fillPricesAtPost()   { return FEFixed.fillPricesAtPost({ force:false }); }
