const FELive = (() => {
  const SRC_POSTS = Config.SHEETS.POSTS;
  const SRC_FIXED = Config.SHEETS.FRONTEND_FIXED;
  const DST_LIVE  = Config.SHEETS.FRONTEND_LIVE;
  const SYM_SHEET = Config.SHEETS.SYMBOLS;
  const LIVE_HDR  = Config.HEADERS.FRONTEND_LIVE;
  const FIXED_HDR = Config.HEADERS.FRONTEND_FIXED;
  const SYM_HDR   = Config.HEADERS.SYMBOLS;

  // Finnhub conf (for cap lookups)
  const FH_BASE = (Config.API && Config.API.FINNHUB_BASE) ? Config.API.FINNHUB_BASE : 'https://finnhub.io/api/v1';
  const FH_KEY  = (() => { try { return Config.getApiKey(Config.API.FINNHUB_KEY); } catch(_) { return ''; } })();

  /* ---------------- Public entry ---------------- */
  function refresh() {
    const ss = SpreadsheetApp.getActive();

    const shPosts = ss.getSheetByName(SRC_POSTS);
    const shFixed = ss.getSheetByName(SRC_FIXED);
    if (!shPosts || !shFixed) throw new Error('Missing sheets: posts or fe_fixed');

    const shLive = ensureSheet_(ss, DST_LIVE, LIVE_HDR);
    const shSyms = ensureSheet_(ss, SYM_SHEET, SYM_HDR);

    // --- POSTS: map id -> {score, upvote_ratio, num_comments}
    const P_hdr = safeHeader_(shPosts);
    const P     = indexFor_(P_hdr);
    const pLR   = shPosts.getLastRow(), pLC = shPosts.getLastColumn();
    const pRows = (pLR > 1 && pLC > 0) ? shPosts.getRange(2,1,pLR-1,pLC).getValues() : [];
    const postsById = new Map();
    for (let i = 0; i < pRows.length; i++) {
      const id = String(pRows[i][P['id']] || '').trim();
      if (!id) continue;
      postsById.set(id, {
        score:        Number(pRows[i][P['score']] || 0),
        upvote_ratio: (pRows[i][P['upvote_ratio']] === '' || pRows[i][P['upvote_ratio']] == null) ? NaN : Number(pRows[i][P['upvote_ratio']]),
        num_comments: Number(pRows[i][P['num_comments']] || 0),
      });
    }

    // --- FE_FIXED: rows to project into LIVE
    const F_hdr = safeHeader_(shFixed);
    const F     = indexFor_(F_hdr);
    const fLR   = shFixed.getLastRow(), fLC = shFixed.getLastColumn();
    const fRows = (fLR > 1 && fLC > 0) ? shFixed.getRange(2,1,fLR-1,fLC).getValues() : [];

    // --- SYMBOLS: ticker -> cap_in_bi (normalize to billions) + row#
    const S_hdr = safeHeader_(shSyms);
    const S     = indexFor_(S_hdr);
    const sLR   = shSyms.getLastRow(), sLC = shSyms.getLastColumn();
    const sRows = (sLR > 1 && sLC > 0) ? shSyms.getRange(2,1,sLR-1,sLC).getValues() : [];

    const capByTicker     = new Map();
    const symRowByTicker  = new Map();
    const normalizeWrites = []; // [{row, valueBi}]

    for (let i = 0; i < sRows.length; i++) {
      const sym = String(sRows[i][S['symbol']] || '').trim().toUpperCase();
      let cap   = sRows[i][S['cap_in_bi']];
      if (!sym) continue;

      symRowByTicker.set(sym, i + 2);

      if (cap !== '' && cap != null && !isNaN(Number(cap))) {
        let capNum = Number(cap);
        if (capNum > 1000) {
          const bi = capNum / 1000;
          capByTicker.set(sym, bi);
          normalizeWrites.push({ row: i + 2, value: bi });
        } else {
          capByTicker.set(sym, capNum);
        }
      }
    }

    if (normalizeWrites.length && S['cap_in_bi'] != null) {
      // Batch write normalizations
      const values = normalizeWrites.map(x => [x.value]);
      shSyms.getRange(normalizeWrites[0].row, S['cap_in_bi'] + 1, values.length, 1).setValues(values);
    }

    // --- Build LIVE rows: compute ups & downs from score + upvote_ratio
    const out = [];
    const L = indexFor_(LIVE_HDR);
    const needCap = new Set();

    for (let i = 0; i < fRows.length; i++) {
      const id = String(fRows[i][F['id']] || '').trim();
      if (!id) continue;

      const p = postsById.get(id);
      if (!p) continue;

      const ticker = String(fRows[i][F['ticker']] || '').trim().toUpperCase();

      // compute [ups, downs] using utils_sheets function if present (fallback local)
      const pair = votesFromScoreRatio_(p.score, p.upvote_ratio);
      const ups   = pair[0];
      const downs = pair[1];

      let cap = '';
      if (ticker) {
        if (capByTicker.has(ticker)) cap = capByTicker.get(ticker);
        else needCap.add(ticker);
      }

      const row = new Array(LIVE_HDR.length).fill('');
      if (L.id            != null) row[L.id]            = id;
      if (L.ticker        != null) row[L.ticker]        = ticker;
      if (L.ups           != null) row[L.ups]           = ups;
      if (L.downs         != null) row[L.downs]         = downs;
      if (L.num_comments  != null) row[L.num_comments]  = p.num_comments;
      if (L.cap_in_bi     != null) row[L.cap_in_bi]     = cap;

      // carry through optional fields if your LIVE header has them
      if (L.direction     != null && F['direction']     != null) row[L.direction]     = fRows[i][F['direction']] || '';
      if (L.created_at    != null && F['created_at']    != null) row[L.created_at]    = fRows[i][F['created_at']] || '';
      if (L.price_at_post != null && F['price_at_post'] != null) row[L.price_at_post] = fRows[i][F['price_at_post']] || '';

      out.push(row);
    }

    // --- Fetch missing caps (Finnhub) and write to symbols in BILLIONS
    if (needCap.size && FH_KEY) {
      const fetched = fetchCapsFromFinnhub_Billions_(Array.from(needCap)); // Map<ticker -> capBi>
      const appends = [];
      fetched.forEach((capBi, tkr) => {
        if (capBi === '' || capBi == null) return;
        const r = symRowByTicker.get(tkr);
        if (r) {
          shSyms.getRange(r, S['cap_in_bi'] + 1, 1, 1).setValue(capBi);
        } else {
          const newRow = new Array(SYM_HDR.length).fill('');
          if (S['symbol']    != null) newRow[S['symbol']]    = tkr;
          if (S['cap_in_bi'] != null) newRow[S['cap_in_bi']] = capBi;
          appends.push(newRow);
          symRowByTicker.set(tkr, (shSyms.getLastRow() + appends.length));
        }
        capByTicker.set(tkr, capBi);
      });
      if (appends.length) {
        const start = shSyms.getLastRow() + 1;
        shSyms.getRange(start, 1, appends.length, SYM_HDR.length).setValues(appends);
      }

      // Patch newly acquired caps into out rows
      if (L.ticker != null && L.cap_in_bi != null) {
        for (let i = 0; i < out.length; i++) {
          const tkr = String(out[i][L.ticker] || '').trim().toUpperCase();
          if (tkr && (out[i][L.cap_in_bi] === '' || out[i][L.cap_in_bi] == null) && capByTicker.has(tkr)) {
            out[i][L.cap_in_bi] = capByTicker.get(tkr);
          }
        }
      }
    }

    // --- Sort newest first if created_at exists
    if (L.created_at != null) {
      out.sort((a, b) => {
        const da = a[L.created_at] instanceof Date ? a[L.created_at].getTime() : 0;
        const db = b[L.created_at] instanceof Date ? b[L.created_at].getTime() : 0;
        return db - da;
      });
    }

    // --- Replace content below header
    const oldRows = Math.max(0, shLive.getLastRow() - 1);
    if (oldRows > 0) shLive.getRange(2, 1, oldRows, shLive.getLastColumn()).clearContent();
    if (out.length) shLive.getRange(2, 1, out.length, LIVE_HDR.length).setValues(out);

    // --- Formats
    const liveHdrNow = safeHeader_(shLive);
    const LI = indexFor_(liveHdrNow);
    if (LI['ups']        != null) shLive.getRange(2, LI['ups']   + 1, Math.max(0, shLive.getLastRow()-1), 1).setNumberFormat('0');
    if (LI['downs']      != null) shLive.getRange(2, LI['downs'] + 1, Math.max(0, shLive.getLastRow()-1), 1).setNumberFormat('0');
    if (LI['num_comments'] != null) shLive.getRange(2, LI['num_comments'] + 1, Math.max(0, shLive.getLastRow()-1), 1).setNumberFormat('0');
    if (LI['cap_in_bi']  != null) shLive.getRange(2, LI['cap_in_bi'] + 1, Math.max(0, shLive.getLastRow()-1), 1).setNumberFormat('0.00');
    if (LI['created_at'] != null) shLive.getRange(2, LI['created_at'] + 1, Math.max(0, shLive.getLastRow()-1), 1).setNumberFormat('yyyy-mm-dd hh:mm');
    if (LI['price_at_post'] != null) shLive.getRange(2, LI['price_at_post'] + 1, Math.max(0, shLive.getLastRow()-1), 1).setNumberFormat('0.00');

    fixRowFormat(shLive, 21);
    SpreadsheetApp.getActive().toast(`FE Live: ${out.length} row(s)`, 'WSB', 3);
  }

  /* ---------------- Finnhub: cap in BILLIONS ---------------- */
  // Returns Map<ticker -> cap_in_bi>; converts millions -> billions; handles sandbox; graceful on errors.
  function fetchCapsFromFinnhub_Billions_(tickers) {
    const out = new Map();
    if (!tickers || !tickers.length || !FH_KEY) return out;

    let base = FH_BASE;
    if (/^sandbox_/i.test(FH_KEY)) base = 'https://sandbox.finnhub.io/api/v1';
    const headers = { 'X-Finnhub-Token': FH_KEY };

    for (let i = 0; i < tickers.length; i++) {
      const tkr = tickers[i];
      if (!tkr) continue;

      const url = base + '/stock/profile2?symbol=' + encodeURIComponent(tkr);
      try {
        const res = (function(){try{var __resp=UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers });return __resp;}catch(e){Utils.Log.append('error','UrlFetch fetch failed',''+e);throw e;}})();
        if (res.getResponseCode() !== 200) { out.set(tkr, ''); continue; }
        const json = JSON.parse(res.getContentText());

        // Finnhub commonly returns marketCapitalization in *millions* USD for profile2.
        // Normalize to *billions* for our sheet.
        let mkt = json && (json.marketCapitalization ?? json.marketCap);
        if (mkt === '' || mkt == null || isNaN(Number(mkt))) { out.set(tkr, ''); continue; }
        const capBi = Number(mkt) / 1000.0; // millions -> billions
        out.set(tkr, capBi);
      } catch (_e) {
        out.set(tkr, '');
      }
      Utilities.sleep(300);
    }
    return out;
  }

  /* ---------------- Vote math helpers ---------------- */
  // Prefer the project's utils function if available; otherwise use a stable fallback.
  function votesFromScoreRatio_(score, ratio) {
    if (typeof estimateVotesFromScoreRatio === 'function') {
      try {
        const pair = estimateVotesFromScoreRatio(score, ratio);
        if (Array.isArray(pair) && pair.length === 2 && isFinite(pair[0]) && isFinite(pair[1])) {
          return [Math.max(0, Math.round(pair[0])), Math.max(0, Math.round(pair[1]))];
        }
      } catch (_) {}
    }
    // Fallback calc: S = U - D, r = U/(U+D) -> N = S/(2r-1), U = rN, D = (1-r)N
    const S = Number(score) || 0;
    const r = Number(ratio);
    if (!isFinite(r) || r <= 0 || r >= 1 || Math.abs(2*r - 1) < 1e-6) {
      // cannot infer from ratio; best-effort: all score as ups, downs 0
      return [Math.max(0, S), 0];
    }
    const N  = S / (2*r - 1);
    let U = r * N;
    let D = (1 - r) * N;

    // integer correction so U - D == S
    U = Math.round(U);
    D = U - S;
    if (D < 0) { D = 0; U = Math.max(0, S); }
    return [U, D];
  }

  /* ---------------- Local sheet helpers ---------------- */
  function ensureSheet_(ss, name, header) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const haveCols = sh.getMaxColumns();
    if (haveCols < header.length) {
      sh.insertColumnsAfter(haveCols, header.length - haveCols);
    }
    const curHdr = safeHeader_(sh);
    if (curHdr.length !== header.length || !sameArray_(curHdr, header)) {
      sh.getRange(1, 1, 1, header.length).setValues([header]);
    }
    sh.setFrozenRows(1);
    return sh;
  }

  function safeHeader_(sh) {
    if (!sh) return [];
    const lc = sh.getLastColumn();
    if (!lc) return [];
    const vals = sh.getRange(1, 1, 1, lc).getValues()[0] || [];
    return vals.map(x => (x || '').toString().trim());
  }

  function indexFor_(hdr) {
    const m = {};
    for (let i = 0; i < hdr.length; i++) {
      const k = (hdr[i] || '').toString().trim().toLowerCase();
      if (k) m[k] = i;
    }
    return m;
  }

  function sameArray_(a, b) {
    if (!a || !b || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if ((a[i] || '') !== (b[i] || '')) return false;
    return true;
  }

  return { refresh };
})();

/* wrapper */
function FELive_refresh() { return FELive.refresh(); }
