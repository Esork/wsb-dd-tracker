const Posts = (() => {
  const SHEET_NAME = Config.SHEETS.POSTS;
  const HEADERS    = Config.HEADERS.POSTS;
  const P          = Config.POSTS;

  function init_() {
    const ss = SpreadsheetApp.getActive();
    return Utils.Sheets.ensureSheet(ss, SHEET_NAME, HEADERS);
  }

  function buildIdIndex_(sh) {
    const last = sh.getLastRow();
    const map = new Map();
    if (last < 2) return map;
    const ids = sh.getRange(2,1,last-1,1).getValues();
    ids.forEach((row,i) => {
      const id = String(row[0] || '').trim();
      if (id) map.set(id, i+2);
    });
    return map;
  }

  function formatCols_(sh) {
    const H = Utils.Sheets.index(Utils.Sheets.header(sh));
    const last = sh.getLastRow();
    if (last < 2) return;
    if (H['created_utc'] != null) sh.getRange(2, H['created_utc']+1, last-1, 1).setNumberFormat('0');
    if (H['created']    != null) sh.getRange(2, H['created']+1,    last-1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }

  function fetchAndStore() {
    const sh = init_();
    const idIndex = buildIdIndex_(sh);

    const cutoffEpoch = (function(){
      const d = new Date();
      d.setDate(d.getDate() - P.CUTOFF_DAYS);
      return Math.floor(d.getTime()/1000);
    })();


    // --- Added: lightweight ticker gate for invalid posts (no content & no ticker) ---
    // Build a symbols index and a flexible extractor (if available)
    let __symIdx = null;
    try {
      const ss = SpreadsheetApp.getActive();
      const symSheet = ss.getSheetByName(Config.SHEETS.SYMBOLS);
      if (symSheet) {
        const symHdr   = Utils.Sheets.header(symSheet);
        const symRows  = (symSheet.getLastRow() > 1 && symSheet.getLastColumn() > 0)
          ? symSheet.getRange(2, 1, symSheet.getLastRow() - 1, symSheet.getLastColumn()).getValues()
          : [];
        const symTable = { header: symHdr, rows: symRows };
        if (typeof SymbolsIndex !== 'undefined' && typeof SymbolsIndex.build === 'function') {
          __symIdx = SymbolsIndex.build(symTable);
        }
      }
    } catch (e) {
      // ignore index build errors
    }

    function __resolveExtractor(idx) {
      try {
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
        if (typeof Ticker !== 'undefined' && typeof Ticker.extract === 'function') {
          return (t,b) => {
            let x = Ticker.extract(t, idx);
            if (!x && typeof Ticker.extractFallback === 'function') x = Ticker.extractFallback(t, b, idx);
            return x || '';
          };
        }
      } catch (_) {}
      return null;
    }

    function __pickTicker(res) {
      if (res == null) return '';
      if (Array.isArray(res))       return res[0] || '';
      if (typeof res === 'string')  return res;
      if (res instanceof Set)       { var it = res.values(); var n = it.next(); return n && !n.done ? n.value : ''; }
      if (typeof res === 'object')  return res.ticker || res.symbol || (Array.isArray(res.tickers) ? res.tickers[0] : '') || '';
      return '';
    }

    const __extractor = __resolveExtractor(__symIdx);

    function __hasTickerInTitleOrBody(title, body) {
      const t = (title || '')+''; const b = (body || '')+'';
      try {
        if (__extractor) {
          const res = __extractor(t, b);
          const picked = __pickTicker(res);
          if (picked && String(picked).trim()) return true;
        }
      } catch (_) {}
      const pat = /(^|\s|\$)[A-Z]{1,5}(\b|\/|\s)/;
      return pat.test(t) || pat.test(b);
    }
    // --- End added helpers ---
    let after = null, pages = 0;
    const updates = [], appends = [];

    while (pages < P.MAX_PAGES) {
      pages++;
      const json = fetchRedditSearch_({
        q: P.QUERY,
        subreddit: P.SUBREDDIT,
        sort: 'new',
        limit: P.PAGE_LIMIT,
        after: after,
        useProxy: P.USE_PROXY,
        useProxyFallback: true,
        raw_json: 1
      });

      const data = json && json.data;
      const children = (data && data.children) ? data.children : [];
      if (!children.length) break;

      for (const c of children) {
        const d = c && c.data;
        if (!d || !d.id) continue;
        const createdUtc = Number(d.created_utc || d.created || 0);
        if (cutoffEpoch && createdUtc && createdUtc < cutoffEpoch) continue;

        const obj = Utils.Mappers.mapRedditPostToRowObj(d);
        const arr = Utils.Sheets.toRowArray(HEADERS, obj);
        const row = idIndex.get(d.id);
        // --- Gate: skip posts with no content AND no ticker in title/body ---
        try {
          const __title = (d && d.title) || '';
          const __body  = (d && d.selftext) || '';
          const __hasContent = !!String(__body).trim();
          const __hasTicker  = __hasTickerInTitleOrBody(__title, __body);
          if (!__hasContent && !__hasTicker) {
            continue; // skip invalid post
          }
        } catch (_e) {}
        // --- End gate ---

        if (row) updates.push({row: row, values: arr});
        else     appends.push(arr);
      }

      after = data.after;
      if (!after) break;
    }

    if (updates.length) {
      updates.sort((a,b)=>a.row-b.row);
      for (const u of updates) sh.getRange(u.row, 1, 1, HEADERS.length).setValues([u.values]);
    }
    if (appends.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, appends.length, HEADERS.length).setValues(appends);
    }

    formatCols_(sh);
    Utils.Sheets.sortByHeader(sh, 'created_utc', { ascending: false });
    fixRowFormat(sh, 21);

    Utils.Log.append('info', `Posts: ${updates.length} updated, ${appends.length} added`, JSON.stringify({updated: updates.length, added: appends.length}));
  }

  return { fetchAndStore };
})();

function Posts_fetchAndStore() { return Posts.fetchAndStore(); }
