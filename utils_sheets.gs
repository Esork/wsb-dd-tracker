var Utils = (function() {
  var Val = {
    str:     function(v){ return (v == null) ? '' : ('' + v).trim(); },
    isBlank: function(v){ return v === '' || v === null || v === undefined; },
    oneline: function(s){ return s == null ? '' : String(s).replace(/\r?\n/g, ' '); }
  };
  var Log = {
  append: function(type, message, detail) {
    // Cloud logs (Executions UI)
    try {
      var lvl = (type || 'info').toLowerCase();
      var payload = { type: lvl, ts: new Date().toISOString(), message: String(message || ''), detail: detail || '' };
      if (lvl === 'error')      { if (console && console.error) console.error(JSON.stringify(payload)); else Logger.log(JSON.stringify(payload)); }
      else if (lvl === 'warn')  { if (console && console.warn)  console.warn(JSON.stringify(payload));  else Logger.log(JSON.stringify(payload)); }
      else                      { if (console && console.log)   console.log(JSON.stringify(payload));   else Logger.log(JSON.stringify(payload)); }
    } catch(_){}

    // Sheet log (DB)
    var ss = SpreadsheetApp.getActive();
    var sh = Sheets.ensureSheet(ss, Config.SHEETS.LOG, Config.HEADERS.LOG);
    var row = [type || 'info', new Date(), message || '', detail || ''];
    sh.appendRow(row);

    // Sort newest first + keep only last 14 days
    try {
      var hdr = Sheets.header(sh);
      var H   = Sheets.index(hdr);
      var timeCol = (H['time'] != null ? H['time'] + 1 : null);
      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      var rows = lastRow - 1;

      if (timeCol && rows > 0) {
        // sort: newest on top
        sh.getRange(2, 1, rows, lastCol).sort([{ column: timeCol, ascending: false }]);

        // retention: 14 days
        var threshold = new Date();
        threshold.setDate(threshold.getDate() - 14);

        var times = sh.getRange(2, timeCol, rows, 1).getValues();
        var toDel = [];
        for (var i = 0; i < times.length; i++) {
          var v = times[i][0];
          if (v && v instanceof Date && v < threshold) toDel.push(i + 2); // 1-based + header
        }
        for (var j = toDel.length - 1; j >= 0; j--) {
          try { sh.deleteRow(toDel[j]); } catch(e) {}
        }
      }
    } catch(e) { /* ignore sorting/retention failures */ }
  },
  info:  function(msg, detail){ return this.append('info',  msg, detail); },
  warn:  function(msg, detail){ return this.append('warn',  msg, detail); },
  error: function(msg, detail){ return this.append('error', msg, detail); }
};

  var Mappers = {
    mapRedditPostToRowObj: function(d) {
      var createdUtc  = Number(d.created_utc || d.created || 0);
      var createdDate = createdUtc ? new Date(createdUtc * 1000) : '';
      var permalink   = d.permalink ? ('https://www.reddit.com' + d.permalink) : '';
      var flair       = (d.link_flair_text || d.author_flair_text || '') + '';
      return {
        id: d.id || '',
        title: Val.oneline(d.title),
        author: d.author || '',
        created_utc: createdUtc || 0,
        created: createdDate,
        score: d.score || 0,
        ups: d.ups || '',
        upvote_ratio: d.upvote_ratio || '',
        selftext: Val.oneline(d.selftext),
        num_comments: d.num_comments || 0,
        permalink: permalink,
        url: d.url || '',
        flair: flair,
        over_18: !!d.over_18,
        stickied: !!d.stickied,
        is_video: !!d.is_video,
        is_self: !!d.is_self,
        subreddit: d.subreddit || '',
        author_fullname: d.author_fullname || '',
        total_awards_received: d.total_awards_received || 0,
        num_crossposts: d.num_crossposts || 0,
        thumbnail: d.thumbnail || ''
      };
    }
  };
  var Sheets = {
    ensureSheet: function(ss, name, header) {
      var sh = ss.getSheetByName(name) || ss.insertSheet(name);
      if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
        if (header && header.length) {
          if (sh.getMaxColumns() < header.length) {
            sh.insertColumnsAfter(sh.getMaxColumns(), header.length - sh.getMaxColumns());
          }
          sh.getRange(1,1,1,header.length).setValues([header]);
          sh.setFrozenRows(1);
        }
      }
      return sh;
    },
    header: function(sh) {
      if (!sh || sh.getLastRow() === 0 || sh.getLastColumn() === 0) return [];
      var hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      return hdr.map(function(h){ return (h || '').toString().trim(); });
    },
    index: function(header) {
      var m = {};
      for (var i=0;i<header.length;i++) {
        var k = (header[i] || '').toString().trim().toLowerCase();
        if (k) m[k] = i;
      }
      return m;
    },
    toRowArray: function(headers, obj) {
      var out = new Array(headers.length);
      for (var i=0;i<headers.length;i++) {
        var k = headers[i];
        out[i] = Object.prototype.hasOwnProperty.call(obj, k) ? obj[k] : '';
      }
      return out;
    },
    sortByHeader: function(sh, headerName, opts) {
      opts = opts || {};
      var hdr = Sheets.header(sh);
      var H   = Sheets.index(hdr);
      var key = (headerName || '').toLowerCase();
      if (!(key in H)) return;
      var col = H[key] + 1;
      var rows = sh.getLastRow() - 1;
      if (rows <= 0) return;
      sh.getRange(2,1,rows,sh.getLastColumn()).sort([{column: col, ascending: !!opts.ascending}]);
    }
  };
  return { Sheets: Sheets, Val: Val, Mappers: Mappers, Log: Log };
})();

function fixRowFormat(sheet, height) {
  height = height || 21;
  SpreadsheetApp.flush();
  try { var filter = sheet.getFilter(); if (filter) filter.remove(); } catch(_) {}
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return;
  var rng = sheet.getRange(1,1,lastRow,lastCol);
  try { rng.setWrap(false); } catch(_) {}
  try { rng.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); } catch(_) {}
  try { rng.setVerticalAlignment('middle'); } catch(_) {}

  SpreadsheetApp.flush();
  for (var r=2; r<=lastRow; r++) sheet.setRowHeight(r, height);
  SpreadsheetApp.flush();
}

function estimateVotesFromScoreRatio(score, ratio) {
  var S = Number(score);
  var R = Number(ratio);
  if (!isFinite(S) || !isFinite(R)) return [null, null];
  if (R <= 0) return [0, Math.max(0, Math.round(-S))];
  if (R >= 1) return [Math.max(0, Math.round(S)), 0];
  var denom = 2 * R - 1;               
  if (Math.abs(denom) < 1e-6) return [null, null]; 
  var T = S / denom;                    
  if (!isFinite(T) || T < 0) return [null, null];
  var U = Math.round(R * T);            
  var D = U - S;                        
  if (D < 0) {                          
    D = 0;
    U = Math.max(0, Math.round(S));
  }
  return [U, Math.round(D)];
}

/**
 * Remove junk rows where ALL are true:
 *   - ticker cell is blank
 *   - post_content/selftext is blank (or column absent)
 *   - NO real ticker can be extracted from title/body using your extractors & SymbolsIndex
 *
 * Logs per-sheet summary to Cloud logs + logger sheet.
 */

function cleanupInvalidRows() {
  const ss = SpreadsheetApp.getActive();
  const targets = [
    Config.SHEETS.POSTS,
    Config.SHEETS.FRONTEND_FIXED,
    Config.SHEETS.FRONTEND_LIVE
  ].filter(Boolean);

  // ---- Build SymbolsIndex once ----
  let symIdx = null;
  try {
    const symSheet = ss.getSheetByName(Config.SHEETS.SYMBOLS);
    if (symSheet) {
      const symHdr   = Utils.Sheets.header(symSheet);
      const symRows  = (symSheet.getLastRow() > 1 && symSheet.getLastColumn() > 0)
        ? symSheet.getRange(2, 1, symSheet.getLastRow() - 1, symSheet.getLastColumn()).getValues()
        : [];
      const symTable = { header: symHdr, rows: symRows };
      if (typeof SymbolsIndex !== 'undefined' && typeof SymbolsIndex.build === 'function') {
        symIdx = SymbolsIndex.build(symTable);
      }
    }
  } catch (e) {
    Utils.Log.warn('cleanupInvalidRows: failed to build SymbolsIndex', String(e));
  }

  // ---- Resolve a flexible extractor (reuses your existing implementations) ----
  function resolveExtractor(idx) {
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

  function pickTicker(res) {
    if (res == null) return '';
    if (Array.isArray(res))       return res[0] || '';
    if (typeof res === 'string')  return res;
    if (res instanceof Set)       { var it = res.values(); var n = it.next(); return n && !n.done ? n.value : ''; }
    if (typeof res === 'object')  return res.ticker || res.symbol || (Array.isArray(res.tickers) ? res.tickers[0] : '') || '';
    return '';
  }

  const extractor = resolveExtractor(symIdx);

  let totalDeleted = 0;
  for (const name of targets) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol === 0) continue;

    const header = sh.getRange(1,1,1,lastCol).getValues()[0]
      .map(h => (h || '').toString().trim().toLowerCase());
    const H = {}; header.forEach((h,i)=> H[h]=i);

    const iTitle   = H['title'];
    const iContent = H['post_content'] != null ? H['post_content'] : H['selftext'];
    const iTicker  = H['ticker'];

    const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
    const rowsToDelete = [];
    let scanned = 0, invalidCandidates = 0;

    for (let r = 0; r < values.length; r++) {
      scanned++;

      // Sheets without title/content (e.g., fe_live) will have undefined indices; that’s OK.
      const title   = (iTitle   != null ? (values[r][iTitle]   || '') : '') + '';
      const content = (iContent != null ? (values[r][iContent] || '') : '') + '';
      const ticker  = (iTicker  != null ? (values[r][iTicker]  || '') : '') + '';

      const hasTickerCell = !!ticker.trim();
      const hasContent    = !!content.trim();

      let extracted = '';
      try {
        if (extractor) {
          const res = extractor(title, content);
          extracted = (pickTicker(res) || '').toString().trim();
        }
      } catch (_) {
        // ignore extractor errors — we only use it as a signal
      }

      // Keep if: hasTickerCell OR extracted ticker OR hasContent
      const keep = hasTickerCell || (!!extracted) || hasContent;

      if (!keep) {
        invalidCandidates++;
        rowsToDelete.push(r + 2); // 1-based + header
      }
    }

    let deleted = 0;
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      try { sh.deleteRow(rowsToDelete[i]); deleted++; }
      catch (e) { Utils.Log.error('cleanupInvalidRows delete failed', String(e)); }
    }
    totalDeleted += deleted;

    Utils.Log.info('cleanupInvalidRows sheet summary',
      JSON.stringify({ sheet: name, scanned, candidates: invalidCandidates, deleted }));
  }

  Utils.Log.info('cleanupInvalidRows total', JSON.stringify({ deleted: totalDeleted }));
  return totalDeleted;
}
