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
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    ids.forEach((row, i) => {
      const id = String(row[0] || '').trim();
      if (id) map.set(id, i + 2);
    });
    return map;
  }

  function formatCols_(sh) {
    const H = Utils.Sheets.index(Utils.Sheets.header(sh));
    const last = sh.getLastRow();
    if (last < 2) return;
    if (H['created_utc'] != null) sh.getRange(2, H['created_utc'] + 1, last - 1, 1).setNumberFormat('0');
    if (H['created']    != null) sh.getRange(2, H['created'] + 1,    last - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }

  function fetchAndStore() {
    const sh = init_();
    const idIndex = buildIdIndex_(sh);
    
    // 1. 精準動態計算時間閘門：目前時間往前推 7 天的 Unix Timestamp
    const cutoffEpoch = Math.floor((Date.now() - (P.CUTOFF_DAYS * 24 * 60 * 60 * 1000)) / 1000);

    // --- 完整保留 Ticker Gate 機制 ---
    let __symIdx = null;
    try {
      const ss = SpreadsheetApp.getActive();
      const symSheet = ss.getSheetByName(Config.SHEETS.SYMBOLS);
      if (symSheet) {
        const symHdr   = Utils.Sheets.header(symSheet);
        const symRows  = (symSheet.getLastRow() > 1 && symSheet.getLastColumn() > 0) ? symSheet.getRange(2, 1, symSheet.getLastRow() - 1, symSheet.getLastColumn()).getValues() : [];
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
          if (typeof TickerExtractor.extract === 'function') return (t, b) => TickerExtractor.extract(t, b);
          if (typeof TickerExtractor.extractOne === 'function') return (t, b) => TickerExtractor.extractOne(t, b);
          if (typeof TickerExtractor.extractAll === 'function') return (t, b) => TickerExtractor.extractAll(t, b);
        }
        if (typeof UtilsTicker !== 'undefined' && typeof UtilsTicker.extract === 'function') {
          return (t, b) => UtilsTicker.extract(t, b);
        }
        if (typeof Ticker !== 'undefined' && typeof Ticker.extract === 'function') {
          return (t, b) => {
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
      const t = (title || '') + '';
      const b = (body || '') + '';
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
    // --- Ticker Gate 結束 ---

    // 調用專案封裝的 getApiKey 確保密鑰讀取安全
    const apiKey = Config.getApiKey('RAPIDAPI_KEY');

    let cursor = null;
    let pages = 0;
    let keepFetching = true; 
    const updates = [], appends = [];

    // 數量保險絲 (MAX_PAGES) 作為最終防禦邊界
    while (keepFetching && pages < P.MAX_PAGES) {
      pages++;
      
      let url = 'https://reddit34.p.rapidapi.com/getSearchPosts?query=DD&subreddit=wallstreetbets&sort=new';
      if (cursor) {
        url += '&cursor=' + encodeURIComponent(cursor);
      }

      Logger.log(`[常態同步] 正在讀取第 ${pages} / ${P.MAX_PAGES} 頁貼文數據...`);
      
      const response = httpFetch_(url, {
        parseAs: 'json',
        fetchOpts: {
          method: 'get',
          headers: {
            'x-rapidapi-key': apiKey,
            'x-rapidapi-host': 'reddit34.p.rapidapi.com'
          }
        }
      });

      const resData = response && response.data;
      const posts = (resData && resData.posts) ? resData.posts : [];
      
      if (!posts.length) {
        break;
      }

      for (const c of posts) {
        const d = c && c.data;
        if (!d || !d.id) continue;

        const createdUtc = Number(d.created_utc || d.created || 0);
        
        // 🚨 優先級判定：若文章時間早於設定天數 (例如 >7天)，觸發時間熔斷機制
        if (createdUtc && createdUtc < cutoffEpoch) {
          Logger.log(`[時間熔斷] 發現超過 ${P.CUTOFF_DAYS} 天前的老文章 (ID: ${d.id})，後續頁面已無須讀取，終止同步流程。`);
          keepFetching = false; 
          break; 
        }

        // 完美套用既有的 Mappers 與 Sheets 轉換工具，100% 確保欄位完全對齊
        const obj = Utils.Mappers.mapRedditPostToRowObj(d);
        const arr = Utils.Sheets.toRowArray(HEADERS, obj);
        const row = idIndex.get(d.id);

        // Ticker Gate 過濾
        try {
          const __title = (d && d.title) || '';
          const __body  = (d && d.selftext) || '';
          const __hasContent = !!String(__body).trim();
          const __hasTicker  = __hasTickerInTitleOrBody(__title, __body);
          if (!__hasContent && !__hasTicker) {
            continue; 
          }
        } catch (_e) {}

        if (row) {
          // 屬於活躍時間（7天內）的舊貼文：加入更新隊列，確保最新 upvotes/comments 能刷新並餵給 fe_live
          updates.push({ row: row, values: arr });
        } else {
          // 全新貼文
          appends.push(arr);
        }
      }

      // 檢查內部迴圈是否已判定時間熔斷，若是則不向 API 發出下一頁請求
      if (!keepFetching) {
        break;
      }

      cursor = resData.cursor;
      if (!cursor) {
        break;
      }
    }

    // 2. 安全批次更新 SOT 區塊
    if (updates.length) {
      updates.sort((a, b) => a.row - b.row);
      for (const u of updates) {
        sh.getRange(u.row, 1, 1, HEADERS.length).setValues([u.values]);
      }
      Logger.log(`[Metrics] 已成功重新整理 ${updates.length} 筆 7 天內既有貼文的互動指標。`);
    }

    // 3. 安全批次追加新資料
    if (appends.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, appends.length, HEADERS.length).setValues(appends);
      Logger.log(`[Database] 已成功寫入 ${appends.length} 筆全新貼文。`);
    }

    // 4. 標準格式化與置頂排序
    formatCols_(sh);
    Utils.Sheets.sortByHeader(sh, 'created_utc', { ascending: false });
    
    if (typeof fixRowFormat === 'function') {
      fixRowFormat(sh, 21);
    }

    Utils.Log.append('info', `RapidAPI Posts Sync: ${updates.length} updated, ${appends.length} added`, JSON.stringify({ updated: updates.length, added: appends.length }));
    Logger.log(`✅ 常態同步圓滿完成。`);
  }

  return { fetchAndStore };
})();

function Posts_fetchAndStore() {
  return Posts.fetchAndStore();
}