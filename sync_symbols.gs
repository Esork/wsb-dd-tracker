const Symbols = (() => {
  const SHEET_NAME = Config.SHEETS.SYMBOLS;
  const HEADERS    = Config.HEADERS.SYMBOLS;

  function init_() {
    const ss = SpreadsheetApp.getActive();
    return Utils.Sheets.ensureSheet(ss, SHEET_NAME, HEADERS);
  }

  function parseNasdaq_(txt) {
    // nasdaqlisted.txt: pipe-delimited, with header + footer
    const lines = txt.split(/\r?\n/).filter(Boolean);
    const out = [];
    for (let i=1; i<lines.length; i++) { // skip header
      const line = lines[i];
      if (line.indexOf('File Creation Time') === 0) break;
      const parts = line.split('|');
      const symbol = parts[0], name = parts[1];
      if (!symbol || symbol === 'Symbol') continue;
      out.push({ symbol: symbol, name: name, exchange: 'NASDAQ', cap_in_bi: '' });
    }
    return out;
  }

  function parseOther_(txt) {
    // otherlisted.txt: ACT Symbol|Security Name|Exchange|...
    const lines = txt.split(/\r?\n/).filter(Boolean);
    const out = [];
    for (let i=1; i<lines.length; i++) {
      const line = lines[i];
      if (line.indexOf('File Creation Time') === 0) break;
      const p = line.split('|');
      const act = p[0], name = p[1], exchCode = p[2];
      if (!act || act === 'ACT Symbol') continue;
      var exch = exchCode;
      if (exchCode === 'N') exch = 'NYSE';
      else if (exchCode === 'A') exch = 'NYSE MKT';
      else if (exchCode === 'P') exch = 'NYSE ARCA';
      out.push({ symbol: act, name: name, exchange: exch, cap_in_bi: '' });
    }
    return out;
  }

  function fetchAndWrite() {
    const sh = init_();
    const nasdaqTxt = fetchNasdaqListed_();
    const otherTxt  = fetchOtherListed_();
    const rows = parseNasdaq_(nasdaqTxt).concat(parseOther_(otherTxt));
    const seen = new Set();
    const dedup = [];
    for (const r of rows) {
      if (seen.has(r.symbol)) continue;
      seen.add(r.symbol);
      dedup.push(r);
    }

    const data = dedup.map(obj => Utils.Sheets.toRowArray(HEADERS, obj));
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2,1,last-1,sh.getLastColumn()).clearContent();
    if (data.length) sh.getRange(2,1,data.length,HEADERS.length).setValues(data);

    Utils.Sheets.sortByHeader(sh, 'symbol', { ascending: true });
    fixRowFormat(sh, 21);

    Utils.Log.append('info', String(`Symbols: wrote ${data.length} rows`, 'WSB', 4), '');
  }

  return { fetchAndWrite };
})();

function Symbols_fetchAndWrite() { return Symbols.fetchAndWrite(); }
