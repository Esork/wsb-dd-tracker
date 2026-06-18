const Direction = (() => {
  function clean(text) {
    if (!text) return '';
    let s = String(text);
    s = s.replace(/<[^>]+>/g, ' ')
         .replace(/&nbsp;/gi, ' ')
         .replace(/&amp;/gi, '&')
         .replace(/&lt;/gi, '<')
         .replace(/&gt;/gi, '>')
         .replace(/\s+/g, ' ')
         .trim()
         .toLowerCase();
    return s;
  }

  function hasNonDilutiveCue(s) {
    return /non[-\s]?dilut/i.test(s) ||
           /no (risk of )?dilution/i.test(s) ||
           /little to no[^.]{0,30}dilution/i.test(s);
  }

  const POS_GENERIC = [
    /(^|\W)long(\W|$)/, /(^|\W)bull(ish)?(\W|$)/, /(^|\W)buy(ing)?(\W|$)/,
    /(^|\W)call(s| option|s option)?(\W|$)/, /price\s*target|(^|\W)pt(\W|$)/,
    /(^|\W)beat(s|en)?(\W|$)/, /upgrade(d|s)?\b/, /raise(d)?\s+guidance/,
    /undervalued/, /break(out|ing)\b/, /up(side|trend)\b/,
    /short\s*squeeze/, /\bbtfd\b|\bbuy\s+the\s+dip\b/,
    /\bmoon(ing)?\b/, /\byolo\b/, /multi[-\s]?bagger|multibagger/,
    /room to grow|growth runway|multi[-\s]?year growth/
  ];

  const POS_DOMAIN = [
    /partnership|exclusive partnership|collab(oratio)?n/,
    /contract(s)?\b|task order(s)?\b|award(s)?\b|backlog\b|order(s)?\b/,
    /\bprogram of record\b|\bpor\b/, /\bdod\b|\busaf\b|\bafwerx\b|\bprime\b/,
    /acquisitio(n|ns)\b|facility\b|hiring\b|expanding|build[-\s]?up/,
    /raise(d)?(\s+\$|\s+capital| funding)/, /\bfunding\b/,
    /\bcertification\b|\btype certification\b|\bfaa\b|\bapproval\b/,
    /\bota\b(?!\w)/,
    /delivered|delivery|first (vehicle|aircraft|unit)\b/
  ];

  const NEG_GENERIC = [
    /(^|\W)short(?!\s*squeeze)(\W|$)/, /(^|\W)bear(ish)?(\W|$)/,
    /(^|\W)sell(ing)?(\W|$)/, /(^|\W)put(s| option|s option)?(\W|$)/,
    /\bmiss(ed)?\b/, /\bdowngrade(d|s)?\b/, /cut\s+guidance/,
    /overvalued/, /\bfraud\b|\bscam\b/, /(chapter\s*11|bankrupt|bk)\b/,
    /dump(ing)?\b/, /(downtrend|sell[-\s]*off|going\s+down)\b/,
    /\bsec\b[^.]{0,15}\b(investigation|probe)\b/,
    /\bdilution\b|\boffering\b|\breverse split\b|\br\/s\b/
  ];

  const EXC_BULL = [
    /(sell(ing)?|writing)\s+puts?/, /\bselling\s+csp(s)?\b/,
    /non[-\s]?dilut(ive|ion)/, /no (risk of )?dilution/, /little to no[^.]{0,30}dilution/
  ];
  const EXC_BEAR = [ /(sell(ing)?|writing)\s+calls?/, /\bcovered calls?\b/ ];

  const TITLE_BOOST = [
    /blow the lid off|game[-\s]?changer|about to (rip|explode|run)/,
    /new contract|big (deal|award)|major (catalyst|program)/,
    /exclusive|official|confirmed|approved/
  ];

  const POS_POSITION = [ /position:\s*\d+(\,\d{3})*\s*share/i ];

  function scoreSection(text, weight, isTitle) {
    if (!text) return 0;
    weight = weight || 1;
    const s = clean(text);
    let score = 0;

    EXC_BULL.forEach(r => { if (r.test(s)) score += 1 * weight; });
    EXC_BEAR.forEach(r => { if (r.test(s)) score -= 1 * weight; });

    POS_DOMAIN.forEach(r => { if (r.test(s)) score += 3 * weight; });
    POS_GENERIC.forEach(r => { if (r.test(s)) score += 2 * weight; });

    if (/\bdilution\b/.test(s) || /\boffering\b/.test(s)) {
      score += hasNonDilutiveCue(s) ? (2 * weight) : (-2 * weight);
    }

    NEG_GENERIC.forEach(r => {
      if (/(dilution|offering)/.test(String(r))) return;
      if (r.test(s)) score -= 2 * weight;
    });

    POS_POSITION.forEach(r => { if (r.test(s)) score += 2 * weight; });

    if (isTitle) TITLE_BOOST.forEach(r => { if (r.test(s)) score += 2 * weight; });

    const bangs = (s.match(/!/g) || []).length;
    if (bangs >= 3) score += 0.5 * Math.sign(score || 1) * weight;

    return score;
  }

  function classify(title, body, meta) {
    const tScore = scoreSection(title, 3, true);
    const bScore = scoreSection(body, 1.5, false);
    let total = tScore + bScore;

    const upr = meta && typeof meta.upvote_ratio === 'number'
      ? meta.upvote_ratio
      : (meta && !isNaN(parseFloat(meta.upvote_ratio)) ? parseFloat(meta.upvote_ratio) : NaN);

    if (!isNaN(upr)) {
      if (upr >= 0.60) total += 0.5;
      else if (upr <= 0.45) total -= 0.5;
    }

    if (total >= 0.75) return 'Bull';
    if (total <= -0.75) return 'Bear';
    return 'Neutral';
  }

  return { classify };
})();

function detectDirection(title, bodyOrMeta, maybeMeta) {
  var body = bodyOrMeta;
  var meta = maybeMeta || {};
  if (body === undefined && typeof title === 'string') {
    body = '';
  }
  return Direction.classify(title, body, meta);
}
function classifyDirection(title, bodyOrMeta, maybeMeta) {
  return detectDirection(title, bodyOrMeta, maybeMeta);
}

var DirectionClassifier = {
  classify: function(title, body, meta) {
    return Direction.classify(title, body, meta || {});
  }
};

var UtilsDirection = {
  classify: function(title, body, meta) {
    return Direction.classify(title, body, meta || {});
  }
};
