var Config = {
  SHEETS: {
    SYMBOLS:        'symbols',
    POSTS:          'posts',
    FRONTEND_FIXED: 'fe_fixed',
    FRONTEND_LIVE:  'fe_live',
    TICKER_CACHE:   'ticker_cache',
    LOG:            'logger'
  },

  HEADERS: {
    SYMBOLS: ['symbol','name','exchange','cap_in_bi'],
    POSTS: ['id','title','author','created_utc','created','score','upvote_ratio','selftext','num_comments','permalink','url','flair','over_18','stickied','is_video','is_self','subreddit','author_fullname','total_awards_received','num_crossposts','thumbnail'],
    FRONTEND_FIXED: ['id','author','title','post_content','direction','ticker','created_at','price_at_post'],
    FRONTEND_LIVE: ['id','ticker','ups','downs','num_comments','cap_in_bi'],
    TICKER_CACHE: ['ticker','last_refreshed','date','open','high','low','close','volume'],
    LOG: ['type','time','message','detail']
  },

  API: {
    FINNHUB_BASE: 'https://finnhub.io/api/v1',
    ALPHA_VANTAGE_BASE: 'https://www.alphavantage.co/query',
    FMP_BASE: 'https://financialmodelingprep.com/api/v3',
    FINNHUB_KEY: 'FINNHUB_KEY',
    ALPHA_VANTAGE_KEY: 'ALPHA_VANTAGE_KEY',
    FMP_KEY: 'FMP_KEY'
  },

  POSTS: {
    QUERY:        'flair:"DD"',
    SUBREDDIT:    'wallstreetbets',
    PAGE_LIMIT:   100,
    USE_PROXY:    true,
    MAX_PAGES:    10,
    CUTOFF_DAYS:  7
  },

  SYMBOLS: {
    NASDAQ_LISTED_URL: 'https://www.nasdaqtrader.com/dynamic/symdir/nasdaqlisted.txt',
    OTHER_LISTED_URL:  'https://www.nasdaqtrader.com/dynamic/symdir/otherlisted.txt'
  },

  getApiKey: function(propName) {
    var props = PropertiesService.getScriptProperties();
    var v = null;
    try { v = props.getProperty(propName); } catch (e) { v = null; }
    if (!v) throw new Error('Missing script property: ' + propName);
    return v;
  }
};
