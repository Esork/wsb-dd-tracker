📝 wsb-dd-tracker / README.md
Markdown
# wsb-dd-tracker

A robust, modular Google Apps Script data pipeline that automatically tracks, extracts, and analyzes WallStreetBets (WSB) "DD" (Due Diligence) posts. The system parses retail sentiment, aggregates ticker mentions, and caches financial market metrics to power downstream frontend visualizations.

## 🏗️ System Architecture & Data Flow

The project functions as a linear, automated data pipeline split into 4 distinct historical and real-time data layers:

[Reddit API via RapidAPI]
│
▼

Raw Ingestion (sheet: posts)  ──► Deduplicates & updates upvotes/comments
│
▼

Immutable Fact (sheet: fe_fixed) ──► Extracts Tickers & classifies sentiment
│
├──► 3. Live Dashboard Cache (sheet: fe_live)  ──► Aggregates recent buzz
│
└──► 4. Market Price Cache (sheet: ticker_cache) ──► Caches historical stock data


### 1. Raw Data Ingestion (`sync_posts.gs` ──► `posts`)
* **Role:** The entry point of the pipeline. It queries the RapidAPI Reddit proxy (`/getSearchPosts`) twice a day to fetch the latest `wallstreetbets` posts filtered by the `DD` flair.
* **Double-Gate Throttling:** Optimizes API quota by implementing a strict time gate (`CUTOFF_DAYS = 7`) and page limit safeguard (`MAX_PAGES = 10`). The loop breaks immediately when encountering posts older than 7 days, preventing redundant API calls.
* **Metrics Synchronization:** Instead of blindly appending rows, it uses an in-memory Map index. New posts are batched into `appends`, while existing posts within the active 7-day window go into `updates` to refresh dynamic metrics like `score` and `num_comments`.

### 2. Static Analysis Layer (`sync_fe_fixed.gs` ──► `fe_fixed`)
* **Role:** Processes raw posts into immutable historical facts.
* **Extraction & Classification:** Invokes `utils_ticker_extractor.gs` to extract verified stock symbols and passes content to `utils_direction_classifier.gs` to evaluate emotional direction (Bullish, Bearish, or Neutral).
* **Immutability:** Once an article's ticker and sentiment are evaluated and committed to `fe_fixed`, they represent historical truth and are never modified. This layer serves as the ultimate dataset for future backtesting.

### 3. Dynamic Cache Layer (`sync_fe_live.gs` ──► `fe_live`)
* **Role:** Acts as a high-performance cache layer for frontend rendering.
* **Aggregation:** Aggregates rolling short-term metrics (e.g., total mentions, cumulative sentiment shifts, and engagement velocities) derived from the freshly updated raw interaction metrics.
* **Performance Optimization:** Prevents frontend queries from scanning tens of thousands of historical logs, ensuring instantaneous leaderboard rendering.

### 4. Market Data Caching (`sync_ticker_cache.gs` ──► `ticker_cache`)
* **Role:** Bridges retail narrative with real-world financial performance.
* **Financial Fetching:** Collects high-frequency tickers from the pipeline and queries financial APIs (e.g., Yahoo Finance) to retrieve corresponding historical price actions and candlestick benchmarks.
* **Quota Defense:** Caches the market data inside `ticker_cache` so that the Web UI reads from Google Sheets directly, preventing external financial API rate-limiting or unexpected overage bills.

---

## 🛠️ Project Configuration & Deployment

### Environment Variables
The pipeline utilizes Google Apps Script `ScriptProperties` to protect sensitive credentials. You must navigate to **Project Settings > Script Properties** and declare:
* `RAPIDAPI_KEY`: Your private token credential generated from your subscribed RapidAPI endpoint.

### Automation Triggers
Crons and automation sequences are centrally configured in `utils_triggers.gs`. Executing `createTriggers()` will automatically register:
* **Time-Driven Synced Framework:** Executes the data ingestion flow twice a day (04:30 & 19:30).
* **Sequential Processing Chains:** Coordinates the down-stream `FEFixed_syncChain`, `sync_ticker_cache`, and `FELive_refreshChain` in staggered 10-30 minute interval loops to guarantee flawless processing order without race conditions.
