function deleteAllTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const t of allTriggers) {
    ScriptApp.deleteTrigger(t);
  }
}

function createTriggers() {
  // First clear everything
  deleteAllTriggers();

  // ---- Weekly Trigger ----
  ScriptApp.newTrigger("Symbols_fetchAndWrite")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(20) // 8 PM
    .create();

  // ---- Daily Runs ----
  // Morning set
  createDailyRun("morning", [
    { name: "Posts_fetchAndStore", hour: 4, minute: 30 },
    { name: "FEFixed_syncChain", hour: 4, minute: 40 },
    { name: "sync_ticker_cache", hour: 4, minute: 50 },
    { name: "FELive_refreshChain", hour: 6, minute: 0 }
  ]);

  // Evening set
  createDailyRun("evening", [
    { name: "Posts_fetchAndStore", hour: 19, minute: 30 },
    { name: "FEFixed_syncChain", hour: 19, minute: 40 },
    { name: "sync_ticker_cache", hour: 19, minute: 50 },
    { name: "FELive_refreshChain", hour: 21, minute: 0 }
  ]);

  // Nightly hygiene: purge invalid rows at 03:00
  ScriptApp.newTrigger("cleanupInvalidRows")
    .timeBased()
    .atHour(3)
    .nearMinute(0)
    .everyDays(1)
    .create();
}

// ---- Helper to add daily triggers ----
function createDailyRun(label, tasks) {
  tasks.forEach(task => {
    ScriptApp.newTrigger(task.name)
      .timeBased()
      .atHour(task.hour)
      .nearMinute(task.minute)
      .everyDays(1)
      .create();
  });
}

// ---- Wrapper Chains ----
// FEFixed step chain
function FEFixed_syncChain() {
  FEFixed_syncFromPosts();
  FEFixed_fillDirections();
  FEFixed_fillTickers();
}

// FELive step chain
function FELive_refreshChain() {
  FELive_refresh();
  FEFixed_fillPricesAtPost();
}
