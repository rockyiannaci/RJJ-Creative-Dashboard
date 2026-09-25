/**
 * Ad performance sync: pulls creative/ad spend + funnel data from BigQuery
 * (the same source Looker Studio's fb_ads_master view reads from) and
 * writes it to a sheet, mirroring sync.gs's pattern for Asana data —
 * doGet() never queries BigQuery live, only this pre-synced sheet.
 *
 * Does nothing (and the dashboard shows a "not connected" state) until
 * BIGQUERY_SERVICE_ACCOUNT_KEY is set in Script Properties.
 */

var AD_DATA_SHEET_NAME = 'AdPerformance';
var AD_LOG_SHEET_NAME = 'AdPerformanceSyncLog';
var AD_SYNC_LOOKBACK_DAYS = 120; // how far back each sync pulls fresh rows

function buildAdPerformanceQuery_(podConfig) {
  var cfg = AD_PERFORMANCE_CONFIG;
  var clientNames = Object.keys(podConfig.adPerformanceClientNameMap).map(function (key) {
    return bqStringLiteral_(podConfig.adPerformanceClientNameMap[key]);
  });

  var sinceDate = Utilities.formatDate(
    new Date(Date.now() - AD_SYNC_LOOKBACK_DAYS * 86400000),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

  return [
    'SELECT',
    '  ' + cfg.dateColumn + ' AS date_of_lead,',
    '  ' + cfg.clientNameColumn + ' AS client_name,',
    '  ' + cfg.mediaBuyerColumn + ' AS media_buyer,',
    '  SUM(total_spend) AS total_spend,',
    '  SUM(count_leads) AS count_leads,',
    '  SUM(count_set) AS count_set,',
    '  SUM(count_demo) AS count_demo,',
    '  SUM(count_sold) AS count_sold,',
    '  SUM(sum_revenue) AS sum_revenue,',
    '  SUM(clicks) AS clicks,',
    '  SUM(impressions) AS impressions',
    'FROM `' + cfg.projectId + '.' + cfg.dataset + '.' + cfg.table + '`',
    'WHERE ' + cfg.dateColumn + ' >= ' + bqStringLiteral_(sinceDate),
    '  AND ' + cfg.clientNameColumn + ' IN (' + clientNames.join(', ') + ')',
    'GROUP BY 1, 2, 3'
  ].join('\n');
}

function writeAdDataSheet_(ss, rows) {
  var sheet = getOrCreateSheetTab_(ss, AD_DATA_SHEET_NAME, [
    'date_of_lead',
    'client_name',
    'media_buyer',
    'total_spend',
    'count_leads',
    'count_set',
    'count_demo',
    'count_sold',
    'sum_revenue',
    'clicks',
    'impressions'
  ]);

  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
  }
  // date_of_lead as plain text, same reasoning as sync.gs's week_start fix:
  // stop Sheets from silently turning it into a Date cell.
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function writeAdLogSheet_(ss, podConfig, aggregate) {
  var sheet = getOrCreateSheetTab_(ss, AD_LOG_SHEET_NAME, [
    'synced_at',
    'pod',
    'rows_returned',
    'unmapped_client_names_seen',
    'unmapped_media_buyers_seen'
  ]);

  var mappedClientNames = {};
  Object.keys(podConfig.adPerformanceClientNameMap).forEach(function (key) {
    mappedClientNames[podConfig.adPerformanceClientNameMap[key]] = true;
  });
  var mappedMediaBuyers = {};
  Object.keys(podConfig.adPerformanceMediaBuyerMap).forEach(function (key) {
    mappedMediaBuyers[podConfig.adPerformanceMediaBuyerMap[key]] = true;
  });

  var unmappedClients = {};
  var unmappedBuyers = {};
  aggregate.rawRows.forEach(function (r) {
    if (!mappedClientNames[r.client_name]) unmappedClients[r.client_name] = true;
    if (!mappedMediaBuyers[r.media_buyer]) unmappedBuyers[r.media_buyer] = true;
  });

  sheet.appendRow([
    new Date(),
    podConfig.label,
    aggregate.rawRows.length,
    Object.keys(unmappedClients).slice(0, 20).join(', '),
    Object.keys(unmappedBuyers).slice(0, 20).join(', ')
  ]);
}

/**
 * Main ad-performance sync entrypoint. Not wired into the Asana trigger —
 * install it separately (createAdPerformanceHourlyTrigger_) once BigQuery
 * access is confirmed working, so a BigQuery outage or bad query can never
 * take down the Asana sync it runs alongside.
 */
function syncAdPerformanceData() {
  var podConfig = getActivePodConfig();
  var sql = buildAdPerformanceQuery_(podConfig);
  var rawRows = runBigQueryQuery_(AD_PERFORMANCE_CONFIG.projectId, sql);

  var rows = rawRows.map(function (r) {
    return [
      r.date_of_lead,
      r.client_name,
      r.media_buyer,
      Number(r.total_spend || 0),
      Number(r.count_leads || 0),
      Number(r.count_set || 0),
      Number(r.count_demo || 0),
      Number(r.count_sold || 0),
      Number(r.sum_revenue || 0),
      Number(r.clicks || 0),
      Number(r.impressions || 0)
    ];
  });

  var ss = getOrCreateSpreadsheet_();
  writeAdDataSheet_(ss, rows);
  writeAdLogSheet_(ss, podConfig, { rawRows: rawRows });

  PropertiesService.getScriptProperties().setProperty('LAST_AD_SYNC_AT', new Date().toISOString());

  return { rowsSynced: rows.length };
}

/**
 * Run once from the editor after confirming syncAdPerformanceData() works,
 * to keep ad performance data refreshed automatically. Safe to re-run.
 */
function createAdPerformanceHourlyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAdPerformanceData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('syncAdPerformanceData').timeBased().everyHours(1).create();
}

function readAdPerformanceRows_() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SHEET_ID');
  if (!sheetId) return [];

  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(AD_DATA_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  return values.map(function (r) {
    return {
      day: normalizeDateCell_(r[0]),
      clientName: r[1],
      mediaBuyer: r[2],
      spend: r[3],
      leads: r[4],
      sets: r[5],
      demos: r[6],
      sold: r[7],
      revenue: r[8],
      clicks: r[9],
      impressions: r[10]
    };
  });
}

function computeAdMetrics_(rows) {
  var spend = 0, leads = 0, sets = 0, demos = 0, sold = 0, revenue = 0, clicks = 0, impressions = 0;
  rows.forEach(function (r) {
    spend += r.spend;
    leads += r.leads;
    sets += r.sets;
    demos += r.demos;
    sold += r.sold;
    revenue += r.revenue;
    clicks += r.clicks;
    impressions += r.impressions;
  });
  return {
    spend: spend,
    leads: leads,
    sets: sets,
    demos: demos,
    sold: sold,
    revenue: revenue,
    clicks: clicks,
    impressions: impressions,
    costPerLead: leads ? spend / leads : 0,
    roas: spend ? revenue / spend : 0,
    ctr: impressions ? clicks / impressions : 0
  };
}

/**
 * Called by the dashboard's Ad Performance tab. Never calls BigQuery —
 * only reads whatever syncAdPerformanceData() last wrote. Returns
 * configured: false until that sync has run at least once, so the rest
 * of the dashboard is never blocked on this.
 */
function getAdPerformanceData() {
  if (!PropertiesService.getScriptProperties().getProperty('LAST_AD_SYNC_AT')) {
    return { configured: false };
  }

  var podConfig = getActivePodConfig();
  var rows = readAdPerformanceRows_();

  var reverseClientMap = {};
  Object.keys(podConfig.adPerformanceClientNameMap).forEach(function (accountKey) {
    reverseClientMap[podConfig.adPerformanceClientNameMap[accountKey]] = accountKey;
  });
  var reverseBuyerMap = {};
  Object.keys(podConfig.adPerformanceMediaBuyerMap).forEach(function (personKey) {
    reverseBuyerMap[podConfig.adPerformanceMediaBuyerMap[personKey]] = personKey;
  });

  var todayKey = ymd_(new Date());
  var currentMonthKey = todayKey.substring(0, 7);
  var currentMonthRows = rows.filter(function (r) {
    return r.day && r.day.substring(0, 7) === currentMonthKey;
  });

  var byAccount = {};
  podConfig.accounts.forEach(function (account) {
    var accountRows = currentMonthRows.filter(function (r) {
      return reverseClientMap[r.clientName] === account;
    });
    byAccount[account] = computeAdMetrics_(accountRows);
  });

  var byPerson = {};
  podConfig.people.forEach(function (person) {
    var personRows = currentMonthRows.filter(function (r) {
      return reverseBuyerMap[r.mediaBuyer] === person;
    });
    byPerson[person] = computeAdMetrics_(personRows);
  });

  return {
    configured: true,
    lastAdSyncAt: PropertiesService.getScriptProperties().getProperty('LAST_AD_SYNC_AT') || null,
    currentMonthKey: currentMonthKey,
    overall: computeAdMetrics_(currentMonthRows),
    byAccount: byAccount,
    byPerson: byPerson,
    accounts: podConfig.accounts,
    accountDisplayNames: podConfig.accountDisplayNames || {},
    people: podConfig.people
  };
}
