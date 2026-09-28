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

// ---------- Creative Performance: per-brief ad results, joined by CRTV code ----------
//
// The AdPerformance sync above aggregates by BigQuery's own "media_buyer"
// column, which tracks whoever operates the ad account — not who wrote
// the creative brief. To see how each Asana brief's own creative actually
// performed, we instead join on the CRTV code: Asana's "CRTV" custom
// field (e.g. "CRTV-22716") against the same code embedded in the ad's
// name in BigQuery (e.g. "CRTV-22716_VID03"). Ads whose name has no CRTV
// code (a raw campaign asset with no matching brief) are excluded at the
// SQL level; ads whose code has no matching Asana brief are excluded at
// join time in getCreativePerformanceData().

var CREATIVE_SHEET_NAME = 'CreativePerformance';
var CRTV_REGEX = '(?i)(CRTV-[0-9]+)';

function buildCreativePerformanceQuery_(podConfig) {
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
    "  REGEXP_EXTRACT(ad_name, '" + CRTV_REGEX + "') AS crtv_code,",
    '  ' + cfg.clientNameColumn + ' AS client_name,',
    '  SUM(total_spend) AS total_spend,',
    '  SUM(count_leads) AS count_leads,',
    '  SUM(count_set) AS count_set,',
    '  SUM(count_demo) AS count_demo,',
    '  SUM(count_sold) AS count_sold,',
    '  SUM(sum_revenue) AS sum_revenue',
    'FROM `' + cfg.projectId + '.' + cfg.dataset + '.' + cfg.table + '`',
    'WHERE ' + cfg.dateColumn + ' >= ' + bqStringLiteral_(sinceDate),
    '  AND ' + cfg.clientNameColumn + ' IN (' + clientNames.join(', ') + ')',
    "  AND REGEXP_CONTAINS(ad_name, '" + CRTV_REGEX + "')",
    'GROUP BY crtv_code, client_name'
  ].join('\n');
}

function writeCreativeDataSheet_(ss, rows) {
  var sheet = getOrCreateSheetTab_(ss, CREATIVE_SHEET_NAME, [
    'crtv_code',
    'client_name',
    'total_spend',
    'count_leads',
    'count_set',
    'count_demo',
    'count_sold',
    'sum_revenue'
  ]);

  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

/**
 * Separate sync + trigger from syncAdPerformanceData, same reasoning as
 * that one being separate from the Asana sync: an unrelated failure
 * should never take down data the rest of the dashboard depends on.
 */
function syncCreativePerformanceData() {
  var podConfig = getActivePodConfig();
  var sql = buildCreativePerformanceQuery_(podConfig);
  var rawRows = runBigQueryQuery_(AD_PERFORMANCE_CONFIG.projectId, sql);

  var rows = rawRows
    .filter(function (r) {
      return r.crtv_code;
    })
    .map(function (r) {
      return [
        r.crtv_code,
        r.client_name,
        Number(r.total_spend || 0),
        Number(r.count_leads || 0),
        Number(r.count_set || 0),
        Number(r.count_demo || 0),
        Number(r.count_sold || 0),
        Number(r.sum_revenue || 0)
      ];
    });

  var ss = getOrCreateSpreadsheet_();
  writeCreativeDataSheet_(ss, rows);

  PropertiesService.getScriptProperties().setProperty('LAST_CREATIVE_SYNC_AT', new Date().toISOString());

  return { rowsSynced: rows.length };
}

function createCreativePerformanceHourlyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncCreativePerformanceData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('syncCreativePerformanceData').timeBased().everyHours(1).create();
}

function readCreativePerformanceRows_() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SHEET_ID');
  if (!sheetId) return [];

  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(CREATIVE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return values.map(function (r) {
    return {
      crtvCode: String(r[0] || '').trim().toUpperCase(),
      clientName: r[1],
      spend: r[2],
      leads: r[3],
      sets: r[4],
      demos: r[5],
      sold: r[6],
      revenue: r[7]
    };
  });
}

/**
 * The 13 metrics, in the exact order requested, computed from a raw
 * totals object ({ spend, leads, sets, demos, sold, revenue }).
 */
function computeCreativeMetrics_(t) {
  return {
    lead: t.leads,
    set: t.sets,
    demo: t.demos,
    sold: t.sold,
    setRate: t.leads ? t.sets / t.leads : 0,
    cpl: t.leads ? t.spend / t.leads : 0,
    cpSet: t.sets ? t.spend / t.sets : 0,
    cpd: t.demos ? t.spend / t.demos : 0,
    closeRate: t.demos ? t.sold / t.demos : 0,
    fbSpend: t.spend,
    revenue: t.revenue,
    avgTicket: t.sold ? t.revenue / t.sold : 0,
    com: t.revenue ? t.spend / t.revenue : 0
  };
}

function sumCreativeTotals_(rows) {
  var t = { spend: 0, leads: 0, sets: 0, demos: 0, sold: 0, revenue: 0 };
  rows.forEach(function (r) {
    t.spend += r.spend;
    t.leads += r.leads;
    t.sets += r.sets;
    t.demos += r.demos;
    t.sold += r.sold;
    t.revenue += r.revenue;
  });
  return t;
}

/**
 * Called by the dashboard's Creative Performance tab. Never calls BigQuery
 * or Asana directly — joins two already-synced sheets (CreativePerformance
 * from BigQuery, Data from Asana) by CRTV code. A creative with no
 * matching brief (or a brief with no matching ad) is silently excluded.
 */
function getCreativePerformanceData() {
  if (!PropertiesService.getScriptProperties().getProperty('LAST_CREATIVE_SYNC_AT')) {
    return { configured: false };
  }

  var podConfig = getActivePodConfig();
  var creativeRows = readCreativePerformanceRows_();
  var briefRows = readDataRows_();

  // First brief wins if a CRTV code somehow appears on more than one task.
  var briefByCode = {};
  briefRows.forEach(function (b) {
    var code = String(b.crtvCode || '').trim().toUpperCase();
    if (code && !briefByCode[code]) {
      briefByCode[code] = b;
    }
  });

  var byPerson = {};
  podConfig.people.forEach(function (person) {
    byPerson[person] = [];
  });

  creativeRows.forEach(function (c) {
    var brief = briefByCode[c.crtvCode];
    if (!brief || podConfig.people.indexOf(brief.person) === -1) {
      return;
    }
    byPerson[brief.person].push({
      crtvCode: c.crtvCode,
      account: brief.account,
      briefName: brief.name,
      creativeType: brief.creativeType,
      metrics: computeCreativeMetrics_({
        spend: c.spend,
        leads: c.leads,
        sets: c.sets,
        demos: c.demos,
        sold: c.sold,
        revenue: c.revenue
      })
    });
  });

  var totalsByPerson = {};
  podConfig.people.forEach(function (person) {
    var rawTotals = sumCreativeTotals_(
      creativeRows.filter(function (c) {
        var brief = briefByCode[c.crtvCode];
        return brief && brief.person === person;
      })
    );
    totalsByPerson[person] = computeCreativeMetrics_(rawTotals);
  });

  return {
    configured: true,
    lastCreativeSyncAt: PropertiesService.getScriptProperties().getProperty('LAST_CREATIVE_SYNC_AT') || null,
    people: podConfig.people,
    accountDisplayNames: podConfig.accountDisplayNames || {},
    byPerson: byPerson,
    totalsByPerson: totalsByPerson
  };
}
