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
// Fixed start date (not a rolling lookback) so every month tab has a
// complete month's worth of data, rather than the oldest visible month
// being partial depending on when the sync last ran. Held at June 2026 for
// now — the data team's Jan–May figures are unreliable at this account's
// current spend/lead volume; move this back once they can re-pull it.
var CREATIVE_SYNC_START_DATE = '2026-06-01';

function buildCreativePerformanceQuery_(podConfig) {
  var cfg = AD_PERFORMANCE_CONFIG;
  var clientNames = Object.keys(podConfig.adPerformanceClientNameMap).map(function (key) {
    return bqStringLiteral_(podConfig.adPerformanceClientNameMap[key]);
  });

  return [
    'SELECT',
    "  REGEXP_EXTRACT(ad_name, '" + CRTV_REGEX + "') AS crtv_code,",
    // Facebook prefixes some ad names with a status marker (e.g. a green
    // circle emoji) that isn't part of the creative's actual name — strip
    // any leading non-alphanumeric characters so "🟢 Collage V19" and
    // "Collage V19" group together as the same creative instead of
    // splitting its revenue across two look-alike rows.
    "  TRIM(REGEXP_REPLACE(ad_name, '^[^A-Za-z0-9]+', '')) AS ad_name,",
    '  ' + cfg.clientNameColumn + ' AS client_name,',
    "  FORMAT_DATE('%Y-%m', " + cfg.dateColumn + ') AS month,',
    '  SUM(total_spend) AS total_spend,',
    '  SUM(count_leads) AS count_leads,',
    '  SUM(count_set) AS count_set,',
    '  SUM(count_demo) AS count_demo,',
    '  SUM(count_sold) AS count_sold,',
    '  SUM(sum_revenue) AS sum_revenue',
    'FROM `' + cfg.projectId + '.' + cfg.dataset + '.' + cfg.table + '`',
    'WHERE ' + cfg.dateColumn + ' >= ' + bqStringLiteral_(CREATIVE_SYNC_START_DATE),
    '  AND ' + cfg.clientNameColumn + ' IN (' + clientNames.join(', ') + ')',
    // "Unassigned ..." rows are BigQuery's own catch-all bucket for spend
    // that never resolved to a specific ad, not a real creative — exclude
    // them rather than crediting that spend/revenue to any media buyer.
    "  AND NOT REGEXP_CONTAINS(ad_name, '(?i)unassigned')",
    // No CRTV code (a static image, per the pod's convention) still counts —
    // getCreativePerformanceData() attributes those to Rocky directly.
    'GROUP BY crtv_code, ad_name, client_name, month'
  ].join('\n');
}

function writeCreativeDataSheet_(ss, rows) {
  var sheet = getOrCreateSheetTab_(ss, CREATIVE_SHEET_NAME, [
    'crtv_code',
    'ad_name',
    'client_name',
    'month',
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
    // Force the month column (D, "yyyy-MM") to plain text so Sheets doesn't
    // auto-convert it into a Date cell, which would break the dashboard's
    // string-based month tabs.
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

/**
 * Sheets can auto-convert a "yyyy-MM" string into a real Date cell despite
 * the '@' text format above (if the cell already held a Date from before
 * that format was applied), so normalize both possible shapes back to
 * "yyyy-MM" on read, the same way normalizeDateCell_ does for week_start.
 */
function normalizeMonthCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return value;
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

  var rows = rawRows.map(function (r) {
    return [
      r.crtv_code || '',
      r.ad_name || '',
      r.client_name,
      r.month,
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

  // Runs after the main data is saved, so a failure here still leaves the
  // month-level data fresh (and shows up as an error on the trigger).
  var windowRows = syncCreativeWindowData_();

  return { rowsSynced: rows.length, windowRowsSynced: windowRows };
}

var CREATIVE_WINDOW_SHEET_NAME = 'CreativeWindow';
var CREATIVE_WINDOW_KEY = 'w4214';
// Leads from 42 days ago through 14 days ago: old enough that the 14-day
// lead-to-set and 42-day lead-to-sold cycles have mostly played out.
var MATURED_WINDOW_START_DAYS_BACK = 42;
var MATURED_WINDOW_END_DAYS_BACK = 14;

function formatDayOffset_(daysBack) {
  return Utilities.formatDate(
    new Date(Date.now() - daysBack * 86400000),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

/**
 * Same shape as the main creative query, but one row per creative over a
 * single rolling date range instead of per month — a 42-to-14-days-ago
 * window can't be built from month buckets.
 */
function buildCreativeWindowQuery_(podConfig, startDate, endDate) {
  var cfg = AD_PERFORMANCE_CONFIG;
  var clientNames = Object.keys(podConfig.adPerformanceClientNameMap).map(function (key) {
    return bqStringLiteral_(podConfig.adPerformanceClientNameMap[key]);
  });

  return [
    'SELECT',
    "  REGEXP_EXTRACT(ad_name, '" + CRTV_REGEX + "') AS crtv_code,",
    "  TRIM(REGEXP_REPLACE(ad_name, '^[^A-Za-z0-9]+', '')) AS ad_name,",
    '  ' + cfg.clientNameColumn + ' AS client_name,',
    '  SUM(total_spend) AS total_spend,',
    '  SUM(count_leads) AS count_leads,',
    '  SUM(count_set) AS count_set,',
    '  SUM(count_demo) AS count_demo,',
    '  SUM(count_sold) AS count_sold,',
    '  SUM(sum_revenue) AS sum_revenue',
    'FROM `' + cfg.projectId + '.' + cfg.dataset + '.' + cfg.table + '`',
    'WHERE ' + cfg.dateColumn + ' >= ' + bqStringLiteral_(startDate),
    '  AND ' + cfg.dateColumn + ' <= ' + bqStringLiteral_(endDate),
    '  AND ' + cfg.clientNameColumn + ' IN (' + clientNames.join(', ') + ')',
    "  AND NOT REGEXP_CONTAINS(ad_name, '(?i)unassigned')",
    'GROUP BY crtv_code, ad_name, client_name'
  ].join('\n');
}

function syncCreativeWindowData_() {
  var podConfig = getActivePodConfig();
  var startDate = formatDayOffset_(MATURED_WINDOW_START_DAYS_BACK);
  var endDate = formatDayOffset_(MATURED_WINDOW_END_DAYS_BACK);
  var rawRows = runBigQueryQuery_(
    AD_PERFORMANCE_CONFIG.projectId,
    buildCreativeWindowQuery_(podConfig, startDate, endDate)
  );

  var rows = rawRows.map(function (r) {
    return [
      r.crtv_code || '',
      r.ad_name || '',
      r.client_name,
      Number(r.total_spend || 0),
      Number(r.count_leads || 0),
      Number(r.count_set || 0),
      Number(r.count_demo || 0),
      Number(r.count_sold || 0),
      Number(r.sum_revenue || 0)
    ];
  });

  var sheet = getOrCreateSheetTab_(getOrCreateSpreadsheet_(), CREATIVE_WINDOW_SHEET_NAME, [
    'crtv_code',
    'ad_name',
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

  PropertiesService.getScriptProperties().setProperties({
    CREATIVE_WINDOW_START: startDate,
    CREATIVE_WINDOW_END: endDate
  });
  return rows.length;
}

function readCreativeWindowRows_() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) return [];

  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(CREATIVE_WINDOW_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues().map(function (r) {
    return {
      crtvCode: String(r[0] || '').trim().toUpperCase(),
      adName: r[1],
      clientName: r[2],
      spend: r[3],
      leads: r[4],
      sets: r[5],
      demos: r[6],
      sold: r[7],
      revenue: r[8]
    };
  });
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

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  return values.map(function (r) {
    return {
      crtvCode: String(r[0] || '').trim().toUpperCase(),
      adName: r[1],
      clientName: r[2],
      month: normalizeMonthCell_(r[3]),
      spend: r[4],
      leads: r[5],
      sets: r[6],
      demos: r[7],
      sold: r[8],
      revenue: r[9]
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

/**
 * BigQuery client_name -> this pod's account key, the reverse of
 * adPerformanceClientNameMap. Used to attribute a static (no-CRTV) ad back
 * to an account without going through the Asana brief join.
 */
function buildReverseClientNameMap_(podConfig) {
  var reverse = {};
  Object.keys(podConfig.adPerformanceClientNameMap).forEach(function (account) {
    reverse[podConfig.adPerformanceClientNameMap[account]] = account;
  });
  return reverse;
}

/**
 * Static creatives (no CRTV code) never go through an Asana brief, so
 * there's no code to join on — attribution here is a fixed convention per
 * account/name, agreed with the pod directly rather than inferred:
 *   - "Collage ..." -> Rocky, on every account
 *   - RBA-ENY / RBA-QC, no CRTV code, not a Collage -> Julian
 *   - Leaf Home / Bath Planet, no CRTV code, not a Collage -> "Other"
 *   - Refloor / RBA-GW, no CRTV code, not a Collage -> Rocky (default)
 * "Other" is intentionally not one of podConfig.people: it's excluded from
 * the media-buyer totals table by construction (nothing in that table
 * loops over a person outside podConfig.people), but still shows up in the
 * Top 25 / creative-level views so its revenue isn't silently dropped.
 */
function attributeStaticCreativeRow_(c, account) {
  var isCollage = /collage/i.test(c.adName || '');
  var person;
  if (isCollage) {
    person = 'Rocky Iannaci';
  } else if (account === 'Renewal By Andersen - ENY' || account === 'Renewal By Andersen - QC') {
    person = 'Julian DiVito';
  } else if (account === 'Leaf Home Enhancements' || account === 'Bath Planet') {
    person = 'Other';
  } else {
    person = 'Rocky Iannaci';
  }
  return {
    person: person,
    account: account,
    briefName: c.adName || '(static image)',
    creativeType: 'Static/Raw Asset'
  };
}

function attributeCreativeRow_(c, briefByCode, reverseClientNameMap, podConfig) {
  var account = reverseClientNameMap[c.clientName] || c.clientName;

  if (!c.crtvCode) {
    return attributeStaticCreativeRow_(c, account);
  }

  var brief = briefByCode[c.crtvCode];
  // A CRTV code is only unique within one account's Asana project, not
  // globally in BigQuery — another account entirely could reuse the same
  // code (or an ad naming coincidence) on an ad tagged with a different
  // client_name. Only trust the match when BigQuery's client_name for this
  // row actually resolves back to the SAME account the brief was written
  // for.
  var expectedClientName = brief ? podConfig.adPerformanceClientNameMap[brief.account] : null;
  var resolved = brief && podConfig.people.indexOf(brief.person) !== -1 && c.clientName === expectedClientName;

  if (resolved) {
    return {
      person: brief.person,
      account: brief.account,
      briefName: brief.name,
      creativeType: brief.creativeType
    };
  }

  // Refloor's CRTV-coded ads that don't resolve to a known brief/media
  // buyer still show up (as "Other") in the Top 25 rather than being
  // dropped, since Refloor's spend is too large to silently exclude.
  // Every other account keeps the old behavior: exclude rather than guess.
  if (account === 'Refloor') {
    return {
      person: 'Other',
      account: 'Refloor',
      briefName: c.adName || c.crtvCode,
      creativeType: 'Unmatched'
    };
  }

  return null;
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
  var reverseClientNameMap = buildReverseClientNameMap_(podConfig);

  // First brief wins if a CRTV code somehow appears on more than one task.
  var briefByCode = {};
  briefRows.forEach(function (b) {
    var code = String(b.crtvCode || '').trim().toUpperCase();
    if (code && !briefByCode[code]) {
      briefByCode[code] = b;
    }
  });

  // Every row that resolves to a person in this pod — either via its
  // brief's CRTV code, or (for a static, no-CRTV ad) the fixed
  // STATIC_CREATIVE_PERSON convention — counts toward that person's totals.
  var matchedRows = [];
  var attribution = {}; // keyed by array index in matchedRows
  creativeRows.forEach(function (c) {
    var attr = attributeCreativeRow_(c, briefByCode, reverseClientNameMap, podConfig);
    if (!attr) return;
    attribution[matchedRows.length] = attr;
    matchedRows.push(c);
  });

  var months = Object.keys(
    matchedRows.reduce(function (acc, c) {
      acc[c.month] = true;
      return acc;
    }, {})
  ).sort();

  function totalsForRows(rows, attrs) {
    var totals = {};
    podConfig.people.forEach(function (person) {
      var rawTotals = sumCreativeTotals_(
        rows.filter(function (c, i) {
          return attrs[i].person === person;
        })
      );
      totals[person] = computeCreativeMetrics_(rawTotals);
    });
    return totals;
  }

  var allAttrs = matchedRows.map(function (c, i) { return attribution[i]; });
  var totalsByMonth = { all: totalsForRows(matchedRows, allAttrs) };
  months.forEach(function (month) {
    var rowsForMonth = [];
    var attrsForMonth = [];
    matchedRows.forEach(function (c, i) {
      if (c.month === month) {
        rowsForMonth.push(c);
        attrsForMonth.push(allAttrs[i]);
      }
    });
    totalsByMonth[month] = totalsForRows(rowsForMonth, attrsForMonth);
  });

  // Brief-level rows for the "top revenue creatives" breakdown and each
  // media buyer's expandable detail table. Filtering/sorting/ranking by
  // month happens client-side against this flat list.
  function buildCreativeEntry(c, attr, month) {
    return {
      crtvCode: c.crtvCode || null,
      creativeName: c.adName || attr.briefName,
      account: attr.account,
      briefName: attr.briefName,
      creativeType: attr.creativeType,
      person: attr.person,
      month: month,
      // Raw totals alongside the pre-computed ratios: the dashboard's "All
      // Time" and "Top 25" views combine the same creative across several
      // months, and ratios (CPL, COM, ...) can't just be averaged — they
      // have to be recomputed from summed raw numbers.
      raw: {
        spend: c.spend,
        leads: c.leads,
        sets: c.sets,
        demos: c.demos,
        sold: c.sold,
        revenue: c.revenue
      },
      metrics: computeCreativeMetrics_({
        spend: c.spend,
        leads: c.leads,
        sets: c.sets,
        demos: c.demos,
        sold: c.sold,
        revenue: c.revenue
      })
    };
  }

  var creatives = matchedRows.map(function (c, i) {
    return buildCreativeEntry(c, allAttrs[i], c.month);
  });

  // The rolling 42-to-14-days-ago window lives in its own sheet (it can't be
  // built from month buckets). Its rows ride along in the same lists under a
  // reserved key; the dashboard keeps them out of "All Time" by that key.
  var props = PropertiesService.getScriptProperties();
  var windowStart = props.getProperty('CREATIVE_WINDOW_START');
  var windowEnd = props.getProperty('CREATIVE_WINDOW_END');
  var windowInfo = null;
  if (windowStart && windowEnd) {
    var windowRows = [];
    var windowAttrs = [];
    readCreativeWindowRows_().forEach(function (c) {
      var attr = attributeCreativeRow_(c, briefByCode, reverseClientNameMap, podConfig);
      if (!attr) return;
      windowRows.push(c);
      windowAttrs.push(attr);
    });
    totalsByMonth[CREATIVE_WINDOW_KEY] = totalsForRows(windowRows, windowAttrs);
    windowRows.forEach(function (c, i) {
      creatives.push(buildCreativeEntry(c, windowAttrs[i], CREATIVE_WINDOW_KEY));
    });
    windowInfo = { key: CREATIVE_WINDOW_KEY, start: windowStart, end: windowEnd };
  }

  return {
    configured: true,
    lastCreativeSyncAt: PropertiesService.getScriptProperties().getProperty('LAST_CREATIVE_SYNC_AT') || null,
    people: podConfig.people,
    accounts: podConfig.accounts,
    accountDisplayNames: podConfig.accountDisplayNames || {},
    creatives: creatives,
    months: months,
    totalsByMonth: totalsByMonth,
    window: windowInfo
  };
}

var RECENTLY_LAUNCHED_WINDOW_DAYS = 35;

/**
 * Called by the dashboard's "Recently Launched Creatives" section. A brief
 * counts as recently launched if its Asana due date falls within the last
 * RECENTLY_LAUNCHED_WINDOW_DAYS days — due_on is what the pod actually
 * treats as a creative's launch date, not created_at. Performance is
 * summed across every ad tied to that brief's CRTV code (all its geo/status
 * variants), since this reports on the BRIEF as a unit, not one specific ad
 * execution the way the Top 25 table does.
 */
function getRecentlyLaunchedCreatives() {
  if (!PropertiesService.getScriptProperties().getProperty('LAST_SYNC_AT')) {
    return { configured: false };
  }

  var podConfig = getActivePodConfig();
  var briefRows = readDataRows_();
  var creativeRows = readCreativePerformanceRows_();

  var cutoff = Utilities.formatDate(
    new Date(Date.now() - RECENTLY_LAUNCHED_WINDOW_DAYS * 86400000),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var recentBriefs = briefRows.filter(function (b) {
    return b.crtvCode && b.dueOn && b.dueOn >= cutoff && b.dueOn <= today;
  });

  var byAccount = {};
  podConfig.accounts.forEach(function (account) {
    byAccount[account] = [];
  });

  recentBriefs.forEach(function (b) {
    var code = String(b.crtvCode || '').trim().toUpperCase();
    var expectedClientName = podConfig.adPerformanceClientNameMap[b.account];
    var matchingAdRows = creativeRows.filter(function (c) {
      return c.crtvCode === code && c.clientName === expectedClientName;
    });
    var rawTotals = sumCreativeTotals_(matchingAdRows);

    // The Asana task name is a brief description, not the actual creative
    // name Ads Manager/Looker use — show the real ad_name from whichever
    // matching BigQuery row has the most revenue (the brief's CRTV code can
    // cover several ad variants). Falls back to the brief name only if it
    // hasn't started spending yet, so there's nothing to join to.
    var topAdRow = matchingAdRows.reduce(function (best, c) {
      return !best || c.revenue > best.revenue ? c : best;
    }, null);
    var creativeName = topAdRow ? topAdRow.adName : b.name;

    if (!byAccount[b.account]) byAccount[b.account] = [];
    byAccount[b.account].push({
      crtvCode: code,
      creativeName: creativeName,
      briefName: b.name,
      account: b.account,
      dueOn: b.dueOn,
      person: b.person,
      creativeType: b.creativeType,
      metrics: computeCreativeMetrics_(rawTotals)
    });
  });

  Object.keys(byAccount).forEach(function (account) {
    byAccount[account].sort(function (a, b) {
      return a.dueOn < b.dueOn ? 1 : a.dueOn > b.dueOn ? -1 : 0;
    });
  });

  return {
    configured: true,
    windowDays: RECENTLY_LAUNCHED_WINDOW_DAYS,
    accounts: podConfig.accounts,
    accountDisplayNames: podConfig.accountDisplayNames || {},
    people: podConfig.people,
    byAccount: byAccount
  };
}

/**
 * Called by the Overview tab's media-buyer leaderboard. Two windows:
 *  - "since": June 2026 through today, from the month-level creative sheet
 *  - "matured": leads from 42 to 14 days ago, from the window sheet, so the
 *    slow down-funnel metrics (CPSet, CPD, Revenue, COM) have had time to
 *    play out. Null until the sync has run once with the window enabled.
 * "Creatives" counts the distinct ads that spent inside the same window.
 */
function getOverviewLeaderboard() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('LAST_CREATIVE_SYNC_AT') || !props.getProperty('LAST_SYNC_AT')) {
    return { configured: false };
  }

  var podConfig = getActivePodConfig();
  var briefRows = readDataRows_();
  var reverseClientNameMap = buildReverseClientNameMap_(podConfig);

  var briefByCode = {};
  briefRows.forEach(function (b) {
    var code = String(b.crtvCode || '').trim().toUpperCase();
    if (code && !briefByCode[code]) briefByCode[code] = b;
  });

  function pack(rows, creativeCount) {
    var m = computeCreativeMetrics_(sumCreativeTotals_(rows));
    return {
      creatives: creativeCount,
      cpl: m.cpl,
      cpSet: m.cpSet,
      cpd: m.cpd,
      fbSpend: m.fbSpend,
      revenue: m.revenue,
      com: m.com
    };
  }

  function buildWindow(creativeRows, startDate, endDate) {
    var rowsByPerson = {};
    var creativeKeysByPerson = {};
    var rowsByPersonAccount = {};
    var creativeKeysByPersonAccount = {};
    podConfig.people.forEach(function (p) {
      rowsByPerson[p] = [];
      creativeKeysByPerson[p] = {};
      rowsByPersonAccount[p] = {};
      creativeKeysByPersonAccount[p] = {};
    });

    // "Creatives" = distinct ads (by cleaned ad name, per account) that
    // actually spent in the window, matching how the Top 25 table counts.
    creativeRows.forEach(function (c) {
      var attr = attributeCreativeRow_(c, briefByCode, reverseClientNameMap, podConfig);
      if (!attr) return;
      if (!rowsByPerson[attr.person]) return;
      rowsByPerson[attr.person].push(c);
      (rowsByPersonAccount[attr.person][attr.account] = rowsByPersonAccount[attr.person][attr.account] || []).push(c);
      if (c.spend > 0) {
        creativeKeysByPerson[attr.person][c.adName + '::' + attr.account] = true;
        (creativeKeysByPersonAccount[attr.person][attr.account] = creativeKeysByPersonAccount[attr.person][attr.account] || {})[c.adName] = true;
      }
    });

    return {
      start: startDate,
      end: endDate,
      rows: podConfig.people.map(function (person) {
        var byAccount = {};
        podConfig.accounts.forEach(function (account) {
          byAccount[account] = pack(
            rowsByPersonAccount[person][account] || [],
            Object.keys(creativeKeysByPersonAccount[person][account] || {}).length
          );
        });
        return {
          person: person,
          metrics: pack(rowsByPerson[person], Object.keys(creativeKeysByPerson[person]).length),
          byAccount: byAccount
        };
      })
    };
  }

  var windowStart = props.getProperty('CREATIVE_WINDOW_START');
  var windowEnd = props.getProperty('CREATIVE_WINDOW_END');

  return {
    configured: true,
    people: podConfig.people,
    accounts: podConfig.accounts,
    accountDisplayNames: podConfig.accountDisplayNames || {},
    since: buildWindow(readCreativePerformanceRows_(), CREATIVE_SYNC_START_DATE, formatDayOffset_(0)),
    matured: windowStart && windowEnd
      ? buildWindow(readCreativeWindowRows_(), windowStart, windowEnd)
      : null
  };
}
