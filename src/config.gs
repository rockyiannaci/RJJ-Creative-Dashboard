/**
 * Pod configuration.
 *
 * Each entry fully describes one micro pod: its Asana project, its people,
 * and its accounts. To stand up the dashboard for another pod (B1 or B2),
 * duplicate one entry below, point it at that pod's Asana filter, and set
 * ACTIVE_POD_KEY to its key. No other file needs to change.
 */

var POD_CONFIGS = {
  B3: {
    key: 'B3',
    label: 'B3 (RJJ)',
    dashboardTitle: 'RJJ Creative Dashboard',
    asanaProjectGid: '1202836664104107', // CDM: Creative Requests
    people: [
      'Julian DiVito',
      'Jay Jeong',
      'Rocky Iannaci'
    ],
    accounts: [
      'Refloor',
      'Leaf Home Enhancements',
      'Bath Planet',
      'Renewal By Andersen - GW',
      'Renewal By Andersen - ENY',
      'Renewal By Andersen - QC'
    ],
    // Short labels for the dashboard UI. Filtering still matches against
    // the real Asana "Client Name" values above; this only changes what's
    // displayed on screen.
    accountDisplayNames: {
      'Refloor': 'Refloor',
      'Leaf Home Enhancements': 'Leaf Home',
      'Renewal By Andersen - GW': 'RBA GW',
      'Renewal By Andersen - QC': 'RBA QC',
      'Renewal By Andersen - ENY': 'RBA E NY',
      'Bath Planet': 'Bath Planet'
    },
    // Custom field names on the Asana task, as they appear in Asana.
    clientNameFieldName: 'Client Name',
    creativeTypeFieldName: 'Creative Request Type',
    // The custom field holding the creative asset code (e.g. "CRTV-22716"),
    // used to join a brief back to its ad's performance in BigQuery via
    // the matching code embedded in the ad's name there.
    crtvFieldName: 'CRTV',
    // Fixed display order + colors for the known creative types, matching
    // their Asana tag colors. Any value not listed here (a new type added
    // later, or a blank field) still renders, just in a fallback color at
    // the end of the legend instead of being dropped.
    creativeTypeOrder: ['Iteration', 'Raw Asset', 'Net New Concept'],
    creativeTypeColors: {
      'Iteration': '#8bc34a',
      'Raw Asset': '#b39ddb',
      'Net New Concept': '#ff9800'
    },
    // Floor for the "Avg / Month" and "Avg / Week" benchmarks on the
    // Volume Comparison tab: never average in data from before this
    // month, even though Asana history goes back further.
    benchmarkStartMonth: '2026-07',
    // Minimum creative briefs to submit per account, per calendar month.
    // Drives the Monthly Target Progress section on Overview and the
    // Volume Comparison tab.
    monthlyAccountTargets: {
      'Refloor': 60,
      'Leaf Home Enhancements': 20,
      'Renewal By Andersen - GW': 15,
      'Renewal By Andersen - ENY': 10,
      'Renewal By Andersen - QC': 3,
      'Bath Planet': 5
    },
    // Bonus targets per account: the most reachable tier for each metric
    // (from the Bonus Estimator sheet). Lower is better for all three. `tier`
    // is the share of the bonus that tier pays. Bath Planet is not in the
    // bonus sheet yet; add it here when it is.
    bonusTargets: {
      'Refloor': { cpl: { target: 101, tier: 66 }, cpSet: { target: 429, tier: 33 }, com: { target: 0.19, tier: 66 } },
      'Leaf Home Enhancements': { cpl: { target: 106, tier: 66 }, cpSet: { target: 442, tier: 33 }, com: { target: 0.19, tier: 66 } },
      'Renewal By Andersen - GW': { cpl: { target: 136, tier: 100 }, cpSet: { target: 652, tier: 33 }, com: { target: 0.19, tier: 66 } },
      'Renewal By Andersen - ENY': { cpl: { target: 145, tier: 100 }, cpSet: { target: 610, tier: 33 }, com: { target: 0.19, tier: 66 } },
      'Renewal By Andersen - QC': { cpl: { target: 85, tier: 100 }, cpSet: { target: 449, tier: 33 }, com: { target: 0.19, tier: 66 } }
    },
    // Maps this pod's account keys and people to the exact string values
    // BigQuery's client_name / media_buyer columns use, so ad performance
    // data can be filtered to just this pod without touching the sync or
    // dashboard logic. Client names confirmed against the live table;
    // media buyer names are still a placeholder guess — see the TODO below.
    adPerformanceClientNameMap: {
      'Refloor': 'Refloor Flooring',
      'Leaf Home Enhancements': 'Leaf Home Bath',
      'Renewal By Andersen - GW': 'Renewal by Andersen GW Windows',
      'Renewal By Andersen - ENY': 'Renewal by Andersen of Eastern NY Windows',
      'Renewal By Andersen - QC': 'Renewal by Andersen Quad Cities Windows',
      'Bath Planet': 'Bath Planet Local'
    },
    // TODO: verify these against real distinct media_buyer values once a
    // sync returns rows — adPerformanceSyncLog's unmapped_media_buyers_seen
    // will list any that don't resolve, the same way Asana's SyncLog flags
    // mismatched names.
    adPerformanceMediaBuyerMap: {
      'Julian DiVito': 'Julian DiVito',
      'Jay Jeong': 'Jay Jeong',
      'Rocky Iannaci': 'Rocky Iannaci'
    }
  }
};

/**
 * BigQuery connection details for the agency-wide ad performance data
 * (Looker Studio's source tables). Not pod-specific — every pod's ad
 * performance sync reads from the same project/table, filtered down via
 * that pod's adPerformanceClientNameMap.
 *
 * TODO: confirm dataset/table against the live schema once the service
 * account exists — table_facebook_ads_totals is the README's best guess
 * at the underlying BigQuery table name (LSR's report only names it
 * loosely as "table_facebook-ads-totals"), and podColumn/podValue need
 * checking against real POD values (this pod may show up as "B3",
 * "B3 (RJJ)", or something else entirely).
 *
 * Confirmed against the live schema (cdm-ads.public.table_facebook-ads-totals):
 * date_of_lead, client_id, client_name, currency, product, media_buyer,
 * csm, active, POD, campaign_id/name, ad_set_id/name, ad_id, ad_name,
 * total_spend, clicks, impressions, outbound_clicks, reach,
 * landing_page_views, three_seconds_video_plays, facebook_likes,
 * post_comments, post_shares, sum_revenue, count_set, count_demo,
 * count_sold, count_duplicate, count_leads — table/column names below
 * are all verified; podValue is still a guess.
 */
var AD_PERFORMANCE_CONFIG = {
  projectId: 'cdm-ads',
  dataset: 'public',
  table: 'table_facebook-ads-totals',
  dateColumn: 'date_of_lead',
  clientNameColumn: 'client_name',
  mediaBuyerColumn: 'media_buyer',
  podColumn: 'POD',
  podValue: 'B3 (RJJ)'
};

// Which pod this deployment renders. Change this (or make it a query param
// later) to switch pods without touching sync/dashboard logic.
var ACTIVE_POD_KEY = 'B3';

function getActivePodConfig() {
  var pod = POD_CONFIGS[ACTIVE_POD_KEY];
  if (!pod) {
    throw new Error('Unknown ACTIVE_POD_KEY: ' + ACTIVE_POD_KEY);
  }
  return pod;
}
