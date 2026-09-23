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
    dashboardTitle: 'RJJ Creative Brief Dashboard',
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
    // Performance Comparison tab: never average in data from before this
    // month, even though Asana history goes back further.
    benchmarkStartMonth: '2026-07',
    // Minimum creative briefs to submit per account, per calendar month.
    // Drives the Monthly Target Progress section on Overview and the
    // Performance Comparison tab.
    monthlyAccountTargets: {
      'Refloor': 60,
      'Leaf Home Enhancements': 20,
      'Renewal By Andersen - GW': 15,
      'Renewal By Andersen - ENY': 10,
      'Renewal By Andersen - QC': 3,
      'Bath Planet': 5
    }
  }
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
