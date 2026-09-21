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
    asanaProjectGid: '1202836664104107', // CDM: Creative Requests
    people: [
      'Rocky Iannaci',
      'Julian DiVito',
      'Jay Jeong'
    ],
    accounts: [
      'Refloor',
      'Leaf Home Enhancements',
      'Renewal By Andersen - GW',
      'Renewal By Andersen - QC',
      'Renewal By Andersen - ENY'
    ],
    // Custom field names on the Asana task, as they appear in Asana.
    clientNameFieldName: 'Client Name',
    creativeTypeFieldName: 'Creative Request'
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
