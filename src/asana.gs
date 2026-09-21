/**
 * Thin wrapper around the Asana REST API. No filtering or aggregation here —
 * that lives in sync.gs. Keeping this file dumb makes it reusable across pods.
 */

var ASANA_API_BASE = 'https://app.asana.com/api/1.0';
var ASANA_TASK_OPT_FIELDS =
  'name,created_at,created_by.name,custom_fields.name,custom_fields.display_value,completed';

function getAsanaToken_() {
  var token = PropertiesService.getScriptProperties().getProperty('ASANA_TOKEN');
  if (!token) {
    throw new Error(
      'ASANA_TOKEN is not set. Run setAsanaToken_() once from the script editor, ' +
        'or set it via Project Settings > Script Properties.'
    );
  }
  return token;
}

/**
 * One-time helper: run this manually from the Apps Script editor to store
 * the token, then delete the literal value from source. Never commit a
 * real token to this file.
 */
function setAsanaToken_() {
  var token = 'PASTE_TOKEN_TEMPORARILY_THEN_RUN_ONCE_THEN_REMOVE';
  PropertiesService.getScriptProperties().setProperty('ASANA_TOKEN', token);
}

/**
 * Fetches every task in the given project, following pagination until
 * exhausted. Returns the full history (no modified_since filter) since we
 * need created_at to bucket into calendar weeks and Asana never deletes
 * tasks.
 */
function fetchAllAsanaTasks_(projectGid) {
  var token = getAsanaToken_();
  var tasks = [];
  var url =
    ASANA_API_BASE +
    '/projects/' +
    projectGid +
    '/tasks?opt_fields=' +
    encodeURIComponent(ASANA_TASK_OPT_FIELDS) +
    '&limit=100';

  while (url) {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code !== 200) {
      throw new Error('Asana API error ' + code + ': ' + response.getContentText());
    }

    var body = JSON.parse(response.getContentText());
    tasks = tasks.concat(body.data || []);

    url = body.next_page && body.next_page.uri ? body.next_page.uri : null;
  }

  return tasks;
}
