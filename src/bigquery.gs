/**
 * BigQuery access, in two modes:
 *
 * 1. As the script's own Google identity (default) — via the built-in
 *    BigQuery Advanced Service (see appsscript.json's enabledAdvancedServices).
 *    Works as long as whoever deployed this web app has BigQuery Data
 *    Viewer + Job User on the target project themselves. The first time
 *    it runs, that person needs to run any function manually from the
 *    Apps Script editor once to grant the added BigQuery scope — a web
 *    app or trigger execution can't show that consent screen itself.
 * 2. As a service account (opt-in) — if BIGQUERY_SERVICE_ACCOUNT_KEY is
 *    set in Script Properties, that's used instead, independent of who
 *    owns this Apps Script project. Useful if the deploying user's own
 *    account shouldn't have standing BigQuery access. This signs its own
 *    JWT and exchanges it for an OAuth token (the "JWT Bearer" flow) —
 *    the service-account key JSON goes in Script Properties directly,
 *    never into code.
 */

var BIGQUERY_TOKEN_URL = 'https://oauth2.googleapis.com/token';
var BIGQUERY_SCOPE = 'https://www.googleapis.com/auth/bigquery.readonly';

function hasBigQueryServiceAccountKey_() {
  return !!PropertiesService.getScriptProperties().getProperty('BIGQUERY_SERVICE_ACCOUNT_KEY');
}

function base64UrlEncode_(input) {
  var bytes = typeof input === 'string' ? Utilities.newBlob(input).getBytes() : input;
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

/**
 * Signs a JWT with the service account's private key and exchanges it for
 * a short-lived BigQuery access token. Not cached across executions
 * (Apps Script executions are short-lived anyway), but cheap enough to
 * redo on every sync run.
 */
function getBigQueryAccessToken_() {
  var keyJson = PropertiesService.getScriptProperties().getProperty('BIGQUERY_SERVICE_ACCOUNT_KEY');
  if (!keyJson) {
    throw new Error(
      'BIGQUERY_SERVICE_ACCOUNT_KEY is not set in Script Properties. Paste the service ' +
        "account's JSON key there (Project Settings > Script Properties) to enable ad performance sync."
    );
  }

  var creds;
  try {
    creds = JSON.parse(keyJson);
  } catch (e) {
    throw new Error('BIGQUERY_SERVICE_ACCOUNT_KEY is not valid JSON: ' + e.message);
  }

  var now = Math.floor(Date.now() / 1000);
  var header = { alg: 'RS256', typ: 'JWT' };
  var claimSet = {
    iss: creds.client_email,
    scope: BIGQUERY_SCOPE,
    aud: BIGQUERY_TOKEN_URL,
    exp: now + 3600,
    iat: now
  };

  var unsigned = base64UrlEncode_(JSON.stringify(header)) + '.' + base64UrlEncode_(JSON.stringify(claimSet));
  var signatureBytes = Utilities.computeRsaSha256Signature(unsigned, creds.private_key);
  var jwt = unsigned + '.' + base64UrlEncode_(signatureBytes);

  var response = UrlFetchApp.fetch(BIGQUERY_TOKEN_URL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('BigQuery auth failed (' + code + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText()).access_token;
}

/**
 * Runs a SQL query against BigQuery and returns rows as plain objects
 * keyed by column name. Uses the service account (JWT) path if a key is
 * configured, otherwise runs as the script's own Google identity via the
 * built-in BigQuery Advanced Service. Both return the same
 * { schema, rows: [{f: [{v}, ...]}] } shape, parsed identically.
 */
function runBigQueryQuery_(projectId, sql) {
  var body = hasBigQueryServiceAccountKey_()
    ? runBigQueryQueryViaServiceAccount_(projectId, sql)
    : runBigQueryQueryAsSelf_(projectId, sql);

  if (!body.jobComplete) {
    throw new Error('BigQuery query did not complete within the timeout. Try narrowing the date range.');
  }

  var fieldNames = (body.schema && body.schema.fields ? body.schema.fields : []).map(function (f) {
    return f.name;
  });

  return (body.rows || []).map(function (row) {
    var obj = {};
    row.f.forEach(function (cell, i) {
      obj[fieldNames[i]] = cell.v;
    });
    return obj;
  });
}

function runBigQueryQueryViaServiceAccount_(projectId, sql) {
  var token = getBigQueryAccessToken_();
  var url = 'https://bigquery.googleapis.com/bigquery/v2/projects/' + encodeURIComponent(projectId) + '/queries';

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      query: sql,
      useLegacySql: false,
      timeoutMs: 30000
    }),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('BigQuery query failed (' + code + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

/**
 * Runs as the script's own Google identity via the BigQuery Advanced
 * Service. Requires that identity to have BigQuery Data Viewer + Job
 * User on `projectId` — an IAM check Google makes, not something this
 * script can grant itself.
 */
function runBigQueryQueryAsSelf_(projectId, sql) {
  var request = { query: sql, useLegacySql: false, timeoutMs: 30000 };
  return BigQuery.Jobs.query(request, projectId);
}

/**
 * Escapes a single value for inline use in a BigQuery string literal.
 * Only used with values from our own config (account/person names), never
 * with external user input, so this is a correctness guard against stray
 * quotes/apostrophes in a name, not an injection defense.
 */
function bqStringLiteral_(value) {
  return "'" + String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
}
