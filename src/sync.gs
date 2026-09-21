/**
 * Pulls tasks from Asana, applies the pod's account + person filter, and
 * writes the filtered rows to a Google Sheet in Drive. doGet() (see code.gs)
 * only ever reads that sheet, so page loads never wait on a live Asana call.
 *
 * Run syncAsanaData() on a time-driven trigger (see createHourlyTrigger_)
 * and also on-demand from the dashboard's Refresh button.
 */

var DATA_SHEET_NAME = 'Data';
var LOG_SHEET_NAME = 'SyncLog';

function getOrCreateSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SHEET_ID');

  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch (e) {
      // Fall through and recreate if the stored ID is no longer valid.
    }
  }

  var ss = SpreadsheetApp.create('RJJ Creative Dashboard Data');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getOrCreateSheetTab_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

/**
 * Reads the "Client Name" / "Creative Request" custom field display_value
 * off a task's custom_fields array by field name.
 */
function getCustomFieldValue_(task, fieldName) {
  var fields = task.custom_fields || [];
  for (var i = 0; i < fields.length; i++) {
    if (fields[i].name === fieldName) {
      return fields[i].display_value || '';
    }
  }
  return '';
}

/**
 * Monday of the calendar week containing the given date, as YYYY-MM-DD.
 */
function getWeekStart_(date) {
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  var day = d.getDay(); // 0 = Sunday
  var diffToMonday = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diffToMonday);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Filters raw Asana tasks down to this pod's briefs (account AND person both
 * match) and returns the matched rows plus per-rule exclusion counts, so a
 * mismatched client-name string is easy to spot in the log.
 */
function filterAndAggregate_(tasks, podConfig) {
  var accountSet = {};
  podConfig.accounts.forEach(function (a) {
    accountSet[a.trim().toLowerCase()] = true;
  });
  var peopleSet = {};
  podConfig.people.forEach(function (p) {
    peopleSet[p.trim().toLowerCase()] = true;
  });

  var rows = [];
  var excludedByAccount = 0;
  var excludedByPerson = 0;
  var excludedByBoth = 0;
  var seenAccountValues = {};

  tasks.forEach(function (task) {
    var account = getCustomFieldValue_(task, podConfig.clientNameFieldName);
    var creativeType = getCustomFieldValue_(task, podConfig.creativeTypeFieldName);
    var person = task.created_by && task.created_by.name ? task.created_by.name : '';

    seenAccountValues[account] = (seenAccountValues[account] || 0) + 1;

    var accountMatches = accountSet[account.trim().toLowerCase()] === true;
    var personMatches = peopleSet[person.trim().toLowerCase()] === true;

    if (accountMatches && personMatches) {
      var created = new Date(task.created_at);
      rows.push([
        task.gid,
        task.name,
        task.created_at,
        getWeekStart_(created),
        person,
        account,
        creativeType,
        task.completed
      ]);
    } else if (!accountMatches && !personMatches) {
      excludedByBoth++;
    } else if (!accountMatches) {
      excludedByAccount++;
    } else {
      excludedByPerson++;
    }
  });

  return {
    rows: rows,
    excludedByAccount: excludedByAccount,
    excludedByPerson: excludedByPerson,
    excludedByBoth: excludedByBoth,
    totalTasks: tasks.length,
    seenAccountValues: seenAccountValues
  };
}

function writeDataSheet_(ss, rows) {
  var sheet = getOrCreateSheetTab_(ss, DATA_SHEET_NAME, [
    'gid',
    'name',
    'created_at',
    'week_start',
    'person',
    'account',
    'creative_type',
    'completed'
  ]);

  // Clear everything below the header, then rewrite from scratch. The sheet
  // only ever holds this pod's already-filtered rows, so this is cheap.
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function writeLogSheet_(ss, podConfig, aggregate) {
  var sheet = getOrCreateSheetTab_(ss, LOG_SHEET_NAME, [
    'synced_at',
    'pod',
    'total_tasks_seen',
    'matched',
    'excluded_by_account_only',
    'excluded_by_person_only',
    'excluded_by_both',
    'unmatched_account_values_seen'
  ]);

  var unmatchedAccounts = Object.keys(aggregate.seenAccountValues)
    .filter(function (a) {
      return (
        podConfig.accounts
          .map(function (x) {
            return x.trim().toLowerCase();
          })
          .indexOf(a.trim().toLowerCase()) === -1
      );
    })
    .slice(0, 20)
    .join(', ');

  sheet.appendRow([
    new Date(),
    podConfig.label,
    aggregate.totalTasks,
    aggregate.rows.length,
    aggregate.excludedByAccount,
    aggregate.excludedByPerson,
    aggregate.excludedByBoth,
    unmatchedAccounts
  ]);
}

/**
 * Main sync entrypoint. Called by the time-driven trigger and by the
 * dashboard's Refresh button (via code.gs).
 */
function syncAsanaData() {
  var podConfig = getActivePodConfig();
  var tasks = fetchAllAsanaTasks_(podConfig.asanaProjectGid);
  var aggregate = filterAndAggregate_(tasks, podConfig);

  var ss = getOrCreateSpreadsheet_();
  writeDataSheet_(ss, aggregate.rows);
  writeLogSheet_(ss, podConfig, aggregate);

  PropertiesService.getScriptProperties().setProperty(
    'LAST_SYNC_AT',
    new Date().toISOString()
  );

  return {
    matched: aggregate.rows.length,
    totalTasksSeen: aggregate.totalTasks,
    excludedByAccount: aggregate.excludedByAccount,
    excludedByPerson: aggregate.excludedByPerson,
    excludedByBoth: aggregate.excludedByBoth
  };
}

/**
 * Run once from the script editor to install the automatic hourly sync.
 * Safe to re-run: clears any existing syncAsanaData triggers first so it
 * never double-schedules.
 */
function createHourlyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAsanaData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('syncAsanaData').timeBased().everyHours(1).create();
}
