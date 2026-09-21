/**
 * Web app entrypoint and the read-only aggregation the dashboard renders.
 * Never calls Asana directly — only reads the sheet that syncAsanaData()
 * (sync.gs) already populated, so page loads stay fast.
 */

var TRAILING_WEEKS = 10;
var TRAILING_MONTHS = 6;

function doGet() {
  var podConfig = getActivePodConfig();
  var template = HtmlService.createTemplateFromFile('index');
  template.podLabel = podConfig.label;
  template.dashboardTitle = podConfig.dashboardTitle || podConfig.label;
  return template
    .evaluate()
    .setTitle(template.dashboardTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function readDataRows_() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SHEET_ID');
  if (!sheetId) {
    return [];
  }
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(DATA_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 8)
    .getValues();

  return values.map(function (r) {
    return {
      gid: r[0],
      name: r[1],
      createdAt: r[2],
      weekStart: normalizeDateCell_(r[3]),
      monthStart: getMonthKeyFromCreatedAt_(r[2]),
      person: r[4],
      account: r[5],
      creativeType: r[6],
      completed: r[7]
    };
  });
}

/**
 * Derives a "yyyy-MM" month key from the created_at cell, which (like
 * week_start) may come back as either a string or an auto-converted Date.
 */
function getMonthKeyFromCreatedAt_(createdAt) {
  var d = createdAt instanceof Date ? createdAt : new Date(createdAt);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
}

/**
 * Ordered list of the last N month keys ("yyyy-MM"), oldest first, ending
 * with the current month.
 */
function getLastNMonths_(n) {
  var months = [];
  var today = new Date();
  for (var i = n - 1; i >= 0; i--) {
    var d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    months.push(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM'));
  }
  return months;
}

/**
 * Sheets auto-converts "YYYY-MM-DD" strings into real Date cells, so a cell
 * read back with getValues() can come back as either a string or a Date
 * depending on what Sheets decided at write time. Normalize both to the
 * same "yyyy-MM-dd" string so week comparisons below are reliable.
 */
function normalizeDateCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value;
}

/**
 * Ordered list of the last N Monday-week-start strings (YYYY-MM-DD),
 * oldest first, ending with the current week.
 */
function getLastNWeekStarts_(n) {
  var weeks = [];
  var today = new Date();
  var currentWeekStart = getWeekStart_(today);
  for (var i = n - 1; i >= 0; i--) {
    var d = new Date(currentWeekStart + 'T00:00:00');
    d.setDate(d.getDate() - i * 7);
    weeks.push(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  }
  return weeks;
}

function tallyBy_(rows, keyFn) {
  var result = {};
  rows.forEach(function (row) {
    var key = keyFn(row);
    result[key] = (result[key] || 0) + 1;
  });
  return result;
}

/**
 * Builds the full dataset the dashboard's charts render, from the already
 * filtered rows in the Data sheet.
 */
function getDashboardData() {
  var podConfig = getActivePodConfig();
  var rows = readDataRows_();
  var weeks = getLastNWeekStarts_(TRAILING_WEEKS);
  var weekSet = {};
  weeks.forEach(function (w) {
    weekSet[w] = true;
  });

  var currentWeek = weeks[weeks.length - 1];
  var lastWeek = weeks[weeks.length - 2];

  var trailingRows = rows.filter(function (r) {
    return weekSet[r.weekStart];
  });
  var currentWeekRows = rows.filter(function (r) {
    return r.weekStart === currentWeek;
  });
  var lastWeekRows = rows.filter(function (r) {
    return r.weekStart === lastWeek;
  });

  // Weekly trend: one series per person, one count per week.
  var weeklyTrend = {};
  podConfig.people.forEach(function (person) {
    weeklyTrend[person] = weeks.map(function (week) {
      return rows.filter(function (r) {
        return r.person === person && r.weekStart === week;
      }).length;
    });
  });

  // This week vs last week, per person.
  var thisWeek = {};
  var lastWeekCounts = {};
  podConfig.people.forEach(function (person) {
    thisWeek[person] = currentWeekRows.filter(function (r) {
      return r.person === person;
    }).length;
    lastWeekCounts[person] = lastWeekRows.filter(function (r) {
      return r.person === person;
    }).length;
  });

  // Per-week account x person matrix, for the weekly-breakdown tabs:
  // for a given week, how many briefs did each person submit for each
  // account, so a mismatch between plan and output per account is visible.
  var weeklyMatrix = {};
  weeks.forEach(function (week) {
    var weekRows = rows.filter(function (r) {
      return r.weekStart === week;
    });
    var matrix = {};
    podConfig.accounts.forEach(function (account) {
      var byPerson = {};
      podConfig.people.forEach(function (person) {
        byPerson[person] = weekRows.filter(function (r) {
          return r.account === account && r.person === person;
        }).length;
      });
      matrix[account] = byPerson;
    });
    weeklyMatrix[week] = matrix;
  });

  // Per-month account x person matrix, same shape as weeklyMatrix but
  // bucketed by calendar month, for the "big picture" monthly view.
  var months = getLastNMonths_(TRAILING_MONTHS);
  var monthlyMatrix = {};
  months.forEach(function (month) {
    var monthRows = rows.filter(function (r) {
      return r.monthStart === month;
    });
    var matrix = {};
    podConfig.accounts.forEach(function (account) {
      var byPerson = {};
      podConfig.people.forEach(function (person) {
        byPerson[person] = monthRows.filter(function (r) {
          return r.account === account && r.person === person;
        }).length;
      });
      matrix[account] = byPerson;
    });
    monthlyMatrix[month] = matrix;
  });

  // Monthly trend: one series per person, one count per month, for the
  // "total briefs per media buyer per month" bar chart.
  var monthlyTrend = {};
  podConfig.people.forEach(function (person) {
    monthlyTrend[person] = months.map(function (month) {
      return rows.filter(function (r) {
        return r.person === person && r.monthStart === month;
      }).length;
    });
  });

  // Cross-media-buyer comparison, over the trailing window: how much did
  // each person produce per account, and per creative type, side by side
  // (rather than one person at a time as in the per-person detail tabs).
  var comparisonByAccount = {};
  podConfig.accounts.forEach(function (account) {
    var byPerson = {};
    podConfig.people.forEach(function (person) {
      byPerson[person] = trailingRows.filter(function (r) {
        return r.account === account && r.person === person;
      }).length;
    });
    comparisonByAccount[account] = byPerson;
  });

  var creativeTypesSeen = {};
  trailingRows.forEach(function (r) {
    creativeTypesSeen[r.creativeType || 'Unspecified'] = true;
  });
  var creativeTypeOrder = (podConfig.creativeTypeOrder || []).filter(function (t) {
    return creativeTypesSeen[t];
  });
  Object.keys(creativeTypesSeen)
    .sort()
    .forEach(function (t) {
      if (creativeTypeOrder.indexOf(t) === -1) {
        creativeTypeOrder.push(t);
      }
    });

  var comparisonByType = {};
  creativeTypeOrder.forEach(function (type) {
    var byPerson = {};
    podConfig.people.forEach(function (person) {
      byPerson[person] = trailingRows.filter(function (r) {
        return (r.creativeType || 'Unspecified') === type && r.person === person;
      }).length;
    });
    comparisonByType[type] = byPerson;
  });

  // Per-person detail: weekly volume, and account / creative-type
  // breakdowns over the trailing window, for the person-detail tabs.
  var perPerson = {};
  podConfig.people.forEach(function (person) {
    var personTrailingRows = trailingRows.filter(function (r) {
      return r.person === person;
    });
    perPerson[person] = {
      weeklyCounts: weeklyTrend[person],
      totalTrailing: personTrailingRows.length,
      accountBreakdown: tallyBy_(personTrailingRows, function (r) {
        return r.account;
      }),
      creativeTypeBreakdown: tallyBy_(personTrailingRows, function (r) {
        return r.creativeType || 'Unspecified';
      })
    };
  });

  return {
    podLabel: podConfig.label,
    dashboardTitle: podConfig.dashboardTitle || podConfig.label,
    people: podConfig.people,
    accounts: podConfig.accounts,
    accountDisplayNames: podConfig.accountDisplayNames || {},
    creativeTypeColors: podConfig.creativeTypeColors || {},
    weeks: weeks,
    months: months,
    currentWeek: currentWeek,
    lastWeek: lastWeek,
    lastSyncAt: PropertiesService.getScriptProperties().getProperty('LAST_SYNC_AT') || null,
    weeklyTrend: weeklyTrend,
    monthlyTrend: monthlyTrend,
    thisWeekByPerson: thisWeek,
    lastWeekByPerson: lastWeekCounts,
    perPerson: perPerson,
    weeklyMatrix: weeklyMatrix,
    monthlyMatrix: monthlyMatrix,
    comparisonByAccount: comparisonByAccount,
    comparisonByType: comparisonByType,
    creativeTypeOrder: creativeTypeOrder,
    accountBreakdown: {
      currentWeek: tallyBy_(currentWeekRows, function (r) {
        return r.account;
      }),
      trailing: tallyBy_(trailingRows, function (r) {
        return r.account;
      })
    },
    creativeTypeBreakdown: {
      currentWeek: tallyBy_(currentWeekRows, function (r) {
        return r.creativeType || 'Unspecified';
      }),
      trailing: tallyBy_(trailingRows, function (r) {
        return r.creativeType || 'Unspecified';
      })
    }
  };
}

/**
 * Called from the dashboard's Refresh button via google.script.run.
 */
function refreshData() {
  return syncAsanaData();
}
