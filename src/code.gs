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
      day: getDayKeyFromCreatedAt_(r[2]),
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
 * Derives a "yyyy-MM-dd" day key from the created_at cell, for arbitrary
 * (non-calendar-week/month) date-range filtering.
 */
function getDayKeyFromCreatedAt_(createdAt) {
  var d = createdAt instanceof Date ? createdAt : new Date(createdAt);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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

/**
 * Number of calendar months from startMonth ("yyyy-MM") to endMonth,
 * inclusive of both ends.
 */
function monthsBetweenInclusive_(startMonth, endMonth) {
  var s = startMonth.split('-').map(Number);
  var e = endMonth.split('-').map(Number);
  return (e[0] - s[0]) * 12 + (e[1] - s[1]) + 1;
}

/**
 * Number of Mon-Sun weeks from startWeekKey to endWeekKey (both
 * "yyyy-MM-dd" Monday keys), inclusive of both ends.
 */
function weeksBetweenInclusive_(startWeekKey, endWeekKey) {
  var start = new Date(startWeekKey + 'T00:00:00');
  var end = new Date(endWeekKey + 'T00:00:00');
  return Math.round((end - start) / (7 * 86400000)) + 1;
}

/**
 * Per-account average pace (briefs / month, briefs / week), based on the
 * full history rather than any trailing window or filter — a stable
 * benchmark for "how much should we be briefing" regardless of which
 * period is selected in the Performance Comparison tab.
 */
function computeAccountBenchmarks_(rows, podConfig) {
  var benchmarks = {};

  if (rows.length === 0) {
    podConfig.accounts.forEach(function (account) {
      benchmarks[account] = { avgPerMonth: 0, avgPerWeek: 0 };
    });
    return { benchmarks: benchmarks, monthsSpanned: 0, weeksSpanned: 0 };
  }

  var minMonth = rows[0].monthStart;
  var minWeek = rows[0].weekStart;
  rows.forEach(function (r) {
    if (r.monthStart < minMonth) minMonth = r.monthStart;
    if (r.weekStart < minWeek) minWeek = r.weekStart;
  });

  var todayMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var todayWeek = getWeekStart_(new Date());
  var monthsSpanned = monthsBetweenInclusive_(minMonth, todayMonth);
  var weeksSpanned = weeksBetweenInclusive_(minWeek, todayWeek);

  podConfig.accounts.forEach(function (account) {
    var total = rows.filter(function (r) {
      return r.account === account;
    }).length;
    benchmarks[account] = {
      avgPerMonth: monthsSpanned ? total / monthsSpanned : 0,
      avgPerWeek: weeksSpanned ? total / weeksSpanned : 0
    };
  });

  return { benchmarks: benchmarks, monthsSpanned: monthsSpanned, weeksSpanned: weeksSpanned };
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

  // Fixed creative-type display order, derived from the full history (not
  // just the trailing window) so it stays stable across every comparison
  // period, including "All Time".
  var creativeTypesSeenAll = {};
  rows.forEach(function (r) {
    creativeTypesSeenAll[r.creativeType || 'Unspecified'] = true;
  });
  var creativeTypeOrder = (podConfig.creativeTypeOrder || []).filter(function (t) {
    return creativeTypesSeenAll[t];
  });
  Object.keys(creativeTypesSeenAll)
    .sort()
    .forEach(function (t) {
      if (creativeTypeOrder.indexOf(t) === -1) {
        creativeTypeOrder.push(t);
      }
    });

  function buildAccountPersonMatrix_(rowsSubset) {
    var matrix = {};
    podConfig.accounts.forEach(function (account) {
      var byPerson = {};
      podConfig.people.forEach(function (person) {
        byPerson[person] = rowsSubset.filter(function (r) {
          return r.account === account && r.person === person;
        }).length;
      });
      matrix[account] = byPerson;
    });
    return matrix;
  }

  function buildTypePersonMatrix_(rowsSubset) {
    var matrix = {};
    creativeTypeOrder.forEach(function (type) {
      var byPerson = {};
      podConfig.people.forEach(function (person) {
        byPerson[person] = rowsSubset.filter(function (r) {
          return (r.creativeType || 'Unspecified') === type && r.person === person;
        }).length;
      });
      matrix[type] = byPerson;
    });
    return matrix;
  }

  // Cross-media-buyer comparison: how much did each person produce per
  // account, and per creative type, side by side (rather than one person
  // at a time as in the per-person detail tabs). Computed once for "all"
  // (the full history) and once per trailing month, so the Performance
  // Comparison tab can filter between All Time and any given month.
  var comparisonPeriods = ['all'].concat(months);
  var comparisonByAccountByPeriod = {};
  var comparisonByTypeByPeriod = {};
  comparisonPeriods.forEach(function (period) {
    var periodRows = period === 'all'
      ? rows
      : rows.filter(function (r) { return r.monthStart === period; });
    comparisonByAccountByPeriod[period] = buildAccountPersonMatrix_(periodRows);
    comparisonByTypeByPeriod[period] = buildTypePersonMatrix_(periodRows);
  });

  // Also kept as the trailing-window-only versions, used by the per-person
  // "Performance vs. Team" tables on the Media Buyer Detail tab.
  var comparisonByAccount = buildAccountPersonMatrix_(trailingRows);
  var comparisonByType = buildTypePersonMatrix_(trailingRows);

  var accountBenchmarkResult = computeAccountBenchmarks_(rows, podConfig);

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
    comparisonPeriods: comparisonPeriods,
    comparisonByAccountByPeriod: comparisonByAccountByPeriod,
    accountBenchmarks: accountBenchmarkResult.benchmarks,
    benchmarkMonthsSpanned: accountBenchmarkResult.monthsSpanned,
    benchmarkWeeksSpanned: accountBenchmarkResult.weeksSpanned,
    comparisonByTypeByPeriod: comparisonByTypeByPeriod,
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

// ---------- Custom-period overview (month picker / All Time / custom range) ----------

function ymd_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function addDays_(dateKey, days) {
  var d = new Date(dateKey + 'T00:00:00');
  d.setDate(d.getDate() + days);
  return ymd_(d);
}

function daysBetween_(startKey, endKey) {
  var start = new Date(startKey + 'T00:00:00');
  var end = new Date(endKey + 'T00:00:00');
  return Math.round((end - start) / 86400000) + 1;
}

function lastDayOfMonth_(monthKey) {
  var parts = monthKey.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  return ymd_(new Date(y, m, 0)); // day 0 of next month = last day of this one
}

function monthLabel_(monthKey) {
  var names = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
    'August', 'September', 'October', 'November', 'December'];
  var parts = monthKey.split('-');
  return names[parseInt(parts[1], 10) - 1] + ' ' + parts[0];
}

/**
 * Every Monday week-start key whose Mon-Sun week overlaps [startKey, endKey].
 */
function weeksOverlapping_(startKey, endKey) {
  var weeks = [];
  var cursor = getWeekStart_(new Date(startKey + 'T00:00:00'));
  var endWeek = getWeekStart_(new Date(endKey + 'T00:00:00'));
  while (cursor <= endWeek) {
    weeks.push(cursor);
    cursor = addDays_(cursor, 7);
  }
  return weeks;
}

/**
 * Splits [startKey, endKey] into consecutive 7-day chunks (last one may be
 * shorter). Used for custom ranges, which aren't calendar-aligned.
 */
function chunksOfSevenDays_(startKey, endKey) {
  var chunks = [];
  var cursor = startKey;
  while (cursor <= endKey) {
    var chunkEnd = addDays_(cursor, 6);
    if (chunkEnd > endKey) chunkEnd = endKey;
    chunks.push({ startKey: cursor, endKey: chunkEnd });
    cursor = addDays_(chunkEnd, 1);
  }
  return chunks;
}

/**
 * Resolves a filter spec from the dashboard's overview filter bar into a
 * concrete date range, its immediately-preceding comparison period (for
 * the delta on the stat cards), and how to bucket it for the trend chart.
 *
 * filterSpec is one of:
 *   { type: 'all' }
 *   { type: 'month', month: 'yyyy-MM' }
 *   { type: 'custom', startDaysAgo: N, endDaysAgo: M }  (both counted back from today)
 */
function resolveFilterRange_(filterSpec, rows, todayKey) {
  if (filterSpec.type === 'all') {
    var minDay = todayKey;
    rows.forEach(function (r) {
      if (r.day < minDay) minDay = r.day;
    });
    var startMonth = minDay.substring(0, 7);
    var endMonth = todayKey.substring(0, 7);
    var months = [];
    var y = parseInt(startMonth.split('-')[0], 10);
    var m = parseInt(startMonth.split('-')[1], 10);
    var endY = parseInt(endMonth.split('-')[0], 10);
    var endM = parseInt(endMonth.split('-')[1], 10);
    while (y < endY || (y === endY && m <= endM)) {
      months.push(y + '-' + (m < 10 ? '0' : '') + m);
      m++;
      if (m > 12) {
        m = 1;
        y++;
      }
    }
    return {
      startKey: minDay,
      endKey: todayKey,
      prevStartKey: null,
      prevEndKey: null,
      granularity: 'month',
      bucketKeys: months,
      rangeLabel: 'All Time'
    };
  }

  if (filterSpec.type === 'month') {
    var month = filterSpec.month;
    var startKey = month + '-01';
    var endKey = lastDayOfMonth_(month);
    if (endKey > todayKey) endKey = todayKey;

    var prevDate = new Date(startKey + 'T00:00:00');
    prevDate.setMonth(prevDate.getMonth() - 1);
    var prevMonth = Utilities.formatDate(prevDate, Session.getScriptTimeZone(), 'yyyy-MM');

    return {
      startKey: startKey,
      endKey: endKey,
      prevStartKey: prevMonth + '-01',
      prevEndKey: lastDayOfMonth_(prevMonth),
      granularity: 'week',
      bucketKeys: weeksOverlapping_(startKey, endKey),
      rangeLabel: monthLabel_(month)
    };
  }

  // 'custom': both values are "days ago", larger one is the start of the range.
  var startDaysAgo = Math.max(filterSpec.startDaysAgo, filterSpec.endDaysAgo);
  var endDaysAgo = Math.min(filterSpec.startDaysAgo, filterSpec.endDaysAgo);
  var rangeEndKey = addDays_(todayKey, -endDaysAgo);
  var rangeStartKey = addDays_(todayKey, -startDaysAgo);
  var lengthDays = daysBetween_(rangeStartKey, rangeEndKey);
  var prevEndKey = addDays_(rangeStartKey, -1);
  var prevStartKey = addDays_(prevEndKey, -(lengthDays - 1));

  return {
    startKey: rangeStartKey,
    endKey: rangeEndKey,
    prevStartKey: prevStartKey,
    prevEndKey: prevEndKey,
    granularity: 'chunk',
    bucketKeys: chunksOfSevenDays_(rangeStartKey, rangeEndKey),
    rangeLabel: startDaysAgo + '–' + endDaysAgo + ' days ago'
  };
}

/**
 * Powers the Overview tab's filter bar (month picker / All Time / custom
 * date range). Reads the same pre-synced sheet as getDashboardData(), just
 * bucketed and totaled differently, so it's just as fast.
 */
function getOverviewForFilter(filterSpec) {
  var podConfig = getActivePodConfig();
  var rows = readDataRows_();
  var todayKey = ymd_(new Date());

  var range = resolveFilterRange_(filterSpec, rows, todayKey);

  var periodRows = rows.filter(function (r) {
    return r.day >= range.startKey && r.day <= range.endKey;
  });
  var prevPeriodRows = range.prevStartKey
    ? rows.filter(function (r) {
        return r.day >= range.prevStartKey && r.day <= range.prevEndKey;
      })
    : [];

  var totals = {};
  var previousTotals = {};
  podConfig.people.forEach(function (person) {
    totals[person] = periodRows.filter(function (r) {
      return r.person === person;
    }).length;
    previousTotals[person] = prevPeriodRows.filter(function (r) {
      return r.person === person;
    }).length;
  });

  var creativeTypesSeenAll = {};
  rows.forEach(function (r) {
    creativeTypesSeenAll[r.creativeType || 'Unspecified'] = true;
  });
  var creativeTypeOrder = (podConfig.creativeTypeOrder || []).filter(function (t) {
    return creativeTypesSeenAll[t];
  });
  Object.keys(creativeTypesSeenAll)
    .sort()
    .forEach(function (t) {
      if (creativeTypeOrder.indexOf(t) === -1) creativeTypeOrder.push(t);
    });

  var trend = {};
  podConfig.people.forEach(function (person) {
    trend[person] = range.bucketKeys.map(function (bucket) {
      var bStart, bEnd;
      if (range.granularity === 'chunk') {
        bStart = bucket.startKey;
        bEnd = bucket.endKey;
      } else if (range.granularity === 'month') {
        bStart = bucket + '-01';
        bEnd = lastDayOfMonth_(bucket);
      } else {
        bStart = bucket;
        bEnd = addDays_(bucket, 6);
      }
      return periodRows.filter(function (r) {
        return r.person === person && r.day >= bStart && r.day <= bEnd;
      }).length;
    });
  });

  return {
    rangeLabel: range.rangeLabel,
    hasPrevious: !!range.prevStartKey,
    granularity: range.granularity,
    bucketKeys: range.bucketKeys,
    people: podConfig.people,
    accounts: podConfig.accounts,
    accountDisplayNames: podConfig.accountDisplayNames || {},
    creativeTypeColors: podConfig.creativeTypeColors || {},
    creativeTypeOrder: creativeTypeOrder,
    totals: totals,
    previousTotals: previousTotals,
    trend: trend,
    accountBreakdown: tallyBy_(periodRows, function (r) {
      return r.account;
    }),
    creativeTypeBreakdown: tallyBy_(periodRows, function (r) {
      return r.creativeType || 'Unspecified';
    })
  };
}

/**
 * Called from the dashboard's Refresh button via google.script.run.
 */
function refreshData() {
  return syncAsanaData();
}
