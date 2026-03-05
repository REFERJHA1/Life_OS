const SHEET_NAMES = {
  work: 'WorkTasks',
  personal: 'PersonalTasks',
  habits: 'Habits',
  journal: 'Journal',
  review: 'DailyReview',
  timetable: 'TimeTable',
  backup: 'Backup'
};

const HABIT_COLUMNS = ['Study', 'Reading', 'Workout', 'Planning'];

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  return routeAction_(action, e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  let payload = {};

  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (error) {
    payload = {};
  }

  if (!payload.action && e && e.parameter && e.parameter.action) {
    payload.action = e.parameter.action;
  }

  return routeAction_(payload.action || '', payload);
}

function routeAction_(action, payload) {
  try {
    switch (action) {
      case 'getDashboardData':
        return jsonResponse_(getDashboardData_(payload));
      case 'getTodayTasks':
        return jsonResponse_(getTodayTasks_(payload));
      case 'getWorkTasks':
        return jsonResponse_(getWorkTasks_(payload));
      case 'getPersonalTasks':
        return jsonResponse_(getPersonalTasks_(payload));
      case 'getJournalEntries':
        return jsonResponse_(getJournalEntries_(payload));
      case 'getJournalHistory':
        return jsonResponse_(getJournalHistory_(payload));
      case 'getAnalytics':
        return jsonResponse_(getAnalytics_(payload));
      case 'getTimeTable':
        return jsonResponse_(getTimeTable_(payload));
      case 'dailyCheckin':
        return jsonResponse_(saveDailyCheckin_(payload));
      case 'addWorkTask':
        return jsonResponse_(addWorkTask_(payload.task || payload));
      case 'addPersonalTask':
        return jsonResponse_(addPersonalTask_(payload.task || payload));
      case 'saveHabit':
        return jsonResponse_(saveHabit_(payload));
      case 'saveJournal':
        return jsonResponse_(saveJournal_(payload));
      case 'saveDailyReview':
        return jsonResponse_(saveDailyReview_(payload));
      case 'createWeeklyBackup':
        return jsonResponse_(createWeeklyBackup_(payload));
      case 'addTimeTableEntry':
        return jsonResponse_(addTimeTableEntry_(payload.entry || payload));
      case 'completeWorkTask':
        return jsonResponse_(completeWorkTask_(payload));
      case 'completePersonalTask':
        return jsonResponse_(completePersonalTask_(payload));
      case 'completeTimeTableEntry':
        return jsonResponse_(completeTimeTableEntry_(payload));
      default:
        return jsonResponse_({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || 'Unexpected server error' });
  }
}

function getDashboardData_(payload) {
  const selectedDate = payload && (payload.date || payload.Date);
  const referenceDate = normalizeDate_(selectedDate) || todayString_();
  const cache = CacheService.getScriptCache();
  const cacheKey = 'dashboard:' + getDataVersion_() + ':' + referenceDate;
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const today = referenceDate;
  const rowsBySheet = loadRowsBySheet_();
  const workRows = rowsBySheet.work;
  const personalRows = rowsBySheet.personal;
  const habitRows = rowsBySheet.habits;
  const journalRows = rowsBySheet.journal;

  const todayTasks = getTodayTasksFromRows_(today, workRows, personalRows);
  const workAllToday = getWorkTasksByDate_(today, false, workRows);
  const personalAllToday = getPersonalTasksByDate_(today, false, personalRows);
  const habitsToday = getHabitsByDate_(today, habitRows);
  const journalToday = getJournalByDate_(today, journalRows);

  const workDone = countDone_(workAllToday, 'Status');
  const personalDone = countDone_(personalAllToday, 'Status');
  const habitsDone = HABIT_COLUMNS.reduce(function (sum, name) {
    return sum + (Number(habitsToday[name]) === 1 ? 1 : 0);
  }, 0);

  const score = calculateDailyScore_(
    workDone,
    workAllToday.length,
    personalDone,
    personalAllToday.length,
    habitsDone,
    HABIT_COLUMNS.length
  );

  const yesterday = dateStringOffsetFrom_(today, 1);
  const yesterdayScore = getDailyScoreByDate_(yesterday, workRows, personalRows, habitRows);

  const result = {
    ok: true,
    date: today,
    score: score,
    moodToday: Number(journalToday.Mood || 0),
    habitsDone: habitsDone,
    habits: {
      Study: Number(habitsToday.Study || 0),
      Reading: Number(habitsToday.Reading || 0),
      Workout: Number(habitsToday.Workout || 0),
      Planning: Number(habitsToday.Planning || 0)
    },
    workTasks: todayTasks.workTasks,
    personalTasks: todayTasks.personalTasks,
    pendingTasks: todayTasks.workTasks.length + todayTasks.personalTasks.length,
    weeklyScores: getWeeklyScores_(workRows, personalRows, habitRows, today),
    motivation: getMotivation_(score),
    reward: habitsDone === HABIT_COLUMNS.length ? 'Perfect habit day' : '',
    reminder: yesterdayScore === 0 ? 'You missed yesterday' : '',
    streaks: {
      Study: getHabitStreak_('Study', today, habitRows),
      Reading: getHabitStreak_('Reading', today, habitRows),
      Workout: getHabitStreak_('Workout', today, habitRows),
      Planning: getHabitStreak_('Planning', today, habitRows),
      Journal: getJournalStreak_(today, journalRows)
    }
  };

  cache.put(cacheKey, JSON.stringify(result), 60);
  return result;
}

function getTodayTasks_(payload) {
  const selectedDate = payload && (payload.date || payload.Date);
  const today = normalizeDate_(selectedDate) || todayString_();
  const rowsBySheet = loadRowsBySheet_();
  return getTodayTasksFromRows_(today, rowsBySheet.work, rowsBySheet.personal);
}

function getTodayTasksFromRows_(dateText, workRows, personalRows) {
  const work = getWorkTasksByDate_(dateText, true, workRows);
  const personal = getPersonalTasksByDate_(dateText, true, personalRows);

  return {
    ok: true,
    date: dateText,
    workTasks: work.map(function (row) {
      return row.Task;
    }),
    personalTasks: personal.map(function (row) {
      return row.Task;
    })
  };
}

function getWorkTasks_(payload) {
  const requestedDate = normalizeDate_(payload && (payload.date || payload.Date));
  const cacheKey = 'workTasks:' + getDataVersion_() + ':' + (requestedDate || 'all');

  return getCachedJson_(cacheKey, function () {
    const rows = getObjects_(getSheet_(SHEET_NAMES.work));

    let tasks = rows.map(function (row) {
      return {
        date: normalizeDate_(row.Date),
        task: row.Task || '',
        priority: row.Priority || '',
        status: row.Status || '',
        notes: row.Notes || '',
        remark: row.Remark || ''
      };
    });

    if (requestedDate) {
      tasks = tasks.filter(function (item) {
        return item.date === requestedDate;
      });
    }

    tasks.sort(function (a, b) {
      if (a.date === b.date) return 0;
      return a.date < b.date ? 1 : -1;
    });

    return {
      ok: true,
      tasks: tasks
    };
  }, 30);
}

function getPersonalTasks_(payload) {
  const requestedDate = normalizeDate_(payload && (payload.date || payload.Date));
  const cacheKey = 'personalTasks:' + getDataVersion_() + ':' + (requestedDate || 'all');

  return getCachedJson_(cacheKey, function () {
    const rows = getObjects_(getSheet_(SHEET_NAMES.personal));

    let tasks = rows.map(function (row) {
      return {
        date: normalizeDate_(row.Date),
        task: row.Task || '',
        category: row.Category || '',
        status: row.Status || '',
        remark: row.Remark || ''
      };
    });

    if (requestedDate) {
      tasks = tasks.filter(function (item) {
        return item.date === requestedDate;
      });
    }

    tasks.sort(function (a, b) {
      if (a.date === b.date) return 0;
      return a.date < b.date ? 1 : -1;
    });

    return {
      ok: true,
      tasks: tasks
    };
  }, 30);
}

function getJournalEntries_() {
  const cacheKey = 'journalEntries:' + getDataVersion_();

  return getCachedJson_(cacheKey, function () {
    const rows = getObjects_(getSheet_(SHEET_NAMES.journal));
    const dateSet = {};

    rows.forEach(function (row) {
      const date = normalizeDate_(row.Date);
      if (date) {
        dateSet[date] = true;
      }
    });

    const dates = Object.keys(dateSet).sort();

    return {
      ok: true,
      dates: dates
    };
  }, 30);
}

function getJournalHistory_(payload) {
  const limit = Number(payload.limit || payload.Limit || 7);
  const normalizedLimit = Math.max(1, Math.min(30, limit));
  const cacheKey = 'journalHistory:' + getDataVersion_() + ':' + normalizedLimit;

  return getCachedJson_(cacheKey, function () {
    const rows = getObjects_(getSheet_(SHEET_NAMES.journal));

    const entries = rows
      .map(function (row) {
        return {
          date: normalizeDate_(row.Date),
          mood: Number(row.Mood || 0)
        };
      })
      .filter(function (entry) {
        return entry.date;
      })
      .sort(function (a, b) {
        if (a.date === b.date) return 0;
        return a.date < b.date ? 1 : -1;
      })
      .slice(0, normalizedLimit);

    return {
      ok: true,
      entries: entries
    };
  }, 30);
}

function getAnalytics_(payload) {
  const selectedDate = payload && (payload.date || payload.Date);
  const today = normalizeDate_(selectedDate) || todayString_();
  const currentMonthKey = today.slice(0, 7);
  const currentYearKey = today.slice(0, 4);
  const rowsBySheet = loadRowsBySheet_();
  const workRows = rowsBySheet.work;
  const personalRows = rowsBySheet.personal;
  const habitRows = rowsBySheet.habits;
  const journalRows = rowsBySheet.journal;

  const weeklyProductivity = [];
  const moodTrend = [];

  for (let daysAgo = 6; daysAgo >= 0; daysAgo--) {
    const date = dateStringOffsetFrom_(today, daysAgo);
    const workAll = getWorkTasksByDate_(date, false, workRows);
    const personalAll = getPersonalTasksByDate_(date, false, personalRows);
    const habits = getHabitsByDate_(date, habitRows);
    const journal = getJournalByDate_(date, journalRows);

    const workDone = countDone_(workAll, 'Status');
    const personalDone = countDone_(personalAll, 'Status');
    const habitsDone = HABIT_COLUMNS.reduce(function (sum, name) {
      return sum + (Number(habits[name]) === 1 ? 1 : 0);
    }, 0);

    const score = calculateDailyScore_(
      workDone,
      workAll.length,
      personalDone,
      personalAll.length,
      habitsDone,
      HABIT_COLUMNS.length
    );

    weeklyProductivity.push({
      date: date,
      label: dayLabel_(date),
      score: score
    });

    moodTrend.push({
      date: date,
      label: dayLabel_(date),
      mood: Number(journal.Mood || 0)
    });
  }

  const monthHabitRows = habitRows.filter(function (row) {
    return normalizeDate_(row.Date).indexOf(currentMonthKey) === 0;
  });

  const monthHabitDone = monthHabitRows.reduce(function (sum, row) {
    return sum + HABIT_COLUMNS.reduce(function (inner, name) {
      return inner + (Number(row[name]) === 1 ? 1 : 0);
    }, 0);
  }, 0);

  const monthHabitPossible = monthHabitRows.length * HABIT_COLUMNS.length;
  const monthHabitPercent = monthHabitPossible > 0
    ? Math.round((monthHabitDone / monthHabitPossible) * 1000) / 10
    : 0;

  const activeDates = getActiveDatesForYear_(currentYearKey, rowsBySheet);
  const daysElapsed = getDaysElapsedInYear_(today);
  const consistencyPercent = daysElapsed > 0
    ? Math.round((activeDates.length / daysElapsed) * 1000) / 10
    : 0;

  return {
    ok: true,
    weeklyProductivity: weeklyProductivity,
    moodTrend: moodTrend,
    monthlyHabitCompletion: {
      month: currentMonthKey,
      completed: monthHabitDone,
      possible: monthHabitPossible,
      percentage: monthHabitPercent
    },
    yearlyConsistency: {
      year: currentYearKey,
      activeDays: activeDates.length,
      daysElapsed: daysElapsed,
      percentage: consistencyPercent
    }
  };
}

function getTimeTable_(payload) {
  const date = payload.date || payload.Date || '';
  const normalizedDate = normalizeDate_(date);
  const cacheKey = 'timeTable:' + getDataVersion_() + ':' + (normalizedDate || 'all');

  return getCachedJson_(cacheKey, function () {
    const rows = getObjects_(getSheet_(SHEET_NAMES.timetable));

    let entries = rows.map(function (row) {
      return {
        date: normalizeDate_(row.Date),
        startTime: row.StartTime || '',
        endTime: row.EndTime || '',
        block: row.Block || '',
        type: row.Type || '',
        linkedTask: row.LinkedTask || '',
        status: row.Status || '',
        remark: row.Remark || ''
      };
    });

    if (normalizedDate) {
      entries = entries.filter(function (entry) {
        return entry.date === normalizedDate;
      });
    }

    entries.sort(function (a, b) {
      if (a.date === b.date) {
        if (a.startTime === b.startTime) return 0;
        return String(a.startTime) > String(b.startTime) ? 1 : -1;
      }
      return a.date < b.date ? 1 : -1;
    });

    return {
      ok: true,
      entries: entries
    };
  }, 30);
}

function getActiveDatesForYear_(yearKey, rowsBySheet) {
  const dateSet = {};
  const rowGroups = rowsBySheet
    ? [rowsBySheet.work || [], rowsBySheet.personal || [], rowsBySheet.habits || [], rowsBySheet.journal || []]
    : [
      getObjects_(getSheet_(SHEET_NAMES.work)),
      getObjects_(getSheet_(SHEET_NAMES.personal)),
      getObjects_(getSheet_(SHEET_NAMES.habits)),
      getObjects_(getSheet_(SHEET_NAMES.journal))
    ];

  rowGroups.forEach(function (rows) {
    rows.forEach(function (row) {
      const date = normalizeDate_(row.Date);
      if (date.indexOf(yearKey) === 0) {
        dateSet[date] = true;
      }
    });
  });

  return Object.keys(dateSet).sort();
}

function getDaysElapsedInYear_(todayText) {
  const year = Number(todayText.slice(0, 4));
  const month = Number(todayText.slice(5, 7)) - 1;
  const day = Number(todayText.slice(8, 10));
  const today = new Date(year, month, day);
  const start = new Date(year, 0, 1);

  const diffMs = today.getTime() - start.getTime();
  return Math.floor(diffMs / 86400000) + 1;
}

function addWorkTask_(task) {
  const normalized = {
    Date: normalizeDate_(task.Date || task.date || todayString_()),
    Task: task.Task || task.task || '',
    Priority: task.Priority || task.priority || 'Medium',
    Status: task.Status || task.status || 'Pending',
    Notes: task.Notes || task.notes || '',
    Remark: task.Remark || task.remark || ''
  };

  validateRequired_(normalized.Task, 'Task is required');
  if (existsTaskDuplicate_(SHEET_NAMES.work, normalized.Date, normalized.Task)) {
    throw new Error('Duplicate work task for this date');
  }

  appendRowByHeaders_(SHEET_NAMES.work, normalized, ['Date', 'Task', 'Priority', 'Status', 'Notes', 'Remark']);
  invalidateDashboardCache_();

  return { ok: true, message: 'Work task added' };
}

function addPersonalTask_(task) {
  const normalized = {
    Date: normalizeDate_(task.Date || task.date || todayString_()),
    Task: task.Task || task.task || '',
    Category: task.Category || task.category || 'General',
    Status: task.Status || task.status || 'Pending',
    Remark: task.Remark || task.remark || ''
  };

  validateRequired_(normalized.Task, 'Task is required');
  if (existsTaskDuplicate_(SHEET_NAMES.personal, normalized.Date, normalized.Task)) {
    throw new Error('Duplicate personal task for this date');
  }

  appendRowByHeaders_(SHEET_NAMES.personal, normalized, ['Date', 'Task', 'Category', 'Status', 'Remark']);
  invalidateDashboardCache_();

  return { ok: true, message: 'Personal task added' };
}

function saveHabit_(payload) {
  const date = normalizeDate_(payload.date || payload.Date || todayString_());
  const habit = payload.habit || payload.Habit;
  const value = Number(payload.value != null ? payload.value : payload.Value);
  const remark = payload.remark || payload.Remark || '';

  validateRequired_(habit, 'Habit name is required');

  if (HABIT_COLUMNS.indexOf(habit) === -1) {
    throw new Error('Invalid habit: ' + habit);
  }

  const sheet = getSheet_(SHEET_NAMES.habits);
  const data = getSheetData_(sheet);
  const rows = data.rows;
  let rowIndex = -1;

  for (let i = 0; i < rows.length; i++) {
    if (normalizeDate_(rows[i].Date) === date) {
      rowIndex = i + 2;
      break;
    }
  }

  if (rowIndex === -1) {
    const seed = { Date: date, Study: 0, Reading: 0, Workout: 0, Planning: 0, Remark: '' };
    appendRowByHeaders_(SHEET_NAMES.habits, seed, ['Date', 'Study', 'Reading', 'Workout', 'Planning', 'Remark']);
    rowIndex = sheet.getLastRow();
  }

  const headers = rowIndex === -1 ? getHeaders_(sheet) : data.headers;
  const columnIndex = headers.indexOf(habit) + 1;
  if (!columnIndex) {
    throw new Error('Habit column missing: ' + habit);
  }

  sheet.getRange(rowIndex, columnIndex).setValue(value === 1 ? 1 : 0);

  const remarkCol = headers.indexOf('Remark') + 1;
  if (remarkCol && remark) {
    sheet.getRange(rowIndex, remarkCol).setValue(remark);
  }

  invalidateDashboardCache_();

  return { ok: true, message: 'Habit saved' };
}

function saveJournal_(payload) {
  const record = {
    Date: normalizeDate_(payload.date || payload.Date || todayString_()),
    Good: payload.good || payload.Good || '',
    Problem: payload.problem || payload.Problem || '',
    Improvement: payload.improvement || payload.Improvement || '',
    Mood: Number(payload.mood != null ? payload.mood : payload.Mood || 0),
    Remark: payload.remark || payload.Remark || ''
  };

  validateRequired_(record.Good, 'Good is required');
  validateRequired_(record.Problem, 'Problem is required');
  validateRequired_(record.Improvement, 'Improvement is required');

  const sheet = getSheet_(SHEET_NAMES.journal);
  upsertByDate_(sheet, record, ['Date', 'Good', 'Problem', 'Improvement', 'Mood', 'Remark']);
  invalidateDashboardCache_();

  return { ok: true, message: 'Journal saved' };
}

function saveDailyReview_(payload) {
  const record = {
    Date: normalizeDate_(payload.date || payload.Date || todayString_()),
    FollowUps: payload.followUps || payload.FollowUps || '',
    TomorrowPlan: payload.tomorrowPlan || payload.TomorrowPlan || '',
    SelfFeedback: payload.selfFeedback || payload.SelfFeedback || '',
    Remark: payload.remark || payload.Remark || ''
  };

  const sheet = getSheet_(SHEET_NAMES.review);
  upsertByDate_(sheet, record, ['Date', 'FollowUps', 'TomorrowPlan', 'SelfFeedback', 'Remark']);
  invalidateDashboardCache_();

  return { ok: true, message: 'Daily review saved' };
}

function addTimeTableEntry_(entry) {
  const normalized = {
    Date: normalizeDate_(entry.Date || entry.date || todayString_()),
    StartTime: entry.StartTime || entry.startTime || '',
    EndTime: entry.EndTime || entry.endTime || '',
    Block: entry.Block || entry.block || '',
    Type: entry.Type || entry.type || 'General',
    LinkedTask: entry.LinkedTask || entry.linkedTask || '',
    Status: entry.Status || entry.status || 'Pending',
    Remark: entry.Remark || entry.remark || ''
  };

  validateRequired_(normalized.StartTime, 'Start time is required');
  validateRequired_(normalized.EndTime, 'End time is required');
  validateRequired_(normalized.Block, 'Block is required');

  if (existsTimeTableDuplicate_(normalized.Date, normalized.Block, normalized.StartTime, normalized.EndTime)) {
    throw new Error('Duplicate timetable block for this date and time');
  }

  appendRowByHeaders_(
    SHEET_NAMES.timetable,
    normalized,
    ['Date', 'StartTime', 'EndTime', 'Block', 'Type', 'LinkedTask', 'Status', 'Remark']
  );

  invalidateDashboardCache_();
  return { ok: true, message: 'Time block added' };
}

function createWeeklyBackup_(payload) {
  const referenceDate = normalizeDate_(payload.date || payload.Date || todayString_());
  const weekStart = getWeekStart_(referenceDate);
  const sheet = getSheet_(SHEET_NAMES.backup);
  const requiredHeaders = [
    'WeekStart',
    'ReferenceDate',
    'GeneratedAt',
    'WorkTotal',
    'WorkDone',
    'PersonalTotal',
    'PersonalDone',
    'HabitPercent',
    'JournalDays',
    'TimetableTotal',
    'TimetableDone',
    'Snapshot'
  ];

  ensureHeaders_(sheet, requiredHeaders);

  const existingRows = getObjects_(sheet);
  const alreadyExists = existingRows.some(function (row) {
    return normalizeDate_(row.WeekStart) === weekStart;
  });

  if (alreadyExists) {
    return { ok: true, message: 'Backup already exists for this week', weekStart: weekStart };
  }

  const summary = buildLast30Summary_(referenceDate);
  const generatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  appendRowByHeaders_(SHEET_NAMES.backup, {
    WeekStart: weekStart,
    ReferenceDate: referenceDate,
    GeneratedAt: generatedAt,
    WorkTotal: summary.workTotal,
    WorkDone: summary.workDone,
    PersonalTotal: summary.personalTotal,
    PersonalDone: summary.personalDone,
    HabitPercent: summary.habitPercent,
    JournalDays: summary.journalDays,
    TimetableTotal: summary.timetableTotal,
    TimetableDone: summary.timetableDone,
    Snapshot: JSON.stringify(summary)
  }, requiredHeaders);

  return {
    ok: true,
    message: 'Weekly backup saved',
    weekStart: weekStart,
    summary: summary
  };
}

function saveDailyCheckin_(payload) {
  const date = normalizeDate_(payload.date || payload.Date || todayString_());

  const habits = payload.habits || {};
  const habitRecord = {
    Date: date,
    Study: Number(habits.Study || habits.study || 0) === 1 ? 1 : 0,
    Reading: Number(habits.Reading || habits.reading || 0) === 1 ? 1 : 0,
    Workout: Number(habits.Workout || habits.workout || 0) === 1 ? 1 : 0,
    Planning: Number(habits.Planning || habits.planning || 0) === 1 ? 1 : 0,
    Remark: payload.habitRemark || payload.remark || payload.Remark || ''
  };
  upsertByDate_(getSheet_(SHEET_NAMES.habits), habitRecord, ['Date', 'Study', 'Reading', 'Workout', 'Planning', 'Remark']);

  const journalRecord = {
    Date: date,
    Good: payload.good || payload.Good || '',
    Problem: payload.problem || payload.Problem || '',
    Improvement: payload.improvement || payload.Improvement || '',
    Mood: Number(payload.mood != null ? payload.mood : payload.Mood || 0),
    Remark: payload.journalRemark || payload.remark || payload.Remark || ''
  };
  upsertByDate_(getSheet_(SHEET_NAMES.journal), journalRecord, ['Date', 'Good', 'Problem', 'Improvement', 'Mood', 'Remark']);

  invalidateDashboardCache_();

  return { ok: true, message: 'Daily check-in saved' };
}

function completeWorkTask_(payload) {
  const task = payload.task || payload.Task || '';
  const date = payload.date || payload.Date || todayString_();

  const updated = markTaskDone_(SHEET_NAMES.work, date, task);
  if (!updated) {
    throw new Error('Pending work task not found');
  }

  invalidateDashboardCache_();
  return { ok: true, message: 'Work task marked done' };
}

function completePersonalTask_(payload) {
  const task = payload.task || payload.Task || '';
  const date = payload.date || payload.Date || todayString_();

  const updated = markTaskDone_(SHEET_NAMES.personal, date, task);
  if (!updated) {
    throw new Error('Pending personal task not found');
  }

  invalidateDashboardCache_();
  return { ok: true, message: 'Personal task marked done' };
}

function completeTimeTableEntry_(payload) {
  const date = payload.date || payload.Date || todayString_();
  const block = payload.block || payload.Block || '';
  const startTime = payload.startTime || payload.StartTime || '';

  const updated = markTimeTableDone_(date, block, startTime);
  if (!updated) {
    throw new Error('Pending timetable entry not found');
  }

  invalidateDashboardCache_();
  return { ok: true, message: 'Timetable entry marked done' };
}

function getCachedJson_(cacheKey, buildFn, ttlSeconds) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const result = buildFn();
  cache.put(cacheKey, JSON.stringify(result), ttlSeconds || 30);
  return result;
}

function getDataVersion_() {
  const cache = CacheService.getScriptCache();
  return cache.get('dataVersion') || '1';
}

function bumpDataVersion_() {
  const cache = CacheService.getScriptCache();
  cache.put('dataVersion', String(new Date().getTime()), 21600);
}

function getWeeklyScores_(workRows, personalRows, habitRows, referenceDate) {
  const output = [];
  const baseDate = normalizeDate_(referenceDate) || todayString_();

  for (let daysAgo = 6; daysAgo >= 0; daysAgo--) {
    const date = dateStringOffsetFrom_(baseDate, daysAgo);
    const workAll = getWorkTasksByDate_(date, false, workRows);
    const personalAll = getPersonalTasksByDate_(date, false, personalRows);
    const habits = getHabitsByDate_(date, habitRows);

    const workDone = countDone_(workAll, 'Status');
    const personalDone = countDone_(personalAll, 'Status');
    const habitsDone = HABIT_COLUMNS.reduce(function (sum, name) {
      return sum + (Number(habits[name]) === 1 ? 1 : 0);
    }, 0);

    const score = calculateDailyScore_(
      workDone,
      workAll.length,
      personalDone,
      personalAll.length,
      habitsDone,
      HABIT_COLUMNS.length
    );

    output.push(dayLabel_(date) + ' ' + score);
  }

  return output;
}

function getWorkTasksByDate_(dateText, pendingOnly, rows) {
  const sourceRows = rows || getObjects_(getSheet_(SHEET_NAMES.work));
  return sourceRows.filter(function (row) {
    if (normalizeDate_(row.Date) !== dateText) return false;
    if (!pendingOnly) return true;
    return !isDoneStatus_(row.Status);
  });
}

function getPersonalTasksByDate_(dateText, pendingOnly, rows) {
  const sourceRows = rows || getObjects_(getSheet_(SHEET_NAMES.personal));
  return sourceRows.filter(function (row) {
    if (normalizeDate_(row.Date) !== dateText) return false;
    if (!pendingOnly) return true;
    return !isDoneStatus_(row.Status);
  });
}

function getHabitsByDate_(dateText, rows) {
  const sourceRows = rows || getObjects_(getSheet_(SHEET_NAMES.habits));
  for (let i = 0; i < sourceRows.length; i++) {
    if (normalizeDate_(sourceRows[i].Date) === dateText) {
      return sourceRows[i];
    }
  }
  return { Date: dateText, Study: 0, Reading: 0, Workout: 0, Planning: 0, Remark: '' };
}

function getJournalByDate_(dateText, rows) {
  const sourceRows = rows || getObjects_(getSheet_(SHEET_NAMES.journal));
  for (let i = 0; i < sourceRows.length; i++) {
    if (normalizeDate_(sourceRows[i].Date) === dateText) {
      return sourceRows[i];
    }
  }
  return { Date: dateText, Mood: 0 };
}

function countDone_(rows, statusKey) {
  return rows.filter(function (row) {
    return isDoneStatus_(row[statusKey]);
  }).length;
}

function isDoneStatus_(status) {
  const value = String(status || '').toLowerCase().trim();
  return value === 'done' || value === 'completed' || value === 'complete' || value === 'delete' || value === 'deleted' || value === 'cancelled' || value === 'canceled';
}

function calculateDailyScore_(workDone, workTotal, personalDone, personalTotal, habitsDone, habitsTotal) {
  const workRatio = workTotal > 0 ? workDone / workTotal : 0;
  const personalRatio = personalTotal > 0 ? personalDone / personalTotal : 0;
  const habitRatio = habitsTotal > 0 ? habitsDone / habitsTotal : 0;

  const weighted = workRatio * 4 + personalRatio * 3 + habitRatio * 3;
  return Math.round(weighted * 10) / 10;
}

function appendRowByHeaders_(sheetName, object, requiredHeaders) {
  const sheet = getSheet_(sheetName);
  ensureHeaders_(sheet, requiredHeaders);
  const headers = getHeaders_(sheet);

  const row = headers.map(function (header) {
    return object[header] != null ? object[header] : '';
  });

  sheet.appendRow(row);
}

function upsertByDate_(sheet, object, requiredHeaders) {
  ensureHeaders_(sheet, requiredHeaders);
  const data = getSheetData_(sheet);
  const headers = data.headers;
  const dateValue = object.Date;
  const rows = data.rows;

  let rowIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (normalizeDate_(rows[i].Date) === normalizeDate_(dateValue)) {
      rowIndex = i + 2;
      break;
    }
  }

  const rowValues = headers.map(function (header) {
    return object[header] != null ? object[header] : '';
  });

  if (rowIndex === -1) {
    sheet.appendRow(rowValues);
    return;
  }

  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
}

function ensureHeaders_(sheet, requiredHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    return;
  }

  const headers = getHeaders_(sheet);
  const missing = requiredHeaders.filter(function (header) {
    return headers.indexOf(header) === -1;
  });

  if (!missing.length) return;

  const updated = headers.concat(missing);
  sheet.getRange(1, 1, 1, updated.length).setValues([updated]);
}

function getObjects_(sheet) {
  return getSheetData_(sheet).rows;
}

function getSheetData_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= 1 || lastColumn === 0) {
    if (lastRow === 1 && lastColumn > 0) {
      return {
        headers: sheet.getRange(1, 1, 1, lastColumn).getValues()[0],
        rows: []
      };
    }

    return {
      headers: [],
      rows: []
    };
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  if (values.length <= 1) {
    return {
      headers: values[0] || [],
      rows: []
    };
  }

  const headers = values[0];
  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const rowObj = {};
    for (let c = 0; c < headers.length; c++) {
      rowObj[headers[c]] = values[r][c];
    }
    rows.push(rowObj);
  }

  return {
    headers: headers,
    rows: rows
  };
}

function getHeaders_(sheet) {
  return getSheetData_(sheet).headers;
}

function getSheet_(name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    throw new Error('Missing sheet: ' + name);
  }
  return sheet;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function validateRequired_(value, message) {
  if (value == null || String(value).trim() === '') {
    throw new Error(message);
  }
}

function normalizeDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const date = new Date(text);
  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return text;
}

function todayString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function dateStringOffset_(daysAgo) {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function dateStringOffsetFrom_(baseDateText, daysAgo) {
  const date = new Date(baseDateText + 'T00:00:00');
  date.setDate(date.getDate() - Number(daysAgo || 0));
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function dayLabel_(dateText) {
  const date = new Date(dateText + 'T00:00:00');
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEE');
}

function getDailyScoreByDate_(dateText, workRows, personalRows, habitRows) {
  const workAll = getWorkTasksByDate_(dateText, false, workRows);
  const personalAll = getPersonalTasksByDate_(dateText, false, personalRows);
  const habits = getHabitsByDate_(dateText, habitRows);

  const workDone = countDone_(workAll, 'Status');
  const personalDone = countDone_(personalAll, 'Status');
  const habitsDone = HABIT_COLUMNS.reduce(function (sum, name) {
    return sum + (Number(habits[name]) === 1 ? 1 : 0);
  }, 0);

  return calculateDailyScore_(
    workDone,
    workAll.length,
    personalDone,
    personalAll.length,
    habitsDone,
    HABIT_COLUMNS.length
  );
}

function getMotivation_(score) {
  if (score >= 8) return 'Excellent day';
  if (score >= 6) return 'Good progress';
  if (score >= 4) return 'Keep pushing';
  return "Let's reset tomorrow";
}

function getHabitStreak_(habitName, referenceDate, rows) {
  const today = referenceDate || todayString_();
  const sourceRows = rows || getObjects_(getSheet_(SHEET_NAMES.habits));
  const dateMap = {};

  sourceRows.forEach(function (row) {
    const date = normalizeDate_(row.Date);
    if (date) {
      dateMap[date] = Number(row[habitName] || 0);
    }
  });

  return getBinaryDateStreak_(dateMap, today);
}

function getJournalStreak_(referenceDate, rows) {
  const today = referenceDate || todayString_();
  const sourceRows = rows || getObjects_(getSheet_(SHEET_NAMES.journal));
  const dateMap = {};

  sourceRows.forEach(function (row) {
    const date = normalizeDate_(row.Date);
    if (date) {
      const hasEntry = String(row.Good || '').trim() !== ''
        || String(row.Problem || '').trim() !== ''
        || String(row.Improvement || '').trim() !== ''
        || Number(row.Mood || 0) > 0;
      dateMap[date] = hasEntry ? 1 : 0;
    }
  });

  return getBinaryDateStreak_(dateMap, today);
}

function getBinaryDateStreak_(dateMap, referenceDate) {
  let streak = 0;
  const cursor = new Date(referenceDate + 'T00:00:00');

  while (true) {
    const dateKey = Utilities.formatDate(cursor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (Number(dateMap[dateKey] || 0) !== 1) {
      break;
    }

    streak += 1;
    cursor.setDate(cursor.getDate() - 1);
  }

  return streak;
}

function markTaskDone_(sheetName, dateText, taskName) {
  const sheet = getSheet_(sheetName);
  const data = getSheetData_(sheet);
  const headers = data.headers;
  const taskCol = headers.indexOf('Task') + 1;
  const dateCol = headers.indexOf('Date') + 1;
  const statusCol = headers.indexOf('Status') + 1;

  if (!taskCol || !dateCol || !statusCol) {
    throw new Error('Task sheet headers are invalid');
  }

  const rows = data.rows;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const sameDate = normalizeDate_(row.Date) === normalizeDate_(dateText);
    const sameTask = String(row.Task || '').trim() === String(taskName).trim();
    const pending = !isDoneStatus_(row.Status);

    if (sameDate && sameTask && pending) {
      sheet.getRange(i + 2, statusCol).setValue('Done');
      return true;
    }
  }

  return false;
}

function markTimeTableDone_(dateText, blockName, startTime) {
  const sheet = getSheet_(SHEET_NAMES.timetable);
  const data = getSheetData_(sheet);
  const headers = data.headers;
  const blockCol = headers.indexOf('Block') + 1;
  const dateCol = headers.indexOf('Date') + 1;
  const startTimeCol = headers.indexOf('StartTime') + 1;
  const statusCol = headers.indexOf('Status') + 1;

  if (!blockCol || !dateCol || !startTimeCol || !statusCol) {
    throw new Error('TimeTable sheet headers are invalid');
  }

  const rows = data.rows;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const sameDate = normalizeDate_(row.Date) === normalizeDate_(dateText);
    const sameBlock = String(row.Block || '').trim() === String(blockName).trim();
    const sameStartTime = String(row.StartTime || '').trim() === String(startTime).trim();
    const pending = !isDoneStatus_(row.Status);

    if (sameDate && sameBlock && sameStartTime && pending) {
      sheet.getRange(i + 2, statusCol).setValue('Done');
      return true;
    }
  }

  return false;
}

function existsTaskDuplicate_(sheetName, dateText, taskName) {
  const rows = getObjects_(getSheet_(sheetName));
  const normalizedDate = normalizeDate_(dateText);
  const normalizedTask = String(taskName || '').trim().toLowerCase();

  return rows.some(function (row) {
    return normalizeDate_(row.Date) === normalizedDate
      && String(row.Task || '').trim().toLowerCase() === normalizedTask;
  });
}

function existsTimeTableDuplicate_(dateText, blockName, startTime, endTime) {
  const rows = getObjects_(getSheet_(SHEET_NAMES.timetable));
  const normalizedDate = normalizeDate_(dateText);
  const normalizedBlock = String(blockName || '').trim().toLowerCase();
  const normalizedStart = String(startTime || '').trim();
  const normalizedEnd = String(endTime || '').trim();

  return rows.some(function (row) {
    return normalizeDate_(row.Date) === normalizedDate
      && String(row.Block || '').trim().toLowerCase() === normalizedBlock
      && String(row.StartTime || '').trim() === normalizedStart
      && String(row.EndTime || '').trim() === normalizedEnd;
  });
}

function buildLast30Summary_(referenceDate) {
  const startDate = dateStringOffsetFrom_(referenceDate, 29);
  const endDate = normalizeDate_(referenceDate);

  const workRows = getObjects_(getSheet_(SHEET_NAMES.work)).filter(function (row) {
    const date = normalizeDate_(row.Date);
    return date >= startDate && date <= endDate;
  });

  const personalRows = getObjects_(getSheet_(SHEET_NAMES.personal)).filter(function (row) {
    const date = normalizeDate_(row.Date);
    return date >= startDate && date <= endDate;
  });

  const habitRows = getObjects_(getSheet_(SHEET_NAMES.habits)).filter(function (row) {
    const date = normalizeDate_(row.Date);
    return date >= startDate && date <= endDate;
  });

  const journalRows = getObjects_(getSheet_(SHEET_NAMES.journal)).filter(function (row) {
    const date = normalizeDate_(row.Date);
    return date >= startDate && date <= endDate;
  });

  const timetableRows = getObjects_(getSheet_(SHEET_NAMES.timetable)).filter(function (row) {
    const date = normalizeDate_(row.Date);
    return date >= startDate && date <= endDate;
  });

  const workDone = countDone_(workRows, 'Status');
  const personalDone = countDone_(personalRows, 'Status');
  const habitDone = habitRows.reduce(function (sum, row) {
    return sum + HABIT_COLUMNS.reduce(function (inner, habitName) {
      return inner + (Number(row[habitName]) === 1 ? 1 : 0);
    }, 0);
  }, 0);

  const habitPossible = habitRows.length * HABIT_COLUMNS.length;
  const habitPercent = habitPossible > 0
    ? Math.round((habitDone / habitPossible) * 1000) / 10
    : 0;

  const journalDays = journalRows.filter(function (row) {
    return String(row.Good || '').trim() !== ''
      || String(row.Problem || '').trim() !== ''
      || String(row.Improvement || '').trim() !== ''
      || Number(row.Mood || 0) > 0;
  }).length;

  const timetableDone = countDone_(timetableRows, 'Status');

  return {
    rangeStart: startDate,
    rangeEnd: endDate,
    workTotal: workRows.length,
    workDone: workDone,
    personalTotal: personalRows.length,
    personalDone: personalDone,
    habitPercent: habitPercent,
    journalDays: journalDays,
    timetableTotal: timetableRows.length,
    timetableDone: timetableDone
  };
}

function getWeekStart_(dateText) {
  const date = new Date(dateText + 'T00:00:00');
  const day = date.getDay();
  const diffToMonday = (day + 6) % 7;
  date.setDate(date.getDate() - diffToMonday);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function invalidateDashboardCache_() {
  const cache = CacheService.getScriptCache();
  cache.remove('dashboard:' + getDataVersion_() + ':' + todayString_());
  bumpDataVersion_();
}

function loadRowsBySheet_() {
  return {
    work: getObjects_(getSheet_(SHEET_NAMES.work)),
    personal: getObjects_(getSheet_(SHEET_NAMES.personal)),
    habits: getObjects_(getSheet_(SHEET_NAMES.habits)),
    journal: getObjects_(getSheet_(SHEET_NAMES.journal))
  };
}
