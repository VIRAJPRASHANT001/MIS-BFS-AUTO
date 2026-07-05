/**
 * ============================================================
 * FILE: 04_KRA_Engine.gs  — v2
 * PURPOSE: KRA/performance scoring with Department KRA support.
 *
 * SCORING FORMULA:
 *   OnTimePercent = CompletedOnTime / TotalAssigned * 100
 *   VolumeScore   = min(TotalAssigned / VOLUME_BENCHMARK, 1) * 100
 *   KRAScore      = (OnTimePct * OnTimeWeight) + (VolumeScore * VolumeWeight)
 *
 * FIX v2: Period label stored AND compared as plain string.
 *         CreatedDate filter widened to last 90 days so demo data
 *         always has signal regardless of when setup was run.
 *         Return value no longer includes the full rows array
 *         (avoids Apps Script 10MB response limit).
 * ============================================================
 */

const VOLUME_BENCHMARK = 10; // tasks/month = "full" volume score

/**
 * Recalculates individual + department KRA scores for the
 * current calendar month. Safe to run repeatedly — wipes only
 * the current-period rows before rewriting.
 */
function recalculateKRAScores() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet  = ss.getSheetByName(SHEET_EMPLOYEES);
  const taskSheet = ss.getSheetByName(SHEET_TASKS);
  const kraSheet  = ss.getSheetByName(SHEET_KRA);
  const weights   = getKRAWeights_();

  const employees = empSheet.getDataRange().getValues();
  const tasks     = taskSheet.getDataRange().getValues();

  const today        = new Date();
  // FIX: always format as plain string — never store a Date object in this column
  const periodLabel  = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM');

  // FIX: widen window to 90 days back so demo data (created weeks ago) is always included
  const periodStart  = new Date(today); periodStart.setDate(periodStart.getDate() - 90);
  const periodEnd    = new Date(today);

  removeExistingPeriodRows_(kraSheet, periodLabel);

  const newRows    = [];
  const deptTotals = {}; // { deptName: { assigned, onTime, late, missed, empCount } }

  for (let e = 1; e < employees.length; e++) {
    const emp = employees[e];
    if (String(emp[6]).trim() !== 'Active') continue;

    const empEmail = emp[2];
    const empDept  = emp[4];
    let assigned = 0, onTime = 0, late = 0, missed = 0;

    for (let t = 1; t < tasks.length; t++) {
      const row    = tasks[t];
      const status = String(row[TASK_COL.AssignedTo - 1]) === empEmail
                     ? String(row[TASK_COL.Status - 1]) : null;
      if (!status) continue;
      if (status === 'Active-Template' || status === 'Pending' || status === 'InProgress') continue;

      // FIX: use CreatedDate with 90-day window instead of strict month boundary
      const created = row[TASK_COL.CreatedDate - 1] ? new Date(row[TASK_COL.CreatedDate - 1]) : null;
      if (!created || created < periodStart || created > periodEnd) continue;

      assigned++;
      if (status === 'Completed' || status === 'Approved') onTime++;
      else if (status === 'Late')   late++;
      else if (status === 'Missed') missed++;
    }

    const onTimePct   = assigned > 0 ? Math.round((onTime / assigned) * 1000) / 10 : 0;
    const volumeScore = Math.min(assigned / VOLUME_BENCHMARK, 1) * 100;
    const kraScore    = assigned > 0
      ? Math.round(((onTimePct * weights.onTimeWeight) + (volumeScore * weights.volumeWeight)) * 10) / 10
      : 0;

    // Row: EmpID, Name, Email, Period, Assigned, OnTime, Late, Missed, OnTimePct, KRAScore, LastCalc, Department
    newRows.push([emp[0], emp[1], empEmail, periodLabel,
                  assigned, onTime, late, missed, onTimePct, kraScore,
                  Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'),
                  empDept]);

    // Accumulate department totals
    if (!deptTotals[empDept]) deptTotals[empDept] = { assigned:0, onTime:0, late:0, missed:0, empCount:0, kraSum:0 };
    deptTotals[empDept].assigned  += assigned;
    deptTotals[empDept].onTime    += onTime;
    deptTotals[empDept].late      += late;
    deptTotals[empDept].missed    += missed;
    deptTotals[empDept].kraSum    += kraScore;
    deptTotals[empDept].empCount  += 1;
  }

  if (newRows.length > 0) {
    kraSheet.getRange(kraSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    applyKRAConditionalFormat_(kraSheet);
  }

  // Save department KRA to Config sheet for fast retrieval
  saveDepartmentKRA_(deptTotals, periodLabel, weights);

  Logger.log('KRA v2 recalculated: ' + periodLabel + ', ' + newRows.length + ' employees.');
  // FIX: return only lightweight summary, not full rows array
  return { period: periodLabel, count: newRows.length };
}

/**
 * Saves department-level KRA summary into a DeptKRA sheet.
 */
function saveDepartmentKRA_(deptTotals, periodLabel, weights) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('DeptKRA');
  if (!sh) {
    sh = ss.insertSheet('DeptKRA');
    sh.getRange(1,1,1,9).setValues([['Department','Period','Employees','Assigned','OnTime','Late','Missed','AvgOnTimePct','AvgKRAScore']])
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
    sh.setFrozenRows(1);
  }

  // Remove existing rows for this period
  const existing = sh.getDataRange().getValues();
  for (let i = existing.length - 1; i >= 1; i--) {
    if (String(existing[i][1]) === periodLabel) sh.deleteRow(i + 1);
  }

  const deptRows = [];
  for (const dept in deptTotals) {
    const d = deptTotals[dept];
    const avgOnTimePct = d.assigned > 0 ? Math.round((d.onTime / d.assigned) * 1000) / 10 : 0;
    const avgKRA       = d.empCount  > 0 ? Math.round((d.kraSum / d.empCount) * 10) / 10 : 0;
    deptRows.push([dept, periodLabel, d.empCount, d.assigned, d.onTime, d.late, d.missed, avgOnTimePct, avgKRA]);
  }

  if (deptRows.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, deptRows.length, deptRows[0].length).setValues(deptRows);
    // Color code AvgKRAScore column (col 9)
    const scoreRange = sh.getRange(2, 9, Math.max(sh.getLastRow()-1,1), 1);
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(85).setBackground('#d9ead3').setRanges([scoreRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(60,84.99).setBackground('#fff2cc').setRanges([scoreRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(60).setBackground('#f4cccc').setRanges([scoreRange]).build()
    ];
    sh.setConditionalFormatRules(rules);
  }
}

function removeExistingPeriodRows_(sh, periodLabel) {
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    // FIX: coerce to string before comparing — Sheets can return dates as Date objects
    if (String(data[i][3]).trim() === String(periodLabel).trim()) sh.deleteRow(i + 1);
  }
}

function getKRAWeights_() {
  const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG).getDataRange().getValues();
  let onTimeWeight = 0.7, volumeWeight = 0.3;
  for (let i = 1; i < cfg.length; i++) {
    if (cfg[i][0] === 'KRA_OnTimeWeight') onTimeWeight = parseFloat(cfg[i][1]) || 0.7;
    if (cfg[i][0] === 'KRA_VolumeWeight') volumeWeight = parseFloat(cfg[i][1]) || 0.3;
  }
  return { onTimeWeight, volumeWeight };
}

function applyKRAConditionalFormat_(sh) {
  const lastRow = Math.max(sh.getLastRow(), 2);
  const range   = sh.getRange(2, 10, lastRow - 1, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(85).setBackground('#d9ead3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(60, 84.99).setBackground('#fff2cc').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(60).setBackground('#f4cccc').setRanges([range]).build()
  ]);
}

/** Returns KRA data for one email — used by web app. */
function getMyKRA_(email) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KRA);
  const data = sh.getDataRange().getValues();
  const periodLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  for (let i = 1; i < data.length; i++) {
    // FIX: coerce both sides to string for safe comparison
    if (String(data[i][2]).trim() === email && String(data[i][3]).trim() === periodLabel) {
      return { period: String(data[i][3]), tasksAssigned: data[i][4], onTime: data[i][5],
               late: data[i][6], missed: data[i][7], onTimePct: data[i][8], kraScore: data[i][9] };
    }
  }
  return null;
}

/** Returns all individual KRA rows for current period. */
function getAllKRAScores_() {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KRA);
  const data = sh.getDataRange().getValues();
  const periodLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === periodLabel) {
      result.push({ empId: data[i][0], name: data[i][1], email: data[i][2],
                    period: String(data[i][3]), assigned: data[i][4], onTime: data[i][5],
                    late: data[i][6], missed: data[i][7], onTimePct: data[i][8],
                    kraScore: data[i][9], department: data[i][11] || '' });
    }
  }
  return result.sort((a, b) => (b.kraScore || 0) - (a.kraScore || 0));
}

/** Returns department KRA scores for current period. */
function getDeptKRAScores_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('DeptKRA');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const periodLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === periodLabel) {
      result.push({ department: data[i][0], period: data[i][1], empCount: data[i][2],
                    assigned: data[i][3], onTime: data[i][4], late: data[i][5],
                    missed: data[i][6], avgOnTimePct: data[i][7], avgKRA: data[i][8] });
    }
  }
  return result.sort((a, b) => (b.avgKRA || 0) - (a.avgKRA || 0));
}
