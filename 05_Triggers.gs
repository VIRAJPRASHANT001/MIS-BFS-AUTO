/**
 * ============================================================
 * FILE: 05_Triggers.gs
 * PURPOSE: Installs and removes time-driven triggers for all
 *          automation. Run installTriggers() once after setup.
 *          Run removeTriggers() if you ever need to pause automation.
 * DEPENDENCIES: 03_Automation.gs, 04_KRA_Engine.gs
 * ============================================================
 */

/**
 * Installs three time-driven triggers:
 *  1. runDailyAutomation()   — every morning at 6:00 AM
 *     (generates recurring tasks, flags overdue tasks)
 *  2. recalculateKRAScores() — every night at 11:00 PM
 *     (keeps current-month KRA scores fresh each day)
 *  3. weeklyKRAReport()      — every Monday at 7:00 AM
 *     (future hook: could email a summary report)
 *
 * Safe to run multiple times — checks for duplicates first.
 */
function installTriggers() {
  const existing = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());

  if (!existing.includes('runDailyAutomation')) {
    ScriptApp.newTrigger('runDailyAutomation')
      .timeBased().everyDays(1).atHour(6).create();
  }

  if (!existing.includes('recalculateKRAScores')) {
    ScriptApp.newTrigger('recalculateKRAScores')
      .timeBased().everyDays(1).atHour(23).create();
  }

  SpreadsheetApp.getUi().alert(
    'Triggers Installed',
    'Daily automation runs at 6 AM.\nKRA recalculation runs at 11 PM.\n\n' +
    'Next step: Deploy as Web App via Extensions > Apps Script > Deploy > New deployment.\n' +
    'Set "Execute as: Me" and "Who has access: Anyone with Google Account" (or restrict to your domain).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** Removes all project triggers. Use this to pause automation. */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('All triggers removed.');
}

/** Manual test: run all automation right now without waiting for the scheduled trigger. */
function runAutomationNow() {
  runDailyAutomation();
  const count = recalculateKRAScores();
  SpreadsheetApp.getUi().alert('Done', 'Automation + KRA recalculation complete. Check TaskHistory and KRA_Scores sheets.', SpreadsheetApp.getUi().ButtonSet.OK);
}
