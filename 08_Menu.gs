/**
 * FILE: 08_Menu.gs — v2.1
 * Adds "MIS System" menu to the Google Sheet UI.
 * DEPENDENCIES: All other .gs files
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏫 MIS System')
    .addItem('1. Initial Setup (blank sheet only)', 'setupMIS')
    .addItem('2. Load / Refresh Demo Data',         'loadDemoData')
    .addItem('3. Install Automation Triggers',       'installTriggers')
    .addSeparator()
    .addItem('▶ Run All Automation Now',             'runAutomationNow')
    .addItem('▶ Recalculate KRA Scores',             'recalculateKRAScores')
    .addSeparator()
    .addItem('⚠ Remove All Triggers',               'removeTriggers')
    .addToUi();
}
