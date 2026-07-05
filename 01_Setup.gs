/**
 * FILE: 01_Setup.gs — v2.1
 * Run setupMIS() once on a blank sheet to initialize everything.
 */

const SHEET_EMPLOYEES  = 'Employees';
const SHEET_ROLES      = 'Roles';
const SHEET_TASKS      = 'Tasks';
const SHEET_WORKFLOWS  = 'Workflows';
const SHEET_HISTORY    = 'TaskHistory';
const SHEET_KRA        = 'KRA_Scores';
const SHEET_CONFIG     = 'Config';
const SHEET_DEPT       = 'Departments';

function setupMIS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createDepartmentsSheet_(ss);
  createEmployeesSheet_(ss);
  createRolesSheet_(ss);
  createWorkflowsSheet_(ss);
  createTasksSheet_(ss);
  createTaskHistorySheet_(ss);
  createKRASheet_(ss);
  createConfigSheet_(ss);
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);
  SpreadsheetApp.getUi().alert('MIS Setup Complete',
    'All sheets created.\nNext: Run loadDemoData(), then installTriggers(), then deploy as Web App.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function createDepartmentsSheet_(ss) {
  if (ss.getSheetByName(SHEET_DEPT)) return;
  const sh = ss.insertSheet(SHEET_DEPT);
  const h = ['DeptID','DeptName','HeadEmail','Description','Status'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  sh.setFrozenRows(1); sh.autoResizeColumns(1,h.length);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Active','Inactive']).build();
  sh.getRange(2,5,998,1).setDataValidation(rule);
}

function createEmployeesSheet_(ss) {
  if (ss.getSheetByName(SHEET_EMPLOYEES)) return;
  const sh = ss.insertSheet(SHEET_EMPLOYEES);
  const h = ['EmpID','Name','Email','Role','Department','SupervisorEmail','Status','JoinDate','Phone','Notes'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  sh.setFrozenRows(1); sh.autoResizeColumns(1,h.length);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Active','Inactive']).build();
  sh.getRange(2,7,998,1).setDataValidation(rule);
}

function createRolesSheet_(ss) {
  if (ss.getSheetByName(SHEET_ROLES)) return;
  const sh = ss.insertSheet(SHEET_ROLES);
  const h = ['RoleName','CanViewTeamTasks','CanApprove','CanViewAllKRA','CanCreateTasks','Description'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  sh.setFrozenRows(1); sh.autoResizeColumns(1,h.length);
  const bool = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE','FALSE']).build();
  sh.getRange(2,2,998,4).setDataValidation(bool);
}

function createWorkflowsSheet_(ss) {
  if (ss.getSheetByName(SHEET_WORKFLOWS)) return;
  const sh = ss.insertSheet(SHEET_WORKFLOWS);
  const h = ['WorkflowID','WorkflowName','StepNumber','StepRole','StepApproverEmail','NextStepNumber','IsFinalStep'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  sh.setFrozenRows(1); sh.autoResizeColumns(1,h.length);
}

function createTasksSheet_(ss) {
  if (ss.getSheetByName(SHEET_TASKS)) return;
  const sh = ss.insertSheet(SHEET_TASKS);
  const h = [
    'TaskID','Title','Type','AssignedTo','AssignedBy','Department',
    'CreatedDate','DueDate','Status','CompletedDate','Priority',
    'RecurrenceRule','RecurrenceParentID','WorkflowID','WorkflowStepNumber',
    'Remarks','RequiresEvidence','EvidenceURL','EvidenceApprovedBy','EvidenceApprovedDate'
  ];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  sh.setFrozenRows(1); sh.autoResizeColumns(1,h.length);

  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['OneTime','Recurring','Workflow']).build();
  sh.getRange(2,3,998,1).setDataValidation(typeRule);
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(
    ['Pending','InProgress','Completed','Late','Missed','Approved','Rejected','Cancelled','Evidence-Pending']).build();
  sh.getRange(2,9,998,1).setDataValidation(statusRule);
  const priorityRule = SpreadsheetApp.newDataValidation().requireValueInList(['Low','Medium','High','Urgent']).build();
  sh.getRange(2,11,998,1).setDataValidation(priorityRule);
  const recurRule = SpreadsheetApp.newDataValidation().requireValueInList(['','Daily','Weekly','Monthly']).setAllowInvalid(true).build();
  sh.getRange(2,12,998,1).setDataValidation(recurRule);
  const evRule = SpreadsheetApp.newDataValidation().requireValueInList(['YES','NO']).build();
  sh.getRange(2,17,998,1).setDataValidation(evRule);
  applyTaskStatusFormatting_(sh);
}

function applyTaskStatusFormatting_(sh) {
  const r = sh.getRange(2,9,998,1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Completed').setBackground('#d9ead3').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Approved').setBackground('#d9ead3').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Late').setBackground('#fce5cd').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Missed').setBackground('#f4cccc').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pending').setBackground('#fff2cc').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Evidence-Pending').setBackground('#e1d5f7').setRanges([r]).build()
  ]);
}

function createTaskHistorySheet_(ss) {
  if (ss.getSheetByName(SHEET_HISTORY)) return;
  const sh = ss.insertSheet(SHEET_HISTORY);
  const h = ['LogID','TaskID','EmpEmail','Action','OldStatus','NewStatus','Timestamp','OnTime'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  sh.setFrozenRows(1);
}

function createKRASheet_(ss) {
  if (ss.getSheetByName(SHEET_KRA)) return;
  const sh = ss.insertSheet(SHEET_KRA);
  const h = ['EmpID','Name','Email','Period','TasksAssigned','OnTime','Late','Missed',
             'OnTimePct','KRAScore','LastCalculated','Department'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  sh.setFrozenRows(1);
}

function createConfigSheet_(ss) {
  if (ss.getSheetByName(SHEET_CONFIG)) return;
  const sh = ss.insertSheet(SHEET_CONFIG);
  sh.getRange(1,1,1,2).setValues([['Key','Value']]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  sh.getRange(2,1,6,2).setValues([
    ['OrgName','Demo School / Office MIS'],
    ['KRA_OnTimeWeight','0.7'],
    ['KRA_VolumeWeight','0.3'],
    ['EvidenceFolderID',''],
    ['KPI_Target_OnTime','85'],
    ['KPI_Target_Completion','90']
  ]);
  sh.setFrozenRows(1); sh.autoResizeColumns(1,2);
}
