/**
 * FILE: 03_Automation.gs — v2.1
 * Recurring task generation, overdue detection, workflow advancement,
 * evidence approval handling.
 */

const TASK_COL = {
  TaskID:1, Title:2, Type:3, AssignedTo:4, AssignedBy:5, Department:6,
  CreatedDate:7, DueDate:8, Status:9, CompletedDate:10, Priority:11,
  RecurrenceRule:12, RecurrenceParentID:13, WorkflowID:14, WorkflowStepNumber:15,
  Remarks:16, RequiresEvidence:17, EvidenceURL:18, EvidenceApprovedBy:19, EvidenceApprovedDate:20
};

function runDailyAutomation() {
  generateRecurringTaskInstances();
  flagOverdueTasks();
  Logger.log('Daily automation done: ' + new Date());
}

// ── RECURRING TASK GENERATION ──────────────────────────────────────────────

function generateRecurringTaskInstances() {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const data = sh.getDataRange().getValues();
  const today = new Date(); today.setHours(0,0,0,0);
  const newRows = [];
  let maxId = getMaxTaskIdNumber_(data);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[TASK_COL.Status-1]) !== 'Active-Template') continue;
    const rule     = row[TASK_COL.RecurrenceRule-1];
    const parentId = String(row[TASK_COL.TaskID-1]);
    if (!isInstanceDueToday_(rule, parentId, data, today)) continue;

    maxId++;
    const newId  = 'TSK' + String(maxId).padStart(4,'0');
    const dueDate = calcDueDate_(rule, today);
    const newRow  = row.slice();
    newRow[TASK_COL.TaskID-1]             = newId;
    newRow[TASK_COL.CreatedDate-1]        = today;
    newRow[TASK_COL.DueDate-1]            = dueDate;
    newRow[TASK_COL.Status-1]             = 'Pending';
    newRow[TASK_COL.CompletedDate-1]      = '';
    newRow[TASK_COL.RecurrenceParentID-1] = parentId;
    newRow[TASK_COL.Remarks-1]            = 'Auto-generated ' + today.toDateString();
    newRow[TASK_COL.EvidenceURL-1]        = '';
    newRow[TASK_COL.EvidenceApprovedBy-1] = '';
    newRow[TASK_COL.EvidenceApprovedDate-1] = '';
    newRows.push(newRow);
  }

  if (newRows.length) {
    sh.getRange(sh.getLastRow()+1, 1, newRows.length, newRows[0].length).setValues(newRows);
    Logger.log('Generated ' + newRows.length + ' recurring instance(s).');
  }
}

function isInstanceDueToday_(rule, parentId, allData, today) {
  let lastDate = null;
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][TASK_COL.RecurrenceParentID-1]) === parentId) {
      const cd = new Date(allData[i][TASK_COL.CreatedDate-1]);
      if (!lastDate || cd > lastDate) lastDate = cd;
    }
  }
  if (!lastDate) return true;
  lastDate.setHours(0,0,0,0);
  const diff = Math.round((today - lastDate) / 86400000);
  if (rule === 'Daily')   return diff >= 1;
  if (rule === 'Weekly')  return diff >= 7;
  if (rule === 'Monthly') return diff >= 28;
  return false;
}

function calcDueDate_(rule, today) {
  const d = new Date(today);
  if (rule === 'Daily')   d.setDate(d.getDate());
  if (rule === 'Weekly')  d.setDate(d.getDate() + 2);
  if (rule === 'Monthly') d.setDate(d.getDate() + 5);
  return d;
}

function getMaxTaskIdNumber_(data) {
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const m = String(data[i][0]).match(/TSK(\d+)/);
    if (m) max = Math.max(max, parseInt(m[1],10));
  }
  return max;
}

// ── OVERDUE DETECTION ─────────────────────────────────────────────────────

function flagOverdueTasks() {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const data = sh.getDataRange().getValues();
  const today = new Date(); today.setHours(0,0,0,0);

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][TASK_COL.Status-1]);
    if (status !== 'Pending' && status !== 'InProgress') continue;
    if (!data[i][TASK_COL.DueDate-1]) continue;
    const due = new Date(data[i][TASK_COL.DueDate-1]); due.setHours(0,0,0,0);
    const days = Math.round((today - due) / 86400000);
    if (days <= 0) continue;
    const newStatus = days >= 4 ? 'Missed' : 'Late';
    if (newStatus === status) continue;
    sh.getRange(i+1, TASK_COL.Status).setValue(newStatus);
    logTaskHistory_(data[i][0], data[i][TASK_COL.AssignedTo-1], 'Auto-Flag', status, newStatus, false);
  }
}

// ── TASK COMPLETION ───────────────────────────────────────────────────────

/**
 * Marks a task complete/approved. If RequiresEvidence=YES, sets status to
 * Evidence-Pending instead of Completed so supervisor must approve evidence.
 */
function completeTask(taskId, completedByEmail, remarks, evidenceUrl) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(taskId)) continue;
    const row         = data[i];
    const dueDate     = row[TASK_COL.DueDate-1] ? new Date(row[TASK_COL.DueDate-1]) : null;
    const now         = new Date();
    const isWorkflow  = row[TASK_COL.Type-1] === 'Workflow';
    const needsEvidence = String(row[TASK_COL.RequiresEvidence-1]).toUpperCase() === 'YES';
    const oldStatus   = String(row[TASK_COL.Status-1]);
    const onTime      = !dueDate || now <= new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate(), 23, 59, 59);

    let newStatus;
    if (isWorkflow) {
      newStatus = 'Approved';
    } else if (needsEvidence && !evidenceUrl) {
      // Doer submitted without evidence — mark as Evidence-Pending for supervisor
      newStatus = 'Evidence-Pending';
    } else {
      newStatus = 'Completed';
    }

    sh.getRange(i+1, TASK_COL.Status).setValue(newStatus);
    sh.getRange(i+1, TASK_COL.CompletedDate).setValue(now);
    if (remarks) sh.getRange(i+1, TASK_COL.Remarks).setValue(remarks);
    if (evidenceUrl) sh.getRange(i+1, TASK_COL.EvidenceURL).setValue(evidenceUrl);

    logTaskHistory_(taskId, completedByEmail, isWorkflow?'Approve':'Complete', oldStatus, newStatus, onTime);

    if (isWorkflow) advanceWorkflow_(row, taskId);

    return {
      success: true, onTime: onTime, newStatus: newStatus,
      message: newStatus === 'Evidence-Pending'
        ? 'Task submitted. Waiting for supervisor to upload/approve evidence.'
        : (isWorkflow ? 'Step approved — advanced to next stage.' : 'Task marked complete.')
    };
  }
  return { success: false, message: 'Task ID not found: ' + taskId };
}

/**
 * Supervisor approves evidence for a task — uploads evidence URL and
 * moves status from Evidence-Pending to Completed.
 */
function approveEvidence(taskId, supervisorEmail, evidenceUrl) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(taskId)) continue;
    const oldStatus = String(data[i][TASK_COL.Status-1]);
    if (oldStatus !== 'Evidence-Pending') return { success:false, message:'Task is not in Evidence-Pending status.' };

    sh.getRange(i+1, TASK_COL.Status).setValue('Completed');
    sh.getRange(i+1, TASK_COL.EvidenceURL).setValue(evidenceUrl || '');
    sh.getRange(i+1, TASK_COL.EvidenceApprovedBy).setValue(supervisorEmail);
    sh.getRange(i+1, TASK_COL.EvidenceApprovedDate).setValue(new Date());
    logTaskHistory_(taskId, supervisorEmail, 'EvidenceApproved', oldStatus, 'Completed', true);
    return { success: true, message: 'Evidence approved. Task marked complete.' };
  }
  return { success: false, message: 'Task not found.' };
}

// ── WORKFLOW ADVANCEMENT ──────────────────────────────────────────────────

function advanceWorkflow_(completedRow, completedTaskId) {
  const workflowId  = completedRow[TASK_COL.WorkflowID-1];
  const currentStep = Number(completedRow[TASK_COL.WorkflowStepNumber-1]);

  const wfSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_WORKFLOWS);
  const wfData  = wfSheet.getDataRange().getValues();
  let nextStepRow = null;

  for (let i = 1; i < wfData.length; i++) {
    if (wfData[i][0] === workflowId && Number(wfData[i][2]) === currentStep) {
      if (String(wfData[i][6]).toUpperCase() === 'TRUE') return; // final step — done
      const nextNum = Number(wfData[i][5]);
      for (let j = 1; j < wfData.length; j++) {
        if (wfData[j][0] === workflowId && Number(wfData[j][2]) === nextNum) {
          nextStepRow = wfData[j]; break;
        }
      }
    }
  }

  if (!nextStepRow) { Logger.log('WARNING: next workflow step not found for ' + workflowId); return; }

  const taskSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const today  = new Date();
  const due    = new Date(today); due.setDate(due.getDate() + 2);
  const maxId  = getMaxTaskIdNumber_(taskSh.getDataRange().getValues());
  const newId  = 'TSK' + String(maxId+1).padStart(4,'0');

  const newRow = [
    newId,
    completedRow[TASK_COL.Title-1],
    'Workflow',
    nextStepRow[4],
    completedRow[TASK_COL.AssignedBy-1],
    completedRow[TASK_COL.Department-1],
    today, due,
    'Pending', '',
    completedRow[TASK_COL.Priority-1],
    '', completedTaskId,
    workflowId, nextStepRow[2],
    'Auto-advanced from step ' + currentStep,
    completedRow[TASK_COL.RequiresEvidence-1] || 'NO',
    '', '', ''
  ];

  taskSh.getRange(taskSh.getLastRow()+1, 1, 1, newRow.length).setValues([newRow]);
  logTaskHistory_(newId, nextStepRow[4], 'Workflow-Advance', '', 'Pending', null);
}

// ── HISTORY LOGGING ───────────────────────────────────────────────────────

function logTaskHistory_(taskId, empEmail, action, oldStatus, newStatus, onTime) {
  const sh  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HISTORY);
  const lid = 'LOG' + new Date().getTime() + Math.floor(Math.random()*1000);
  sh.appendRow([lid, taskId, empEmail, action, oldStatus, newStatus, new Date(),
                onTime === null ? 'N/A' : onTime]);
}
