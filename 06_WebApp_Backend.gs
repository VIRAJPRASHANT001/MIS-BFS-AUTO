/**
 * FILE: 06_WebApp_Backend.gs — v2.1
 * All server-side API. Fixed task filtering, added evidence upload,
 * KPI targets, department CRUD, employee CRUD.
 */


function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('07_Dashboard_UI')
    .setTitle('MIS Dashboard v2')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── AUTH ─────────────────────────────────────────────────────────────────

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return { error: 'Could not detect Google account. Make sure you are signed in.' };
  const emp = getEmployeeByEmail_(email);
  if (!emp) return { error: 'Your email (' + email + ') is not in the Employees sheet. Ask your admin to add you.' };
  const perms = getRolePermissions_(emp.role);
  return { ...emp, ...perms };
}

function getEmployeeByEmail_(email) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === String(email).trim() && String(data[i][6]) === 'Active') {
      return { empId:data[i][0], name:data[i][1], email:data[i][2],
               role:data[i][3], department:data[i][4], supervisorEmail:data[i][5] };
    }
  }
  return null;
}

function getRolePermissions_(roleName) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROLES).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === roleName) {
      return {
        canViewTeamTasks: data[i][1]==='TRUE'||data[i][1]===true,
        canApprove:       data[i][2]==='TRUE'||data[i][2]===true,
        canViewAllKRA:    data[i][3]==='TRUE'||data[i][3]===true,
        canCreateTasks:   data[i][4]==='TRUE'||data[i][4]===true
      };
    }
  }
  return { canViewTeamTasks:false, canApprove:false, canViewAllKRA:false, canCreateTasks:false };
}

// ── TASKS ─────────────────────────────────────────────────────────────────

function getMyTasks() {
  const email = Session.getActiveUser().getEmail();
  const raw   = getAllTasksRaw_();
  const result = { oneTime:[], recurring:[], workflow:[], evidencePending:[] };

  raw.forEach(t => {
    if (t.assignedTo !== email && t.assignedBy !== email) return;
    if (t.status === 'Active-Template') return;

    const item = formatTaskForUI_(t);

    // Evidence-pending tasks where YOU are the supervisor (assignedBy)
    if (t.status === 'Evidence-Pending' && t.assignedBy === email) {
      result.evidencePending.push(item);
      return;
    }

    // Only show tasks directly assigned to me in the main lists
    if (t.assignedTo !== email) return;

    if (t.type === 'Recurring') {
      // Show instances (have a parent ID) not templates
      if (t.recurParentId && t.recurParentId !== '') result.recurring.push(item);
    } else if (t.type === 'Workflow') {
      item.isActionable = (t.status === 'Pending');
      result.workflow.push(item);
    } else {
      result.oneTime.push(item);
    }
  });

  const sortFn = (a,b) => {
    const o = {Pending:0,InProgress:1,Late:2,'Evidence-Pending':3,Missed:4,Completed:5,Approved:5,Rejected:6};
    return ((o[a.status]||9)-(o[b.status]||9)) || (new Date(a.rawDueDate||0)-new Date(b.rawDueDate||0));
  };
  result.oneTime.sort(sortFn);
  result.recurring.sort(sortFn);
  result.workflow.sort(sortFn);
  result.evidencePending.sort(sortFn);
  return result;
}

function getTeamTasks() {
  const email = Session.getActiveUser().getEmail();
  const emp   = getEmployeeByEmail_(email);
  if (!emp) return { error:'Not registered.' };
  const perms = getRolePermissions_(emp.role);
  if (!perms.canViewTeamTasks) return { error:'Access denied.' };

  const tasks  = getAllTasksRaw_();
  const emps   = getEmployeesForTeamView_(emp, perms);
  const byEmp  = {};
  emps.forEach(e => { byEmp[e.email] = { ...e, tasks:[] }; });

  tasks.forEach(t => {
    if (!byEmp[t.assignedTo] || t.status === 'Active-Template') return;
    byEmp[t.assignedTo].tasks.push(formatTaskForUI_(t));
  });
  return { team: Object.values(byEmp) };
}

function getEmployeesForTeamView_(callerEmp, perms) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES).getDataRange().getValues();
  const res  = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] !== 'Active') continue;
    if (perms.canViewAllKRA || data[i][4] === callerEmp.department)
      res.push({ empId:data[i][0], name:data[i][1], email:data[i][2], role:data[i][3], department:data[i][4] });
  }
  return res;
}

// ── KRA + KPI ─────────────────────────────────────────────────────────────

function getKRAData() {
  const email = Session.getActiveUser().getEmail();
  const emp   = getEmployeeByEmail_(email);
  if (!emp) return { error:'Not registered.' };
  const perms = getRolePermissions_(emp.role);
  const kpis  = getKPITargets_();

  if (perms.canViewAllKRA) {
    return { scope:'all', scores:getAllKRAScores_(), deptScores:getDeptKRAScores_(), kpis };
  }
  return { scope:'self', score:getMyKRA_(email), deptScore:getMyDeptKRA_(emp.department), kpis };
}

function getKPITargets_() {
  const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG).getDataRange().getValues();
  let onTimeTarget = 85, completionTarget = 90;
  for (let i = 1; i < cfg.length; i++) {
    if (cfg[i][0]==='KPI_Target_OnTime')    onTimeTarget    = parseFloat(cfg[i][1])||85;
    if (cfg[i][0]==='KPI_Target_Completion') completionTarget = parseFloat(cfg[i][1])||90;
  }
  return { onTimeTarget, completionTarget };
}

function getMyDeptKRA_(department) {
  return (getDeptKRAScores_()).find(d => d.department === department) || null;
}

// ── TASK MUTATIONS ────────────────────────────────────────────────────────

function updateTaskStatus(taskId, remarks) {
  const email = Session.getActiveUser().getEmail();
  const emp   = getEmployeeByEmail_(email);
  if (!emp) return { success:false, message:'Not registered.' };

  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(taskId)) continue;
    const assignedTo = data[i][TASK_COL.AssignedTo-1];
    const perms = getRolePermissions_(emp.role);
    if (assignedTo !== email && !perms.canApprove)
      return { success:false, message:'You are not the assignee for this task.' };
    const result = completeTask(taskId, email, remarks||'', null);
    if (result.success) {
      const assignedBy = data[i][TASK_COL.AssignedBy-1];
      addNotification_(assignedBy, '✅ Task completed: '+data[i][1], emp.name+' marked this complete. '+remarks);
    }
    return result;
  }
  return { success:false, message:'Task not found.' };
}

function rejectWorkflowStep(taskId, remarks) {
  const email = Session.getActiveUser().getEmail();
  const emp   = getEmployeeByEmail_(email);
  if (!emp) return { success:false, message:'Not registered.' };
  const perms = getRolePermissions_(emp.role);
  if (!perms.canApprove) return { success:false, message:'Your role cannot reject steps.' };

  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(taskId)) continue;
    sh.getRange(i+1, TASK_COL.Status).setValue('Rejected');
    sh.getRange(i+1, TASK_COL.Remarks).setValue('REJECTED: '+(remarks||'No reason'));
    sh.getRange(i+1, TASK_COL.CompletedDate).setValue(new Date());
    logTaskHistory_(taskId, email, 'Reject', data[i][TASK_COL.Status-1], 'Rejected', false);
    addNotification_(data[i][TASK_COL.AssignedBy-1], '❌ Workflow rejected: '+data[i][1], 'Rejected by '+emp.name+'. Reason: '+remarks);
    return { success:true, message:'Step rejected. Workflow stopped.' };
  }
  return { success:false, message:'Task not found.' };
}

/**
 * Supervisor approves evidence — called from UI with evidence URL.
 */
function supervisorApproveEvidence(taskId, evidenceUrl) {
  const email = Session.getActiveUser().getEmail();
  const emp   = getEmployeeByEmail_(email);
  if (!emp) return { success:false, message:'Not registered.' };

  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(taskId)) continue;
    // Only supervisor (assignedBy) or admin can approve evidence
    const perms = getRolePermissions_(emp.role);
    if (data[i][TASK_COL.AssignedBy-1] !== email && !perms.canApprove)
      return { success:false, message:'Only the task supervisor can approve evidence.' };
    return approveEvidence(taskId, email, evidenceUrl);
  }
  return { success:false, message:'Task not found.' };
}

function createTask(payload) {
  const email = Session.getActiveUser().getEmail();
  const emp   = getEmployeeByEmail_(email);
  if (!emp) return { success:false, message:'Not registered.' };
  const perms = getRolePermissions_(emp.role);
  if (!perms.canCreateTasks) return { success:false, message:'Your role cannot create tasks.' };

  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const data = sh.getDataRange().getValues();
  const maxId = getMaxTaskIdNumber_(data);
  const newId = 'TSK' + String(maxId+1).padStart(4,'0');

  let assignTo = payload.assignedTo, wfStep = '';
  if (payload.type === 'Workflow' && payload.workflowId) {
    const wf = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_WORKFLOWS).getDataRange().getValues();
    for (let i = 1; i < wf.length; i++) {
      if (wf[i][0]===payload.workflowId && Number(wf[i][2])===1) {
        assignTo = wf[i][4]; wfStep = 1; break;
      }
    }
  }

  const newRow = [
    newId, payload.title, payload.type, assignTo, email,
    payload.department||emp.department, new Date(), new Date(payload.dueDate),
    'Pending', '', payload.priority||'Medium',
    payload.recurrenceRule||'', '', payload.workflowId||'', wfStep,
    payload.remarks||'', payload.requiresEvidence?'YES':'NO', '', '', ''
  ];
  sh.appendRow(newRow);
  logTaskHistory_(newId, assignTo, 'Create', '', 'Pending', null);
  addNotification_(assignTo, '📋 New task: '+payload.title, 'Assigned by '+emp.name+'. Due: '+payload.dueDate+(payload.requiresEvidence?' [Evidence required]':''));
  return { success:true, taskId:newId, message:'Task created: '+newId };
}

// ── DEPARTMENT CRUD ───────────────────────────────────────────────────────

function getDepartments() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DEPT).getDataRange().getValues();
  const res  = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    res.push({ deptId:data[i][0], name:data[i][1], headEmail:data[i][2], description:data[i][3], status:data[i][4] });
  }
  return res;
}

function saveDepartment(payload) {
  const email = Session.getActiveUser().getEmail();
  const emp   = getEmployeeByEmail_(email);
  if (!emp) return { success:false, message:'Not registered.' };
  const perms = getRolePermissions_(emp.role);
  if (!perms.canCreateTasks) return { success:false, message:'Access denied.' };

  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DEPT);
  const data = sh.getDataRange().getValues();

  if (payload.deptId) {
    // Update existing
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.deptId) {
        sh.getRange(i+1,1,1,5).setValues([[payload.deptId, payload.name, payload.headEmail, payload.description, payload.status||'Active']]);
        return { success:true, message:'Department updated.' };
      }
    }
  }
  // New department
  const maxNum = data.slice(1).reduce((m,r) => {
    const n = parseInt(String(r[0]).replace('DEPT',''))||0; return Math.max(m,n);
  },0);
  const newId = 'DEPT' + String(maxNum+1).padStart(3,'0');
  sh.appendRow([newId, payload.name, payload.headEmail, payload.description, payload.status||'Active']);
  return { success:true, deptId:newId, message:'Department added.' };
}

// ── EMPLOYEE CRUD ─────────────────────────────────────────────────────────

function getEmployees() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES).getDataRange().getValues();
  const res  = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    res.push({ empId:data[i][0], name:data[i][1], email:data[i][2], role:data[i][3],
               department:data[i][4], supervisorEmail:data[i][5], status:data[i][6],
               joinDate:data[i][7], phone:data[i][8]||'', notes:data[i][9]||'' });
  }
  return res;
}

function saveEmployee(payload) {
  const email = Session.getActiveUser().getEmail();
  const emp   = getEmployeeByEmail_(email);
  if (!emp) return { success:false, message:'Not registered.' };
  const perms = getRolePermissions_(emp.role);
  if (!perms.canCreateTasks) return { success:false, message:'Access denied.' };

  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
  const data = sh.getDataRange().getValues();

  const row = [
    payload.empId||'', payload.name, payload.email, payload.role,
    payload.department, payload.supervisorEmail||'', payload.status||'Active',
    payload.joinDate||'', payload.phone||'', payload.notes||''
  ];

  if (payload.empId) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.empId) {
        sh.getRange(i+1,1,1,10).setValues([row]);
        return { success:true, message:'Employee updated.' };
      }
    }
  }
  // New employee — generate ID
  const maxNum = data.slice(1).reduce((m,r)=>{
    const n=parseInt(String(r[0]).replace('EMP',''))||0; return Math.max(m,n);
  },0);
  row[0] = 'EMP' + String(maxNum+1).padStart(3,'0');
  sh.appendRow(row);
  return { success:true, empId:row[0], message:'Employee added: '+row[0] };
}

// ── ANALYTICS ─────────────────────────────────────────────────────────────

function getAnalytics() {
  const email = Session.getActiveUser().getEmail();
  const emp   = getEmployeeByEmail_(email);
  if (!emp) return { error:'Not registered.' };
  const perms = getRolePermissions_(emp.role);
  const tasks = getAllTasksRaw_().filter(t => t.status !== 'Active-Template');
  const mine  = perms.canViewAllKRA ? tasks : tasks.filter(t => t.assignedTo === email);

  const statusCount={}, typeCount={}, priorityCount={}, deptBreakdown={};
  mine.forEach(t => {
    statusCount[t.status]   = (statusCount[t.status]||0)+1;
    typeCount[t.type]       = (typeCount[t.type]||0)+1;
    priorityCount[t.priority]=(priorityCount[t.priority]||0)+1;
    if (perms.canViewAllKRA) {
      if (!deptBreakdown[t.department]) deptBreakdown[t.department]={total:0,done:0,late:0};
      deptBreakdown[t.department].total++;
      if (['Completed','Approved'].includes(t.status)) deptBreakdown[t.department].done++;
      if (['Late','Missed'].includes(t.status))        deptBreakdown[t.department].late++;
    }
  });

  const today = new Date();
  const trend = [];
  for (let d=6; d>=0; d--) {
    const day = new Date(today); day.setDate(day.getDate()-d); day.setHours(0,0,0,0);
    const label = Utilities.formatDate(day, Session.getScriptTimeZone(), 'dd MMM');
    const completed = mine.filter(t => {
      if (!t.rawCompletedDate) return false;
      const cd = new Date(t.rawCompletedDate); cd.setHours(0,0,0,0);
      return cd.getTime()===day.getTime();
    }).length;
    trend.push({ day:label, completed });
  }

  const total      = mine.length;
  const done       = (statusCount['Completed']||0)+(statusCount['Approved']||0);
  const overdue    = (statusCount['Late']||0)+(statusCount['Missed']||0);
  const completion = total>0 ? Math.round(done/total*100) : 0;
  return { total, done, overdue, completion, statusCount, typeCount, priorityCount, trend, deptBreakdown, isAdmin:perms.canViewAllKRA };
}

// ── DASHBOARD SUMMARY ─────────────────────────────────────────────────────

function getDashboardSummary() {
  const email  = Session.getActiveUser().getEmail();
  const tasks  = getAllTasksRaw_().filter(t => t.assignedTo===email && t.status!=='Active-Template');
  const notifs = getMyNotifications();
  const evidPending = getAllTasksRaw_().filter(t =>
    t.assignedBy===email && t.status==='Evidence-Pending').length;
  return {
    pending:    tasks.filter(t=>t.status==='Pending').length,
    overdue:    tasks.filter(t=>['Late','Missed'].includes(t.status)).length,
    done:       tasks.filter(t=>['Completed','Approved'].includes(t.status)).length,
    wfPending:  tasks.filter(t=>t.type==='Workflow'&&t.status==='Pending').length,
    evidPending: evidPending,
    unreadNotifs: notifs.filter(n=>!n.isRead).length
  };
}

// ── NOTIFICATIONS ─────────────────────────────────────────────────────────

function addNotification_(toEmail, title, body) {
  if (!toEmail) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Notifications');
  if (!sh) {
    sh = ss.insertSheet('Notifications');
    sh.getRange(1,1,1,6).setValues([['NotifID','ToEmail','Title','Body','Timestamp','IsRead']])
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  sh.appendRow(['NTF'+new Date().getTime(), toEmail, title, body, new Date(), 'FALSE']);
}

function getMyNotifications() {
  const email = Session.getActiveUser().getEmail();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();
  return data.slice(1)
    .filter(r => r[1]===email)
    .map(r => ({
      id: r[0], title:r[2], body:r[3],
      time: r[4] ? Utilities.formatDate(new Date(r[4]), tz, 'dd MMM, HH:mm') : '',
      isRead: String(r[5])==='TRUE'
    }))
    .reverse().slice(0,20);
}

function markAllNotificationsRead() {
  const email = Session.getActiveUser().getEmail();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]===email && String(data[i][5])!=='TRUE') sh.getRange(i+1,6).setValue('TRUE');
  }
}

// ── HELPERS ───────────────────────────────────────────────────────────────

function getAllTasksRaw_() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS).getDataRange().getValues();
  return data.slice(1).map(r => ({
    taskId:r[0], title:r[1], type:r[2], assignedTo:r[3], assignedBy:r[4],
    department:r[5], createdDate:r[6], dueDate:r[7], status:r[8],
    completedDate:r[9], priority:r[10], recurrenceRule:r[11],
    recurParentId:String(r[12]||''), workflowId:r[13], workflowStep:r[14],
    remarks:r[15], requiresEvidence:r[16], evidenceUrl:r[17],
    evidenceApprovedBy:r[18], rawDueDate:r[7], rawCompletedDate:r[9]
  }));
}

function formatTaskForUI_(t) {
  const now = new Date(); now.setHours(0,0,0,0);
  const due = t.dueDate ? new Date(t.dueDate) : null;
  if (due) due.setHours(0,0,0,0);
  const tz  = Session.getScriptTimeZone();
  return {
    taskId:t.taskId, title:t.title, type:t.type,
    assignedTo:t.assignedTo, assignedBy:t.assignedBy, department:t.department,
    dueDate: t.dueDate ? Utilities.formatDate(new Date(t.dueDate),tz,'dd MMM yyyy') : '—',
    rawDueDate:t.rawDueDate, status:t.status, priority:t.priority,
    recurrenceRule:t.recurrenceRule, workflowId:t.workflowId, workflowStep:t.workflowStep,
    remarks:t.remarks, requiresEvidence:String(t.requiresEvidence||'').toUpperCase()==='YES',
    evidenceUrl:t.evidenceUrl||'', evidenceApprovedBy:t.evidenceApprovedBy||'',
    isOverdue: !!(due && due < now && ['Pending','InProgress'].includes(t.status)),
    completedDate: t.completedDate ? Utilities.formatDate(new Date(t.completedDate),tz,'dd MMM yyyy') : '',
    rawCompletedDate:t.rawCompletedDate, recurParentId:t.recurParentId
  };
}

function getWorkflowList() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_WORKFLOWS).getDataRange().getValues();
  const seen={}, res=[];
  for (let i=1;i<data.length;i++) {
    if (!seen[data[i][0]]) { seen[data[i][0]]=true; res.push({id:data[i][0],name:data[i][1]}); }
  }
  return res;
}

function getEmployeesByDepartment(department) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES).getDataRange().getValues();
  const res  = [];
  for (let i=1;i<data.length;i++) {
    if (data[i][6]!=='Active') continue;
    if (!department || department==='' || data[i][4]===department)
      res.push({ empId:data[i][0], name:data[i][1], email:data[i][2], department:data[i][4] });
  }
  return res;
}

function getEmployeeList() {
  return getEmployeesByDepartment('');
}
