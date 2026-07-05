/**
 * FILE: 02_DemoData.gs — v2.1
 * KEY FIX: Active workflow task (Purchase Request step 2) is now
 * assigned to YOUR email so it shows up immediately in the dashboard.
 * All tasks also use CreatedDate within last 30 days for KRA visibility.
 */

function loadDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearDemoSheets_(ss);
  loadDemoDepartments_(ss);
  loadDemoEmployees_(ss);
  loadDemoRoles_(ss);
  loadDemoWorkflows_(ss);
  loadDemoTasks_(ss);
  SpreadsheetApp.getUi().alert('Demo Data Loaded',
    'All demo data loaded fresh.\n' +
    'Your email is mapped to "Anita Sharma" (Coordinator role with full permissions).\n' +
    'You have pending tasks AND a workflow step awaiting your approval.\n\n' +
    'Next: MIS System → Run All Automation Now → then deploy/refresh the web app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function clearDemoSheets_(ss) {
  [SHEET_DEPT, SHEET_EMPLOYEES, SHEET_ROLES, SHEET_WORKFLOWS, SHEET_TASKS].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1, sh.getLastColumn()).clearContent();
  });
}

function loadDemoDepartments_(ss) {
  const sh = ss.getSheetByName(SHEET_DEPT);
  const myEmail = Session.getActiveUser().getEmail() || 'you@example.com';
  sh.getRange(2,1,4,5).setValues([
    ['DEPT001','Academics',   myEmail,               'All academic staff and teachers','Active'],
    ['DEPT002','Accounts',    'karan.mehta@example.com', 'Finance and accounts team',  'Active'],
    ['DEPT003','IT',          'vikram.singh@example.com','IT support and infrastructure','Active'],
    ['DEPT004','HR',          'deepa.iyer@example.com',  'Human resources',            'Active']
  ]);
}

function loadDemoEmployees_(ss) {
  const sh = ss.getSheetByName(SHEET_EMPLOYEES);
  const myEmail = Session.getActiveUser().getEmail() || 'you@example.com';
  // YOU are the Coordinator with full permissions — so you see team tasks, workflow approvals, KRA
  sh.getRange(2,1,8,10).setValues([
    ['EMP001','Anita Sharma (You)', myEmail,                    'Coordinator','Academics','sunil.rao@example.com',  'Active','2022-06-01','9800000001','Demo user - your email'],
    ['EMP002','Rohit Verma',       'rohit.verma@example.com',  'Teacher',    'Academics', myEmail,                 'Active','2021-04-15','9800000002','Reports to you'],
    ['EMP003','Neha Joshi',        'neha.joshi@example.com',   'Teacher',    'Academics', myEmail,                 'Active','2022-09-01','9800000003','Reports to you'],
    ['EMP004','Karan Mehta',       'karan.mehta@example.com',  'Accountant', 'Accounts', 'sunil.rao@example.com',  'Active','2020-08-20','9800000004',''],
    ['EMP005','Sunil Rao',         'sunil.rao@example.com',    'Principal',  'Management','',                      'Active','2015-03-01','9800000005','Top level'],
    ['EMP006','Deepa Iyer',        'deepa.iyer@example.com',   'HR',         'HR',       'sunil.rao@example.com',  'Active','2018-11-05','9800000006',''],
    ['EMP007','Vikram Singh',      'vikram.singh@example.com', 'IT Support', 'IT',       'sunil.rao@example.com',  'Active','2023-02-14','9800000007',''],
    ['EMP008','Priya Menon',       'priya.menon@example.com',  'Teacher',    'Academics', myEmail,                 'Active','2019-01-10','9800000008','Reports to you']
  ]);
}

function loadDemoRoles_(ss) {
  const sh = ss.getSheetByName(SHEET_ROLES);
  sh.getRange(2,1,6,6).setValues([
    ['Teacher',     'FALSE','FALSE','FALSE','FALSE','Sees only own tasks and own KRA'],
    ['Coordinator', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'Department head — team tasks, workflow approval, KRA'],
    ['Accountant',  'FALSE','FALSE','FALSE','FALSE','Sees only own tasks and own KRA'],
    ['Principal',   'TRUE', 'TRUE', 'TRUE', 'TRUE', 'Full visibility across all departments'],
    ['HR',          'TRUE', 'TRUE', 'TRUE', 'TRUE', 'Sees all KRA scores for performance review'],
    ['IT Support',  'FALSE','FALSE','FALSE','FALSE','Sees only own tasks and own KRA']
  ]);
}

function loadDemoWorkflows_(ss) {
  const sh = ss.getSheetByName(SHEET_WORKFLOWS);
  sh.getRange(2,1,5,7).setValues([
    ['WF001','Leave Approval',   1,'Accountant', 'karan.mehta@example.com',2,'FALSE'],
    ['WF001','Leave Approval',   2,'Coordinator','__MY_EMAIL__',           '','TRUE'],
    ['WF002','Purchase Request', 1,'Accountant', 'karan.mehta@example.com',2,'FALSE'],
    ['WF002','Purchase Request', 2,'Coordinator','__MY_EMAIL__',           3,'FALSE'],
    ['WF002','Purchase Request', 3,'Principal',  'sunil.rao@example.com',  '','TRUE']
  ]);
  // Replace placeholder with real email
  const myEmail = Session.getActiveUser().getEmail() || 'you@example.com';
  const data = sh.getRange(2,1,5,7).getValues();
  const fixed = data.map(r => r.map(c => c === '__MY_EMAIL__' ? myEmail : c));
  sh.getRange(2,1,5,7).setValues(fixed);
}

function loadDemoTasks_(ss) {
  const sh = ss.getSheetByName(SHEET_TASKS);
  const myEmail = Session.getActiveUser().getEmail() || 'you@example.com';
  const now = new Date();

  function d(daysOffset) {
    const dt = new Date(now);
    dt.setDate(dt.getDate() + daysOffset);
    return dt;
  }

  // Columns: TaskID,Title,Type,AssignedTo,AssignedBy,Department,
  //          CreatedDate,DueDate,Status,CompletedDate,Priority,
  //          RecurrenceRule,RecurrenceParentID,WorkflowID,WorkflowStepNumber,
  //          Remarks,RequiresEvidence,EvidenceURL,EvidenceApprovedBy,EvidenceApprovedDate

  const tasks = [
    // ── YOUR ONE-TIME TASKS (assigned to your email) ──────────────────────
    ['TSK0001','Prepare monthly academic report','OneTime',myEmail,'sunil.rao@example.com','Academics',
     d(-5),d(2),'Pending','','High','','','','','Monthly report for principal','NO','','',''],

    ['TSK0002','Review curriculum for next term','OneTime',myEmail,'sunil.rao@example.com','Academics',
     d(-10),d(-2),'Late','','High','','','','','Overdue — needs immediate attention','NO','','',''],

    ['TSK0003','Submit attendance audit','OneTime',myEmail,'sunil.rao@example.com','Academics',
     d(-20),d(-15),'Completed',d(-16),'Medium','','','','','Submitted on time','NO','','',''],

    ['TSK0004','Conduct lab safety inspection','OneTime',myEmail,'sunil.rao@example.com','Academics',
     d(-3),d(5),'Pending','','Urgent','','','','','Requires physical inspection of all labs','YES','','',''],

    // ── TASKS YOU ASSIGNED TO YOUR TEAM ──────────────────────────────────
    ['TSK0005','Complete unit test papers','OneTime','rohit.verma@example.com',myEmail,'Academics',
     d(-8),d(-3),'Missed','','High','','','','','Not submitted by deadline','NO','','',''],

    ['TSK0006','Update student grade records','OneTime','neha.joshi@example.com',myEmail,'Academics',
     d(-6),d(1),'Pending','','Medium','','','','','','NO','','',''],

    ['TSK0007','Submit classroom observation report','OneTime','priya.menon@example.com',myEmail,'Academics',
     d(-15),d(-10),'Completed',d(-11),'Low','','','','','Completed early','YES','','',''],

    // ── RECURRING TASKS (your email) ─────────────────────────────────────
    ['TSK0008','Mark daily attendance','Recurring',myEmail,'sunil.rao@example.com','Academics',
     d(-30),d(0),'Active-Template','','High','Daily','','','','Recurring template — do not complete directly','NO','','',''],

    ['TSK0009','Mark daily attendance','Recurring',myEmail,'sunil.rao@example.com','Academics',
     d(-2),d(-2),'Completed',d(-2),'High','Daily','TSK0008','','','','NO','','',''],

    ['TSK0010','Mark daily attendance','Recurring',myEmail,'sunil.rao@example.com','Academics',
     d(-1),d(-1),'Completed',d(-1),'High','Daily','TSK0008','','','','NO','','',''],

    ['TSK0011','Mark daily attendance','Recurring',myEmail,'sunil.rao@example.com','Academics',
     d(0),d(0),'Pending','','High','Daily','TSK0008','','','TODAY — mark attendance','NO','','',''],

    ['TSK0012','Submit weekly dept report','Recurring',myEmail,'sunil.rao@example.com','Academics',
     d(-14),d(-7),'Completed',d(-7),'Medium','Weekly','','','','','NO','','',''],

    ['TSK0013','Submit weekly dept report','Recurring',myEmail,'sunil.rao@example.com','Academics',
     d(-7),d(0),'Pending','','Medium','Weekly','TSK0012','','','Due today','NO','','',''],

    // ── EVIDENCE-REQUIRED TASK (completed, waiting for your evidence approval) ─
    ['TSK0014','Fire drill documentation','OneTime','rohit.verma@example.com',myEmail,'Academics',
     d(-4),d(-1),'Evidence-Pending','','High','','','','',
     'Rohit completed the drill but evidence photo not yet uploaded','YES','','',''],

    // ── WORKFLOW: Purchase Request — Step 2 PENDING ON YOUR APPROVAL ─────
    ['TSK0015','Purchase Request — New Lab Equipment (Step 1)','Workflow',
     'karan.mehta@example.com','vikram.singh@example.com','Accounts',
     d(-6),d(-4),'Approved',d(-5),'High','','','WF002',1,
     'Budget verified by accounts dept','NO','','',''],

    ['TSK0016','Purchase Request — New Lab Equipment (Step 2)','Workflow',
     myEmail,'vikram.singh@example.com','Accounts',
     d(-5),d(2),'Pending','','High','','TSK0015','WF002',2,
     'AWAITING YOUR APPROVAL — approve or reject this purchase request','NO','','',''],

    // ── WORKFLOW: Leave Approval — FULLY COMPLETED ────────────────────────
    ['TSK0017','Leave Approval — Rohit Verma 3 days (Step 1)','Workflow',
     'karan.mehta@example.com','rohit.verma@example.com','Accounts',
     d(-10),d(-8),'Approved',d(-9),'Medium','','','WF001',1,
     'Verified leave balance','NO','','',''],

    ['TSK0018','Leave Approval — Rohit Verma 3 days (Step 2)','Workflow',
     myEmail,'rohit.verma@example.com','Academics',
     d(-9),d(-7),'Approved',d(-8),'Medium','','TSK0017','WF001',2,
     'Final approval granted','NO','','',''],

    // ── OTHER DEPT TASKS (for team view and analytics) ────────────────────
    ['TSK0019','Server backup verification','OneTime','vikram.singh@example.com','sunil.rao@example.com','IT',
     d(-3),d(-1),'Completed',d(-2),'Urgent','','','','','','NO','','',''],

    ['TSK0020','Reconcile vendor invoices','OneTime','karan.mehta@example.com','sunil.rao@example.com','Accounts',
     d(-8),d(-3),'Completed',d(-4),'Medium','','','','','','NO','','',''],

    ['TSK0021','Update HR onboarding docs','OneTime','deepa.iyer@example.com','sunil.rao@example.com','HR',
     d(-12),d(-7),'Late',d(-5),'Low','','','','','Late by 2 days','NO','','',''],

    ['TSK0022','Fix projector Room 204','OneTime','vikram.singh@example.com',myEmail,'IT',
     d(-2),d(3),'Pending','','High','','','','','Reported by Anita','YES','','','']
  ];

  sh.getRange(2,1,tasks.length,tasks[0].length).setValues(tasks);
  applyTaskStatusFormatting_(sh);
}
