# MIS Dashboard v2.1 — Complete Deployment Guide

## What's in this package

| File | Purpose | Change |
|------|---------|--------|
| `01_Setup.gs` | Creates all sheets including new Departments sheet, Evidence columns in Tasks | Updated |
| `02_DemoData.gs` | Demo data — YOUR email gets Coordinator role, active workflow step assigned to you | Fixed |
| `03_Automation.gs` | Recurring tasks, overdue detection, workflow advancement, evidence approval | Updated |
| `04_KRA_Engine.gs` | KRA scoring + Department KRA, 90-day window fix | Updated |
| `05_Triggers.gs` | Daily automation triggers | Unchanged |
| `06_WebApp_Backend.gs` | All API functions — evidence, dept CRUD, employee CRUD, KPI, dept-filtered assignees | Updated |
| `07_Dashboard_UI.html` | Full web app UI | Updated |
| `08_Menu.gs` | Sheets menu | Updated |

---

## Fresh install steps (10 minutes)

### Step 1 — Create project
1. Open a **blank Google Sheet**
2. **Extensions → Apps Script**
3. Delete all content from the default `Code.gs`

### Step 2 — Add all files

**For each `.gs` file** (01 through 08):
- Click **+** → **Script**
- Name it exactly as shown (e.g. `01_Setup`) — no `.gs` extension needed
- Paste full file contents → Save

**For `07_Dashboard_UI.html`:**
- Click **+** → **HTML** ← must choose HTML not Script
- Name it exactly `07_Dashboard_UI`
- Paste full file contents → Save

> ⚠️ HTML file name MUST be `07_Dashboard_UI` exactly.
> ⚠️ Never add `<!DOCTYPE html>` or `<?xml …?>` to the HTML file.

### Step 3 — Run setup
1. In the function dropdown select `setupMIS` → **Run**
2. Accept permissions when prompted
3. Wait for "MIS Setup Complete" alert

### Step 4 — Load demo data
1. Go back to your **Google Sheet** (reload the tab)
2. Click **🏫 MIS System → 2. Load / Refresh Demo Data**
3. Wait for confirmation alert

> Your Google account email is automatically set as "Anita Sharma" with **Coordinator** role.
> This gives you full permissions: team tasks, workflow approval, KRA scores, create tasks.
> You will see an active workflow task (Purchase Request Step 2) awaiting YOUR approval.

### Step 5 — Install triggers
- **MIS System → 3. Install Automation Triggers**

### Step 6 — Deploy as Web App
1. Apps Script → **Deploy → New deployment**
2. Click ⚙ gear next to Type → **Web app**
3. Settings:
   - **Execute as:** Me
   - **Who has access:** Anyone with a Google Account
4. Click **Deploy** → copy the Web App URL

### Step 7 — First run checklist
Open the Web App URL and verify:
- ✅ Your name appears in the sidebar
- ✅ Dashboard shows pending/overdue counts
- ✅ "My Tasks" → One-Time tab has tasks
- ✅ "My Tasks" → Workflow tab shows "Purchase Request Step 2" with Approve/Reject buttons
- ✅ "My Tasks" → Recurring tab shows daily attendance instances
- ✅ KRA/KPI page shows your score (click "Calculate My Score" if blank)
- ✅ Team Tasks shows your Academics team
- ✅ Analytics shows charts
- ✅ Employees and Departments pages appear in sidebar (Coordinator role has access)

---

## If you already have data (upgrading from v1/v2)

Only replace these files — do NOT re-run setupMIS (it will skip existing sheets):
- `01_Setup.gs` — adds Evidence columns and Departments sheet
- `02_DemoData.gs` — refreshed demo data
- `03_Automation.gs` — evidence support
- `04_KRA_Engine.gs` — KRA fix + dept KRA
- `06_WebApp_Backend.gs` — all new features
- `07_Dashboard_UI.html` — full UI
- `08_Menu.gs` — updated menu

Then run **MIS System → 2. Load / Refresh Demo Data** to reload fresh demo data.
Then create a **new deployment version** (Deploy → Manage deployments → Edit → New version).

---

## Key features in v2.1

### Evidence / Photo upload flow
1. When creating a task, check **"Requires evidence/photo upload"**
2. Assignee completes task → prompted to paste a Google Drive link
3. Status becomes **Evidence-Pending** instead of Completed
4. Supervisor sees an **Evidence** tab in My Tasks with "Upload Evidence" button
5. Supervisor uploads to Drive, pastes link → task becomes **Completed**

### Department KRA
- Visible on KRA page for Coordinators/Principals/HR
- Shows avg KRA score, on-time %, employee count per department
- Auto-calculated when individual KRA scores are recalculated

### KPI Cards
- **On-Time Rate** vs target (default 85%) — green if hitting, red if not
- **Completion Rate** vs target (default 90%)
- Change targets in the **Config** sheet: `KPI_Target_OnTime` and `KPI_Target_Completion`

### Task assignee filtered by department
- In Create Task: select Department first → Assignee dropdown shows only people in that department
- Reduces wrong-department assignments

### Employee & Department management
- Full add/edit forms accessible from sidebar (Coordinator+ roles)
- Supervisor dropdown filtered by selected department
- Role dropdown with all defined roles

---

## KRA Scoring formula

```
OnTimePct  = CompletedOnTime / TotalAssigned × 100
VolumePct  = min(TotalAssigned / 10, 1) × 100
KRAScore   = (OnTimePct × 0.70) + (VolumePct × 0.30)
```

KRA uses a **90-day rolling window** so demo data always has signal.
Change weights in Config sheet: `KRA_OnTimeWeight` / `KRA_VolumeWeight` (must sum to 1.0).

---

## Role permissions

| Role | Team Tasks | Approve WF | All KRA | Create Tasks | Admin Forms |
|------|:---:|:---:|:---:|:---:|:---:|
| Teacher | ✗ | ✗ | ✗ | ✗ | ✗ |
| Coordinator | ✓ | ✓ | ✓ | ✓ | ✓ |
| Accountant | ✗ | ✗ | ✗ | ✗ | ✗ |
| Principal | ✓ | ✓ | ✓ | ✓ | ✓ |
| HR | ✓ | ✓ | ✓ | ✓ | ✓ |
| IT Support | ✗ | ✗ | ✗ | ✗ | ✗ |

Add new roles by adding rows to the **Roles** sheet — no code change needed.

---

## Common issues

**Dashboard loads but shows no tasks**
→ Run **MIS System → ▶ Run All Automation Now**, then refresh the web app.

**KRA page shows "No KRA data yet"**
→ Click "Calculate My Score" button on the KRA page, or run `recalculateKRAScores()` from Apps Script.

**Workflow tasks not visible**
→ Confirm your email in Employees sheet matches your Google login exactly (case-sensitive).
→ Re-run **Load Demo Data** which maps the active workflow step to your email.

**"Your email is not registered" error**
→ Add your email to the Employees sheet with Status = Active.

**Evidence URL field doesn't appear**
→ Only shows when creating a task with "Requires evidence" checked, OR when completing such a task.

**Web app shows old version after code changes**
→ Always create a **New Version** under Manage Deployments — never just save and refresh.

**Assignee dropdown shows "Select department first"**
→ Select the Department dropdown before trying to pick an assignee — it filters by department.
