// TasksService.gs

// ───────────────────────────────────────────────────────────────────────────────
// Constants (adjust sheet names and folder ID to match your setup)
//── 1. Core Read APIs ──────────────────────────────────────────────────────
function getAllTasks() {
  return readCampaignSheet(TASKS_SHEET).map(r => ({
    ID: r.ID,
    Task: r.Task,
    Owner: r.Owner,
    Delegations: r.Delegations,
    StartDate: r.StartDate,
    EndDate: r.EndDate,
    DueDate: r.DueDate,
    RecurrenceRule: r.RecurrenceRule,
    Dependencies: JSON.parse(r.Dependencies || '[]'),
    Priority: r.Priority,
    Calendar: r.Calendar,
    Status: r.Status,
    NotifyOnComplete: JSON.parse(r.NotifyOnComplete || '[]'),
    ApprovalStatus: r.ApprovalStatus,
    SharedLink: r.SharedLink,
    Notes: r.Notes,
    Attachments: r.Attachments ? JSON.parse(r.Attachments) : [],
    CreatedAt: r.CreatedAt,
    UpdatedAt: r.UpdatedAt
  }));
}

function getTaskById(id) {
  const t = getAllTasks().find(t => t.ID === id);
  if (!t) throw new Error('Task not found: ' + id);
  return t;
}

//── 2. Create / Update ────────────────────────────────────────────────────
function addOrUpdateTask(form) {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(TASKS_SHEET);
  const now = new Date().toISOString();
  const isNew = !form.id;
  const id = isNew ? Utilities.getUuid() : form.id;

  const rowObj = {
    ID: id,
    Task: form.task,
    Owner: form.owner || Session.getActiveUser().getEmail(),  // fallback to email
    Delegations: form.delegations || '',
    StartDate: form.startDate || '',
    StartTime: form.startTime || '',
    EndDate: form.endDate || '',
    EndTime: form.endTime || '',
    AllDay: form.allDay ? 'Yes' : 'No',
    Priority: form.priority,
    Calendar: form.calendar,
    Category: form.category || '',
    Status: form.status,
    Notes: form.notes || '',
    CompletedDate: form.status === 'Completed' ? now.slice(0, 10) : '',
    CreatedAt: isNew ? now : form.createdAt,
    UpdatedAt: now
  };

  // — handle attachments —
  if (form.attachments) {
    const blobs = Array.isArray(form.attachments)
                ? form.attachments
                : [form.attachments];
    const folder = DriveApp.getFolderById(FOLDER_ID)
                          .createFolder(id + '-attachments');
    const urls = blobs.map(b => {
      const f = folder.createFile(b.setName(b.name));
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return f.getUrl();
    });
    rowObj.Attachments = JSON.stringify(urls);
  }

  // — write to sheet —
  const data = sh.getDataRange().getValues();
  const hdrs = data.shift();
  const values = hdrs.map(h => rowObj[h] || '');

  if (isNew) {
    sh.appendRow(values);
  } else {
    const idCol = hdrs.indexOf('ID');
    for (let i = 0; i < data.length; i++) {
      if (data[i][idCol] === form.id) {
        sh.getRange(i + 2, 1, 1, values.length).setValues([values]);
        break;
      }
    }
  }

  // sync calendar if needed
  syncTaskToCalendar(getTaskById(id));
}

//── 3. Delete ─────────────────────────────────────────────────────────────
function deleteTask(id) {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(TASKS_SHEET);
  const data = sh.getDataRange().getValues();
  const col = data[0].indexOf('ID');
  for (let i = 1; i < data.length; i++) {
    if (data[i][col] === id) {
      sh.deleteRow(i + 1);
      return;
    }
  }
  throw new Error('Task not found: ' + id);
}

//── 4. EOD Reporting ──────────────────────────────────────────────────────
function getEODTasks(dateStr) {
  const sh = getIBTRSpreadsheet().getSheetByName(TASKS_SHEET);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const hdr = data.shift().map(String);
  return data
    .map(r => hdr.reduce((o, h, i) => (o[h] = r[i], o), {}))
    .filter(t =>
      (t.Status === 'Done' || t.Status === 'Completed')
      && Utilities.formatDate(
        new Date(t.CompletedDate),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      ) === dateStr
    );
}

function sendEODEmail() {
  const date = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
  const tasks = getEODTasks(date);
  const tpl = HtmlService.createTemplateFromFile('EODReport');
  tpl.reportDate = date;
  tpl.tasks = JSON.stringify(tasks).replace(/<\/script>/g, '<\\/script>');
  const htmlBody = tpl.evaluate().getContent();

  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: `EOD Tasks Report — ${date}`,
    htmlBody
  });
}

function createEODTrigger() {
  ScriptApp.newTrigger('sendEODEmail')
    .timeBased().everyDays(1).atHour(18).create();
}

//── 5. Notifications ─────────────────────────────────────────────────────
function getTaskNotifications() {
  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
  return getAllTasks()
    .filter(t =>
      (t.Status !== 'Done' && t.Status !== 'Completed') || t.DueDate === today
    ).length;
}
//── 6. Calendar Sync & Snooze ────────────────────────────────────────────
function syncTaskToCalendar(task) {
  if (!task.Calendar || !task.DueDate) return;
  const cal = CalendarApp.getCalendarById(task.Calendar);
  const start = task.StartDate ? new Date(task.StartDate) : new Date(task.DueDate);
  const end = task.EndDate ? new Date(task.EndDate) : new Date(task.DueDate);

  // remove old events
  cal.getEvents(start, end)
     .filter(e => e.getDescription().includes(task.ID))
     .forEach(e => e.deleteEvent());

  // create new
  cal.createAllDayEvent(task.Task, new Date(task.DueDate), {
    description: `TaskID:${task.ID}`,
    guests: task.Delegations.split(',').filter(u => u)
  });
}

function snoozeTask(id, newDueDate) {
  const t = getTaskById(id);
  t.DueDate = newDueDate;
  addOrUpdateTask({ ...t, id });

  // ↪ Log that the user snoozed this Task
  addOrUpdateTask({
    id: Utilities.getUuid(),
    task: `Reminder extended for “${t.Task}”`,
    owner: t.Owner,
    dueDate: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    ),
    priority: 'Low',
    status: 'To Do',
    notes: `New due date: ${newDueDate}`
  });
}

//── 7. Recurring Engine (daily trigger) ──────────────────────────────────
function generateRecurringTasks() {
  const tasks = getAllTasks();
  tasks.forEach(t => {
    if (!t.RecurrenceRule) return;
    const parts = t.RecurrenceRule.split(';').reduce((m, p) => {
      const [k, v] = p.split('='); m[k] = v; return m;
    }, {});
    const today = new Date();
    let should = false;

    if (parts.FREQ === 'DAILY') {
      should = true;
    } else if (parts.FREQ === 'WEEKLY' && parts.BYDAY) {
      const wd = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'][today.getDay()];
      should = parts.BYDAY.split(',').includes(wd);
    }

    if (should) {
      const instanceId = `${t.ID}#${Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMdd')}`;
      try {
        getTaskById(instanceId);
      } catch (e) {
        addOrUpdateTask({
          id: instanceId,
          task: t.Task,
          owner: t.Owner,
          dueDate: Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          priority: t.Priority,
          status: 'To Do',
          recurrenceRule: '',
          dependencies: t.Dependencies,
          notifyOnComplete: t.NotifyOnComplete,
          approvalStatus: 'Pending',
          sharedLink: '',
          notes: `(Recurring from ${t.ID})`
        });
      }
    }
  });
}

//── 8. Completion Hooks ───────────────────────────────────────────────────
function onTaskComplete(task) {
  (task.NotifyOnComplete || []).forEach(n => {
    if (n.action === 'createTask') {
      addOrUpdateTask({ task: n.template, owner: n.target });
    } else if (n.action === 'email') {
      MailApp.sendEmail(n.target, `Task Completed: ${task.Task}`, `Task ${task.ID} is done.`);
    }
  });
}

//── 9. Comments & Approval ────────────────────────────────────────────────
function addComment(taskId, author, text) {
  const sh = getIBTRSpreadsheet().getSheetByName(COMMENTS_SHEET);
  sh.appendRow([taskId, new Date().toISOString(), author, text]);

  // ↪ Notify the Task owner there’s a new comment
  const t = getTaskById(taskId);
  addOrUpdateTask({
    id: Utilities.getUuid(),
    task: `New comment on “${t.Task}”`,
    owner: t.Owner,
    dueDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    priority: 'Normal',
    status: 'To Do',
    notes: `Comment by ${author}: ${text}`
  });
}

function getComments(taskId) {
  return getIBTRSpreadsheet()
    .getSheetByName(COMMENTS_SHEET)
    .getDataRange().getValues()
    .filter(r => r[0] === taskId)
    .map(r => ({ timestamp: r[1], author: r[2], text: r[3] }));
}

function setApproval(taskId, status, approver) {
  const t = getTaskById(taskId);
  t.ApprovalStatus = status;
  addOrUpdateTask({ ...t, id: taskId });
  MailApp.sendEmail(t.Owner, `Your Task “${t.Task}” was ${status}`, `Approved by ${approver}.`);

  // ↪ Enqueue a “see approval” Task for the owner
  addOrUpdateTask({
    id: Utilities.getUuid(),
    task: `Review approval for “${t.Task}”`,
    owner: t.Owner,
    dueDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    priority: 'Low',
    status: 'To Do',
    notes: `Approved by ${approver}`
  });
}

//── 10. Reporting & Export ────────────────────────────────────────────────
function getBurnDownData(projectId, weekStart) {
  const msDay = 1000 * 60 * 60 * 24;
  const start = new Date(weekStart);
  let remaining = getAllTasks().filter(t => t.ID.startsWith(projectId)).length;
  return Array.from({ length: 7 }).map((_, i) => {
    const date = new Date(start.getTime() + i * msDay);
    const doneCount = getAllTasks().filter(t =>
      t.ID.startsWith(projectId) &&
      (t.Status === 'Done' || t.Status === 'Completed') &&
      t.CompletedDate &&
      t.CompletedDate.startsWith(Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'))
    ).length;
    remaining -= doneCount;
    return {
      date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      remaining
    };
  });
}

// 1) Pull all events from all calendars in the given range
function getAllCalendarsEvents(startIso, endIso) {
  const start = new Date(startIso);
  const end = new Date(endIso);
  return CalendarApp.getAllCalendars().flatMap(cal =>
    cal.getEvents(start, end).map(e => ({
      calendarName: cal.getName(),
      id: e.getId(),
      title: e.getTitle(),
      startTime: e.getStartTime().toISOString(),
      endTime: e.getEndTime().toISOString(),
      allDay: e.isAllDayEvent()
    }))
  );
}

/**
 * Server‐side: fetch Calendar events + all Tasks from every list
 */
function getTasksAndEvents() {
  const CALENDAR_ID = 'primary';
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 2, 0);

  // ── 1) Calendar events ───────────────────
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const events = (cal ? cal.getEvents(start, end) : [])
    .map(ev => ({
      id: ev.getId(),
      title: ev.getTitle(),
      startTime: ev.getStartTime().toISOString(),
      endTime: ev.getEndTime().toISOString(),
      allDay: ev.isAllDayEvent(),
      calendarName: cal.getName()
    }));

  // ── 2) Google Tasks ──────────────────────
  let tasks = [];
  if (typeof Tasks !== 'undefined') {
    const lists = Tasks.Tasklists.list().items || [];
    lists.forEach(list => {
      const items = (Tasks.Tasks.list(list.id).items || []);
      items.forEach(t => {
        if (t.due) {
          const dueISO = new Date(t.due).toISOString();
          tasks.push({
            id: t.id,
            title: t.title,
            startTime: dueISO,
            endTime: dueISO,
            allDay: true,
            calendarName: list.title
          });
        }
      });
    });
  }
  return events.concat(tasks);
}

function exportTasks(format = 'CSV') {
  const rows = getAllTasks().map(t => Object.values(t));
  const csv = rows.map(r => r.join(',')).join('\n');
  if (format === 'CSV') {
    return DriveApp.createFile('Tasks.csv', csv, 'text/csv').getUrl();
  }
  // otherwise, render HTML → PDF in your own template…
}
