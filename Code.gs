// ============================================================
// MIDNIGHT COACH TASK MANAGER - Code.gs
// Pure JSON API — served to GitHub Pages front-end
// All responses go through sendJson() with CORS headers
// ============================================================

// ── CONFIG ───────────────────────────────────────────────────

var PHOTOS_FOLDER_ID = '1_PO2F_bPEsRolfNmIQe7AbrjaydQ5Sdj';

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) throw new Error('Config sheet not found');
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) config[String(data[i][0])] = data[i][1];
  }
  return config;
}

function getListFromSheet(sheetName, col) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .map(r => r[col || 0]).filter(v => v !== '' && v !== null);
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { headers: [], rows: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return { headers: [], rows: [] };
  const headers = data[0].map(h => h.toString());
  const rows = data.slice(1).map((row, i) => {
    const obj = { _rowIndex: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });
  return { headers, rows };
}

// ── CORS & ROUTING ───────────────────────────────────────────

/**
 * All responses go through here to ensure consistent CORS headers.
 * Apps Script ContentService is the only way to set headers.
 */
function sendJson(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET handler — routes based on ?action= parameter
 * Used for read operations and JSONP-style requests
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const callback = (e && e.parameter && e.parameter.callback) || '';

  try {
    let result;

    switch (action) {
      case 'getPeopleForLogin':
        result = getPeopleForLogin();
        break;
      case 'getFormData':
        result = getFormData();
        break;
      case 'getTasks':
        result = getAllTasksUnfiltered();
        break;
      case 'getDashboardData':
        result = getDashboardData();
        break;
      case 'getPeopleAdmin':
        result = getSheetData('People').rows;
        break;
      case 'getFlaggedTasks':
        result = getFlaggedTasks();
        break;
      case 'getTasksWithMissingCategory':
        result = getTasksWithMissingCategory();
        break;
      case 'getBuses':
        result = getSheetData('Buses').rows;
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }

    // Support JSONP for cross-origin GET requests
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify({ success: true, data: result }) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return sendJson({ success: true, data: result });

  } catch (err) {
    Logger.log('doGet error: ' + err);
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify({ success: false, error: err.toString() }) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return sendJson({ success: false, error: err.toString() });
  }
}

/**
 * POST handler — routes based on action in JSON body
 * Used for all write operations
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';
    const userName = body.userName || '';

    let result;

    switch (action) {

      // Auth
      case 'verifyPin':
        result = verifyPin(body.name, body.pin);
        break;

      // Tasks
      case 'saveTask':
        result = saveTask(body.data, userName);
        break;
      case 'updateTask':
        result = updateTask(body.rowIndex, body.data, userName);
        break;
      case 'deleteTask':
        result = deleteTask(body.rowIndex, userName);
        break;
      case 'completeTask':
        result = completeTask(body.rowIndex, body.data, userName);
        break;

      // People
      case 'savePerson':
        result = savePerson(body.data, userName);
        break;
      case 'deletePerson':
        result = deletePerson(body.rowIndex, userName);
        break;

      // Admin
      case 'reassignTasks':
        result = reassignTasks(body.rows, body.newAssignee, userName);
        break;
      case 'recategorizeTasks':
        result = recategorizeTasks(body.rows, body.newCat, userName);
        break;
      case 'sendDailyDigest':
        sendDailyDigest();
        result = { success: true };
        break;

      // Photos
      case 'uploadPhoto':
        result = uploadPhoto(body.taskId, body.fileName, body.base64Data, body.mimeType, userName);
        break;
      case 'deletePhoto':
        result = deletePhoto(body.fileId, userName);
        break;

      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }

    return sendJson(result);

  } catch (err) {
    Logger.log('doPost error: ' + err);
    return sendJson({ success: false, error: err.toString() });
  }
}

// ── AUTH ─────────────────────────────────────────────────────

function verifyPin(name, pin) {
  try {
    if (!name || !pin) return { success: false, error: 'Name and PIN are required.' };
    const { rows } = getSheetData('People');
    const config = getConfig();
    const admins = (config['admin_emails'] || '').split(',').map(s => s.trim().toLowerCase());

    const person = rows.find(p => {
      const active = p['active'];
      if (active === 'No' || active === false || active === 'FALSE' || active === 'no') return false;
      return p['name'] === name;
    });

    if (!person) return { success: false, error: 'User not found.' };
    const storedPin = String(person['pin'] || '').trim();
    if (!storedPin) return { success: false, error: 'No PIN set for this user. Contact an admin.' };
    if (String(pin).trim() !== storedPin) return { success: false, error: 'Incorrect PIN.' };

    const safePerson = {
      name:       person['name'],
      email:      person['email'],
      alt_emails: person['alt_emails'],
      color:      person['color'],
      phone:      person['phone'],
      carrier:    person['carrier'],
      sms_opt_in: person['sms_opt_in'],
      _rowIndex:  person['_rowIndex']
    };

    const isAdmin = admins.includes((person['email'] || '').toLowerCase().trim()) ||
      (person['alt_emails'] || '').split(',').some(e => admins.includes(e.trim().toLowerCase()));

    return { success: true, person: safePerson, isAdmin, name: person['name'] };
  } catch (e) {
    Logger.log('verifyPin: ' + e);
    return { success: false, error: 'Server error. Please try again.' };
  }
}

function getPeopleForLogin() {
  const { rows } = getSheetData('People');
  return rows
    .filter(p => {
      const a = p['active'];
      return a !== 'No' && a !== false && a !== 'FALSE' && a !== 'no';
    })
    .map(p => ({ name: p['name'], color: p['color'] || '#64748b' }));
}

function getVerifiedPerson(name) {
  if (!name) throw new Error('ACCESS_DENIED: No user session.');
  const { rows } = getSheetData('People');
  const person = rows.find(p => {
    const active = p['active'];
    if (active === 'No' || active === false || active === 'FALSE' || active === 'no') return false;
    return p['name'] === name;
  });
  if (!person) throw new Error('ACCESS_DENIED: ' + name + ' is not an authorized user.');
  return person;
}

// ── TASKS ─────────────────────────────────────────────────────

function getAllTasksUnfiltered() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    return data.map((row, i) => {
      const obj = { _rowIndex: i + 2 };
      headers.forEach((h, j) => { obj[h] = row[j]; });
      ['due_date', 'start_date', 'created_at', 'updated_at', 'completed_at'].forEach(f => {
        if (obj[f] instanceof Date)
          obj[f] = Utilities.formatDate(obj[f], Session.getScriptTimeZone(),
            (f === 'due_date' || f === 'start_date') ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm');
      });
      return obj;
    });
  } catch (e) { Logger.log('getAllTasksUnfiltered: ' + e); return []; }
}

function saveTask(d, userName) {
  getVerifiedPerson(userName);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const now = new Date();
    d.created_at = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    d.task_id = 'T-' + now.getTime();
    d.status = d.status || 'To Do';
    sheet.appendRow(headers.map(h => d[h] !== undefined ? d[h] : ''));
    return { success: true, task_id: d.task_id };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateTask(rowIndex, d, userName) {
  getVerifiedPerson(userName);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    d.updated_at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    headers.forEach((h, i) => { if (d[h] !== undefined) sheet.getRange(rowIndex, i + 1).setValue(d[h]); });
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteTask(rowIndex, userName) {
  getVerifiedPerson(userName);
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks').deleteRow(rowIndex);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function completeTask(rowIndex, d, userName) {
  getVerifiedPerson(userName);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const u = { status: 'Done', completed_at: now, updated_at: now };
    headers.forEach((h, i) => { if (u[h] !== undefined) sheet.getRange(rowIndex, i + 1).setValue(u[h]); });
    if (d.recurring && d.recurring !== 'None' && d.due_date) createNextRecurrence(d);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function createNextRecurrence(d) {
  const next = new Date(d.due_date);
  if      (d.recurring === 'Daily')   next.setDate(next.getDate() + 1);
  else if (d.recurring === 'Weekly')  next.setDate(next.getDate() + 7);
  else if (d.recurring === 'Monthly') next.setMonth(next.getMonth() + 1);
  else if (d.recurring === 'Yearly')  next.setFullYear(next.getFullYear() + 1);
  else if (d.recurrence_interval)     next.setDate(next.getDate() + parseInt(d.recurrence_interval));
  else return;
  const t = Object.assign({}, d);
  t.due_date = Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  t.status = 'To Do';
  delete t.completed_at; delete t.task_id; delete t._rowIndex;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const now = new Date();
  t.created_at = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  t.task_id = 'T-' + now.getTime();
  sheet.appendRow(headers.map(h => t[h] !== undefined ? t[h] : ''));
}

// ── PHOTOS ────────────────────────────────────────────────────

function uploadPhoto(taskId, fileName, base64Data, mimeType, userName) {
  getVerifiedPerson(userName);
  try {
    const folder = DriveApp.getFolderById(PHOTOS_FOLDER_ID);

    // Find or create subfolder for this task
    const subFolderName = 'Task_' + taskId;
    let subFolder;
    const existing = folder.getFoldersByName(subFolderName);
    if (existing.hasNext()) {
      subFolder = existing.next();
    } else {
      subFolder = folder.createFolder(subFolderName);
    }

    // Decode base64 and create file
    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);
    const file = subFolder.createFile(blob);

    // Make file publicly readable so it can be displayed in the app
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';
    const viewUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

    // Save photo reference to task
    savePhotoToTask(taskId, fileId, fileName, thumbnailUrl, viewUrl, userName);

    return { success: true, fileId, thumbnailUrl, viewUrl, fileName };
  } catch (e) {
    Logger.log('uploadPhoto: ' + e);
    return { success: false, error: e.toString() };
  }
}

function savePhotoToTask(taskId, fileId, fileName, thumbnailUrl, viewUrl, userName) {
  // Store photo metadata in a Photos sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('TaskPhotos');
  if (!sheet) {
    sheet = ss.insertSheet('TaskPhotos');
    sheet.appendRow(['task_id', 'file_id', 'file_name', 'thumbnail_url', 'view_url', 'uploaded_by', 'uploaded_at']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([taskId, fileId, fileName, thumbnailUrl, viewUrl, userName, now]);
}

function getPhotosForTask(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TaskPhotos');
  if (!sheet) return [];
  const { rows } = getSheetData('TaskPhotos');
  return rows.filter(r => r['task_id'] === taskId);
}

function deletePhoto(fileId, userName) {
  getVerifiedPerson(userName);
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    // Remove from TaskPhotos sheet
    const { rows } = getSheetData('TaskPhotos');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TaskPhotos');
    const row = rows.find(r => r['file_id'] === fileId);
    if (row && sheet) sheet.deleteRow(row._rowIndex);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ── FORM DATA ─────────────────────────────────────────────────

function getFormData() {
  try {
    const config = getConfig();
    return {
      people:     getPeople(),
      buses:      getBusesList(),
      categories: getListFromSheet('Categories', 0),
      statuses:   (config['statuses']         || 'To Do,In Progress,Done,On Hold,Cancelled').split(',').map(s => s.trim()),
      priorities: (config['priorities']        || 'High,Medium,Low').split(',').map(s => s.trim()),
      recurring:  (config['recurring_options'] || 'None,Daily,Weekly,Monthly,Yearly').split(',').map(s => s.trim()),
      config
    };
  } catch (e) {
    Logger.log('getFormData: ' + e);
    return { people: [], buses: [], categories: [], statuses: [], priorities: [], recurring: [], config: {} };
  }
}

function getPeople() {
  const { rows } = getSheetData('People');
  return rows.filter(r => {
    const a = r['active'];
    return a !== false && a !== 'FALSE' && a !== 'No' && a !== 'no';
  });
}

function getBusesList() {
  return getSheetData('Buses').rows.map(r => r['bus_name']).filter(b => b && b !== '');
}

// ── DASHBOARD ─────────────────────────────────────────────────

function getDashboardData() {
  try {
    const allTasks = getAllTasksUnfiltered();
    const people = getPeople();
    const config = getConfig();
    const today = new Date(); today.setHours(0, 0, 0, 0);

    // Filter to only tasks that should be visible
    const tasks = allTasks.filter(t => {
      const sd = t['start_date'];
      if (sd && new Date(sd) > today) return false;
      if (t['status'] === 'Cancelled') return false;
      // Hide Done tasks older than completed_task_days
      if (t['status'] === 'Done') {
        const days = parseInt(config['completed_task_days'] || '3');
        if (t['completed_at']) {
          const completedDate = new Date(t['completed_at']);
          const diffDays = (today - completedDate) / (1000 * 60 * 60 * 24);
          if (diffDays > days) return false;
        }
      }
      return true;
    });

    const po = { 'High': 0, 'Medium': 1, 'Low': 2 };
    const personMap = {};
    people.forEach(p => { personMap[p['name']] = { person: p, tasks: [] }; });

    tasks.forEach(t => {
      const a = t['assigned_to'];
      if (a && personMap[a]) personMap[a].tasks.push(t);
    });

    // Sort each person's tasks
    Object.values(personMap).forEach(e => {
      e.tasks.sort((a, b) => {
        const ad = a['status'] === 'Done', bd = b['status'] === 'Done';
        if (ad && !bd) return 1; if (!ad && bd) return -1;
        const aD = a['due_date'] ? new Date(a['due_date']) : new Date('9999-12-31');
        const bD = b['due_date'] ? new Date(b['due_date']) : new Date('9999-12-31');
        if (aD < bD) return -1; if (aD > bD) return 1;
        return (po[a['priority']] || 99) - (po[b['priority']] || 99);
      });
    });

    // Bus shop status
    const busStatus = {};
    getSheetData('Buses').rows.forEach(r => {
      if (r['bus_name']) busStatus[r['bus_name']] = {
        shopStatus: r['shop_status'] || '',
        expectedReturn: r['expected_return'] || ''
      };
    });

    return {
      personMap,
      peopleOrder: people.map(p => p['name']),
      busStatus,
      lastUpdated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a'),
      config
    };
  } catch (e) {
    Logger.log('getDashboardData: ' + e);
    return { personMap: {}, peopleOrder: [], busStatus: {}, lastUpdated: '', config: {} };
  }
}

// ── PEOPLE MANAGEMENT ─────────────────────────────────────────

function savePerson(d, userName) {
  getVerifiedPerson(userName);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('People');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (d._rowIndex) {
      headers.forEach((h, i) => {
        if (h === 'pin' && (d[h] === undefined || d[h] === '')) return;
        if (d[h] !== undefined) sheet.getRange(d._rowIndex, i + 1).setValue(d[h]);
      });
      return { success: true, action: 'updated' };
    } else {
      sheet.appendRow(headers.map(h => d[h] !== undefined ? d[h] : ''));
      return { success: true, action: 'added' };
    }
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deletePerson(rowIndex, userName) {
  getVerifiedPerson(userName);
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('People').deleteRow(rowIndex);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ── ADMIN UTILITIES ───────────────────────────────────────────

function getFlaggedTasks() {
  const active = getPeople().map(p => p['name']);
  return getAllTasksUnfiltered().filter(t =>
    t['assigned_to'] && !active.includes(t['assigned_to']) &&
    t['status'] !== 'Done' && t['status'] !== 'Cancelled'
  );
}

function reassignTasks(rows, newAssignee, userName) {
  getVerifiedPerson(userName);
  try {
    rows.forEach(i => updateTaskInternal(i, { assigned_to: newAssignee }));
    return { success: true, count: rows.length };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getTasksWithMissingCategory() {
  const cats = getListFromSheet('Categories', 0);
  return getAllTasksUnfiltered().filter(t =>
    t['category'] && !cats.includes(t['category']) && t['status'] !== 'Done'
  );
}

function recategorizeTasks(rows, newCat, userName) {
  getVerifiedPerson(userName);
  try {
    rows.forEach(i => updateTaskInternal(i, { category: newCat }));
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateTaskInternal(rowIndex, d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  d.updated_at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  headers.forEach((h, i) => { if (d[h] !== undefined) sheet.getRange(rowIndex, i + 1).setValue(d[h]); });
}

// ── EMAIL / SMS ───────────────────────────────────────────────

function getSmsCarriers() {
  const c = {};
  getSheetData('SMS_Carriers').rows.forEach(r => {
    if (r['carrier_name'] && r['gateway']) c[r['carrier_name']] = r['gateway'];
  });
  return c;
}

function sendDailyDigest() {
  try {
    const people = getPeople();
    const tasks = getAllTasksUnfiltered();
    const config = getConfig();
    const title = config['app_title'] || 'Task Manager';
    const carriers = getSmsCarriers();
    const today = new Date(); today.setHours(0, 0, 0, 0);

    people.forEach(function(p) {
      const mine = tasks.filter(function(t) {
        return t['assigned_to'] === p['name'] &&
               t['status'] !== 'Done' &&
               t['status'] !== 'Cancelled';
      });
      if (!mine.length || !p['email']) return;

      const ov = mine.filter(t => t['due_date'] && new Date(t['due_date']) < today);
      const dt = mine.filter(t => t['due_date'] && new Date(t['due_date']).toDateString() === today.toDateString());
      const up = mine.filter(t => !t['due_date'] || new Date(t['due_date']) > today);

      let body = '<h2 style="color:#1e40af;">' + title + '</h2>';
      body += '<h3>Tasks for ' + p['name'] + '</h3>';

      if (ov.length) {
        body += '<h4 style="color:#dc2626;">⚠️ Overdue (' + ov.length + ')</h4><ul>';
        ov.forEach(t => { body += '<li><strong>' + t['task_name'] + '</strong> — Due: ' + t['due_date'] + ' [' + t['priority'] + ']' + (t['bus'] ? ' · ' + t['bus'] : '') + '</li>'; });
        body += '</ul>';
      }
      if (dt.length) {
        body += '<h4 style="color:#d97706;">📅 Due Today (' + dt.length + ')</h4><ul>';
        dt.forEach(t => { body += '<li><strong>' + t['task_name'] + '</strong> [' + t['priority'] + ']' + (t['bus'] ? ' · ' + t['bus'] : '') + '</li>'; });
        body += '</ul>';
      }
      if (up.length) {
        body += '<h4 style="color:#16a34a;">📋 Upcoming (' + up.length + ')</h4><ul>';
        up.forEach(t => { body += '<li><strong>' + t['task_name'] + '</strong>' + (t['due_date'] ? ' — Due: ' + t['due_date'] : '') + ' [' + t['priority'] + ']</li>'; });
        body += '</ul>';
      }

      GmailApp.sendEmail(p['email'], title + ' — Your Tasks for Today', '', { htmlBody: body });

      if (p['phone'] && p['carrier'] && p['sms_opt_in'] === 'Yes') {
        const gw = carriers[p['carrier']];
        if (gw) {
          const ph = p['phone'].toString().replace(/\D/g, '');
          GmailApp.sendEmail(ph + '@' + gw, '', title + ': ' + mine.length + ' tasks (' + ov.length + ' overdue, ' + dt.length + ' today)');
        }
      }
    });
  } catch (e) { Logger.log('sendDailyDigest: ' + e); }
}

// ── SETUP ─────────────────────────────────────────────────────

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function mkSheet(name, headers) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name); else s.clearContents();
    s.appendRow(headers);
    s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    return s;
  }

  mkSheet('Tasks', ['task_id','task_name','assigned_to','status','priority','category','bus','due_date','start_date','recurring','recurrence_interval','next_due','notes','created_at','updated_at','completed_at']);

  const ppl = mkSheet('People', ['name','email','alt_emails','pin','phone','carrier','sms_opt_in','color','active']);
  [['David','dwooswald@gmail.com','','1234','7708415038','AT&T','Yes','#3b82f6','Yes'],
   ['Zach','zach@midnightcoach.com','','1234','7706683325','AT&T','Yes','#10b981','Yes'],
   ['Bradley','bradleyedwards7448@gmail.com','','1234','7066699976','AT&T','Yes','#f59e0b','Yes'],
   ['Nich','nich@midnightcoach.com','','1234','7706683338','AT&T','Yes','#8b5cf6','Yes'],
   ['Brody','','','1234','','AT&T','Yes','#14b8a6','Yes'],
   ['Jimmy','cochranjiv@gmail.com','','1234','4706268850','AT&T','Yes','#ef4444','Yes']
  ].forEach(r => ppl.appendRow(r));

  const buses = mkSheet('Buses', ['bus_name','shop_status','expected_return','notes']);
  ['Night Owl','Faith','Storm Front','Midnight Cowboy','Jupiter','Dark Horse','Daylight','Layla','Moonshine','Lakeshore','Sgt Pepper','Ella','Evening Star','Newsong','Solitude','Black Betty','Black Pearl'].forEach(b => buses.appendRow([b,'On Tour','','']));

  const cats = mkSheet('Categories', ['category_name','color']);
  [['Mechanical','#ef4444'],['Interior','#8b5cf6'],['Electrical','#f59e0b'],['Purchasing','#3b82f6'],['Admin','#10b981'],['Safety','#14b8a6'],['DOT Compliance','#f97316'],['General','#6b7280']].forEach(r => cats.appendRow(r));

  const sms = mkSheet('SMS_Carriers', ['carrier_name','gateway']);
  [['AT&T','txt.att.net'],['Verizon','vtext.com'],['T-Mobile','tmomail.net'],['Sprint','messaging.sprintpcs.com'],['US Cellular','email.uscc.net'],['Boost Mobile','myboostmobile.com'],['Cricket','sms.cricketwireless.net']].forEach(r => sms.appendRow(r));

  const cfg = mkSheet('Config', ['key','value','description']);
  [['app_title','Midnight Coach Tasks','Title shown in app and emails'],
   ['app_subtitle','Shop Task Manager','Subtitle for dashboard header'],
   ['statuses','To Do,In Progress,Done,On Hold,Cancelled','Task statuses'],
   ['priorities','High,Medium,Low','Priority levels'],
   ['recurring_options','None,Daily,Weekly,Monthly,Yearly','Recurring options'],
   ['dashboard_refresh_seconds','120','TV dashboard refresh interval in seconds'],
   ['daily_digest_hour','7','Hour to send daily digest (0-23)'],
   ['completed_task_days','3','Days to show completed tasks on dashboard'],
   ['admin_emails','dwooswald@gmail.com','Comma-separated admin emails']
  ].forEach(r => cfg.appendRow(r));

  SpreadsheetApp.getUi().alert(
    'Setup complete!\n\n' +
    'All PINs set to 1234.\n\n' +
    'IMPORTANT: Deploy as Web App:\n' +
    '• Execute as: Me\n' +
    '• Who has access: Anyone\n\n' +
    'Then update API_URL in your GitHub Pages index.html'
  );
}

function createTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  const h = parseInt(getConfig()['daily_digest_hour'] || '7');
  ScriptApp.newTrigger('sendDailyDigest').timeBased().atHour(h).everyDays(1).create();
  Logger.log('Trigger set: daily digest at ' + h + ':00');
}
