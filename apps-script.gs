/**
 * Miguel's better HubSpot — Google Apps Script backend
 *
 * SETUP:
 * 1. Open your Google Sheet → Extensions → Apps Script
 * 2. Delete any boilerplate code, paste this entire file in
 * 3. Save (disk icon)
 * 4. Click "Deploy" → "New deployment"
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone (required so the static page can call it)
 * 5. Copy the Web app URL it gives you
 * 6. Paste that URL into the CRM's Settings panel and click Connect
 */

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const CONTACTS_SHEET = 'Contacts';
const TEMPLATES_SHEET = 'Email Templates';
const ACTIVITY_SHEET = 'Activity Log';

function doGet(e) {
  return handle(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: 'Invalid JSON' });
  }
  return handle(body);
}

function handle(req) {
  try {
    const action = req.action || 'list';
    switch (action) {
      case 'list':            return json({ ok: true, ...listAll() });
      case 'addContact':      return json({ ok: true, contact: addContact(req.contact) });
      case 'updateContact':   return json({ ok: true, contact: updateContact(req.id, req.fields) });
      case 'deleteContact':   return json({ ok: true, deleted: deleteContact(req.id) });
      case 'logActivity':     return json({ ok: true, logged: logActivity(req.entry) });
      case 'recordEmailSent': return json({ ok: true, contact: recordEmailSent(req.id, req.templateId) });
      case 'listActivity':    return json({ ok: true, activities: listActivity(req.limit || 30) });
      default:                return json({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json({ ok: false, error: err.message, stack: err.stack });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const s = ss.getSheetByName(name);
  if (!s) throw new Error('Missing sheet: ' + name);
  return s;
}

function readRows(name) {
  const s = sheet(name);
  const data = s.getDataRange().getValues();
  if (data.length < 2) return { headers: data[0] || [], rows: [] };
  const headers = data[0];
  const rows = data.slice(1).map((row, i) => {
    const obj = { _rowIndex: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });
  return { headers, rows };
}

function listAll() {
  const contactsData = readRows(CONTACTS_SHEET);
  const templatesData = readRows(TEMPLATES_SHEET);

  const contacts = contactsData.rows
    .filter(r => r['ID'] || r['Name'] || r['Email'])
    .map(r => ({
      id: String(r['ID'] || ''),
      name: String(r['Name'] || ''),
      email: String(r['Email'] || ''),
      phone: String(r['Phone'] || ''),
      comment: String(r['Comment'] || ''),
      createdAt: r['Created At'] ? new Date(r['Created At']).toISOString() : '',
      lastEmailed: r['Last Emailed'] ? new Date(r['Last Emailed']).toISOString() : '',
      emailsSent: Number(r['Emails Sent']) || 0,
      status: String(r['Status'] || 'New'),
      owner: String(r['Owner'] || '')
    }));

  const templates = templatesData.rows
    .filter(r => r['Template ID'])
    .map(r => ({
      id: String(r['Template ID']),
      name: String(r['Name'] || ''),
      desc: String(r['Description'] || ''),
      subject: String(r['Subject'] || ''),
      body: String(r['Body'] || '')
    }));

  return { contacts, templates };
}

function addContact(c) {
  if (!c || !c.name || !c.email) throw new Error('name and email required');
  const id = c.id || (Date.now().toString(36) + Math.random().toString(36).slice(2, 6));
  const createdAt = new Date();
  sheet(CONTACTS_SHEET).appendRow([
    id, c.name, c.email, c.phone || '', c.comment || '',
    createdAt, '', 0, c.status || 'New', c.owner || ''
  ]);
  logActivity({
    contactId: id, contactName: c.name, action: 'Contact Added',
    templateUsed: '', notes: ''
  });
  return {
    id, name: c.name, email: c.email, phone: c.phone || '',
    comment: c.comment || '', createdAt: createdAt.toISOString(),
    lastEmailed: '', emailsSent: 0, status: c.status || 'New', owner: c.owner || ''
  };
}

function findRowById(sheetName, id) {
  const s = sheet(sheetName);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function updateContact(id, fields) {
  const rowIndex = findRowById(CONTACTS_SHEET, id);
  if (rowIndex < 0) throw new Error('Contact not found: ' + id);
  const s = sheet(CONTACTS_SHEET);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const map = {
    name: 'Name', email: 'Email', phone: 'Phone', comment: 'Comment',
    status: 'Status', owner: 'Owner', lastEmailed: 'Last Emailed',
    emailsSent: 'Emails Sent'
  };
  Object.keys(fields).forEach(k => {
    const header = map[k];
    if (!header) return;
    const col = headers.indexOf(header) + 1;
    if (col > 0) s.getRange(rowIndex, col).setValue(fields[k]);
  });
  return { id, ...fields };
}

function deleteContact(id) {
  const rowIndex = findRowById(CONTACTS_SHEET, id);
  if (rowIndex < 0) throw new Error('Contact not found: ' + id);
  const s = sheet(CONTACTS_SHEET);
  const name = s.getRange(rowIndex, 2).getValue();
  s.deleteRow(rowIndex);
  logActivity({
    contactId: id, contactName: name, action: 'Contact Deleted',
    templateUsed: '', notes: ''
  });
  return id;
}

function logActivity(entry) {
  if (!entry) return false;
  sheet(ACTIVITY_SHEET).appendRow([
    new Date(),
    entry.contactId || '',
    entry.contactName || '',
    entry.action || '',
    entry.templateUsed || '',
    entry.notes || ''
  ]);
  return true;
}

function listActivity(limit) {
  const s = sheet(ACTIVITY_SHEET);
  const data = s.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1)
    .filter(r => r[0] || r[3])
    .map(r => ({
      timestamp:    r[0] ? new Date(r[0]).toISOString() : '',
      contactId:    String(r[1] || ''),
      contactName:  String(r[2] || ''),
      action:       String(r[3] || ''),
      templateUsed: String(r[4] || ''),
      notes:        String(r[5] || '')
    }))
    .reverse()
    .slice(0, limit);
}

function recordEmailSent(id, templateId) {
  const rowIndex = findRowById(CONTACTS_SHEET, id);
  if (rowIndex < 0) throw new Error('Contact not found: ' + id);
  const s = sheet(CONTACTS_SHEET);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const lastEmailedCol = headers.indexOf('Last Emailed') + 1;
  const sentCol = headers.indexOf('Emails Sent') + 1;
  const nameCol = headers.indexOf('Name') + 1;
  const statusCol = headers.indexOf('Status') + 1;
  const name = s.getRange(rowIndex, nameCol).getValue();
  const currentSent = Number(s.getRange(rowIndex, sentCol).getValue()) || 0;
  const currentStatus = String(s.getRange(rowIndex, statusCol).getValue() || '');
  s.getRange(rowIndex, lastEmailedCol).setValue(new Date());
  s.getRange(rowIndex, sentCol).setValue(currentSent + 1);
  if (currentStatus === 'New' || currentStatus === '') {
    s.getRange(rowIndex, statusCol).setValue('Contacted');
  }
  logActivity({
    contactId: id, contactName: name, action: 'Email Sent',
    templateUsed: templateId || '', notes: ''
  });
  return { id, emailsSent: currentSent + 1 };
}
