/**
 * Colleague Zero – application form lead capture
 *
 * Setup:
 *   1. Go to script.google.com > New project and paste this file in as Code.gs
 *      (or open the sheet below and use Extensions > Apps Script).
 *   2. Run setup() once from the editor and accept the permissions prompt.
 *   3. Deploy > New deployment > Web app
 *        Execute as: Me
 *        Who has access: Anyone
 *   4. Copy the /exec URL into LEADS_URL on the landing page.
 *
 * Each time you change this file, use Deploy > Manage deployments > Edit > New version,
 * otherwise the /exec URL keeps running the old code.
 */

const CONFIG = {
  SHEET_ID: '1AIATY0zY4Oy2QwwnsNUmoUSCNHrFFbPPCXVwGSrNO5k',
  // The tab to write to (the gid in the sheet URL). 0 is the first tab.
  SHEET_GID: 0,
  // Comma-separated addresses to notify on each new application. Leave empty to disable.
  NOTIFY_EMAIL: '',
  // Send the applicant a short confirmation email.
  SEND_CONFIRMATION: false,
  MAX_FIELD_LENGTH: 2000,
};

const COLUMNS = [
  { key: 'timestamp',         header: 'Submitted at' },
  { key: 'name',              header: 'Full name' },
  { key: 'email',             header: 'Email' },
  { key: 'phone',             header: 'Mobile number' },
  { key: 'org',               header: 'School, university or company' },
  { key: 'team',              header: 'Team name' },
  { key: 'size',              header: 'Team size' },
  { key: 'role',              header: 'Main role' },
  { key: 'exp',               header: 'Experience with AI agents' },
  { key: 'idea',              header: 'Digital colleague idea' },
  { key: 'inperson',          header: 'Can attend Sat 14 Nov' },
  { key: 'media_consent',     header: 'Media consent' },
  { key: 'marketing_consent', header: 'Marketing consent' },
  { key: 'duplicate',         header: 'Duplicate email' },
  { key: 'source',            header: 'Source page' },
];

const ROLES = ['Developer', 'Designer', 'Educator', 'Student', 'Product or business', 'Other'];
const EXPERIENCE = ['New to agents', 'Built a few prototypes', 'Shipped agents in production'];

/** Run once from the editor to create the sheet and header row. */
function setup() {
  const sheet = getSheet_();
  Logger.log('Ready: ' + sheet.getParent().getUrl());
}

function doGet() {
  return json_({ ok: true, service: 'colleague-zero-leads' });
}

function doPost(e) {
  let data;
  try {
    data = parseBody_(e);
  } catch (err) {
    return json_({ ok: false, error: 'Could not read the submission.' });
  }

  // Honeypot: bots fill hidden fields, people don't. Pretend success so bots move on.
  if (data.website) return json_({ ok: true });

  const errors = validate_(data);
  if (errors.length) return json_({ ok: false, error: errors.join(' ') });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'The server is busy. Please try again in a moment.' });
  }

  let record;
  try {
    const sheet = getSheet_();
    const email = clean_(data.email).toLowerCase();

    record = {
      timestamp:         new Date(),
      name:              clean_(data.name),
      email:             email,
      phone:             clean_(data.phone),
      org:               clean_(data.org),
      team:              clean_(data.team),
      size:              Number(data.size),
      role:              clean_(data.role),
      exp:               clean_(data.exp),
      idea:              clean_(data.idea),
      inperson:          yesNo_(data.inperson),
      media_consent:     yesNo_(data.media_consent),
      marketing_consent: yesNo_(data.marketing_consent),
      duplicate:         emailExists_(sheet, email) ? 'Yes' : '',
      source:            clean_(data.source),
    };

    sheet.appendRow(COLUMNS.map(c => safeCell_(record[c.key])));
  } finally {
    lock.releaseLock();
  }

  // Email failures (e.g. daily quota) shouldn't fail an application that is already saved.
  try { notify_(record); } catch (err) { console.error('Notify failed: ' + err); }
  try { confirm_(record); } catch (err) { console.error('Confirmation failed: ' + err); }

  return json_({ ok: true });
}

// ---------------------------------------------------------------------------

/** Accepts URL-encoded / multipart form posts (e.parameter) or a JSON body. */
function parseBody_(e) {
  const params = (e && e.parameter) || {};
  if (Object.keys(params).length) return params;

  const raw = e && e.postData && e.postData.contents;
  if (!raw) return {};
  return JSON.parse(raw);
}

function validate_(d) {
  const errors = [];
  const email = clean_(d.email);

  if (!clean_(d.name)) errors.push('Full name is required.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) errors.push('A valid email is required.');
  if (clean_(d.phone).replace(/\D/g, '').length < 9) errors.push('A valid mobile number is required.');
  if (!['1', '2', '3', '4', '5'].includes(String(d.size))) errors.push('Choose a team size from 1 to 5.');
  if (!ROLES.includes(clean_(d.role))) errors.push('Choose your main role.');
  if (!EXPERIENCE.includes(clean_(d.exp))) errors.push('Choose your experience with AI agents.');
  if (yesNo_(d.inperson) !== 'Yes') errors.push('You must be able to attend in person on Saturday 14 November.');
  if (yesNo_(d.media_consent) !== 'Yes') errors.push('Media consent is required to take part.');

  return errors;
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === CONFIG.SHEET_GID);
  if (!sheet) throw new Error('No tab with gid ' + CONFIG.SHEET_GID + ' in ' + ss.getUrl());

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(c => c.header));
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function emailExists_(sheet, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const col = COLUMNS.findIndex(c => c.key === 'email') + 1;
  return sheet.getRange(2, col, lastRow - 1, 1).getValues()
    .some(r => String(r[0]).replace(/^'/, '').toLowerCase() === email);
}

function notify_(r) {
  if (!CONFIG.NOTIFY_EMAIL) return;
  const body = COLUMNS
    .filter(c => c.key !== 'timestamp')
    .map(c => c.header + ': ' + (r[c.key] === '' ? '-' : r[c.key]))
    .join('\n');
  MailApp.sendEmail({
    to: CONFIG.NOTIFY_EMAIL,
    replyTo: r.email,
    subject: 'New Colleague Zero application: ' + r.name + (r.team ? ' (' + r.team + ')' : ''),
    body: body + '\n\nSheet: ' + 'https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID + '/edit#gid=' + CONFIG.SHEET_GID,
  });
}

function confirm_(r) {
  if (!CONFIG.SEND_CONFIRMATION) return;
  MailApp.sendEmail({
    to: r.email,
    subject: 'Your Colleague Zero application',
    name: 'Melsoft Academy',
    body:
      'Hi ' + r.name.split(' ')[0] + ',\n\n' +
      'Thanks for applying to Colleague Zero. We have your application and will email you ' +
      'with next steps before kickoff at 10:00 on Friday 13 November.\n\n' +
      'Melsoft Academy',
  });
}

function clean_(v) {
  return String(v == null ? '' : v).trim().slice(0, CONFIG.MAX_FIELD_LENGTH);
}

/** Checkboxes post "on" when ticked and are omitted when not. */
function yesNo_(v) {
  return ['on', 'true', 'yes', '1'].includes(String(v).toLowerCase()) ? 'Yes' : 'No';
}

/** Stop user input being run as a formula, and keep phone numbers like +27… as text. */
function safeCell_(v) {
  if (typeof v !== 'string') return v;
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
