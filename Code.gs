/**
 * Colleague Zero – application form lead capture
 *
 * Setup:
 *   1. Go to script.google.com > New project and paste this file in as Code.gs
 *      (or open the sheet below and use Extensions > Apps Script).
 *   2. Run setup() once from the editor and accept the permissions prompt. It creates the
 *      header rows, the "Email queue" tab and the 10-minute timer that sends follow-ups.
 *   3. Deploy > New deployment > Web app
 *        Execute as: Me
 *        Who has access: Anyone
 *   4. Copy the /exec URL into LEADS_URL on the landing page.
 *
 * Each time you change this file, use Deploy > Manage deployments > Edit > New version,
 * otherwise the /exec URL keeps running the old code.
 *
 * Emails to applicants (times are SAST). Whether the event is online or in person isn't decided
 * yet, so no email says either; the day-of details go out separately once they're set.
 *   1. Straight away: confirmation with their details, both days at a glance and a calendar
 *      invite (.ics) for both days.
 *   2. Two hours later: a welcome with the Discord invite.
 *   3. Lead-up: are you excited? (Thu 5 Nov 12:00), RSVP: are you coming? (Fri 6 Nov 09:00), sharpen
 *      your idea (Mon 9 Nov 09:00) and two days to go (Wed 11 Nov 09:00), then the morning of
 *      each day (Fri 13 Nov 08:00, Sat 14 Nov 08:00). Set JOIN_URL before 13 Nov if there's a link to join.
 *   Scheduled emails live in the "Email queue" tab; processFollowUps() runs every 10 minutes,
 *   queues reminders for anyone who doesn't have them yet and sends what's due.
 *   Preview with sendTestConfirmation(), sendTestWelcome() or sendTestReminders().
 *
 * RSVPs: each applicant gets a personal link to rsvp.html (signed, so nobody can answer for
 * someone else). Answers land in a separate spreadsheet, one row per person; setup() creates
 * it and logs its URL. The idea and two-days emails ask anyone who hasn't answered yet, and
 * people who say they can't come get no further reminders.
 *
 * Attendance: attend.html?day=1 and attend.html?day=2 (the QR codes in assets/qr point there).
 * Each sign-in lands in the "Attendance" tab of the applications sheet, one row per person per
 * day; signing in twice on the same day keeps the first time. A day only accepts sign-ins on
 * its own date (SAST) unless ATTENDANCE_ANY_DAY is true, which you can use to test beforehand.
 *
 * Feedback: feedback.html posts ratings (1–5) and comments to the "Feedback" tab of the
 * applications sheet, one row per answer. Only the overall rating is required, and answers
 * are anonymous (no name or email is asked for).
 */

const CONFIG = {
  SHEET_ID: '1AIATY0zY4Oy2QwwnsNUmoUSCNHrFFbPPCXVwGSrNO5k',
  // The tab to write to (the gid in the sheet URL). 0 is the first tab.
  SHEET_GID: 0,
  // Comma-separated addresses to notify on each new application. Leave empty to disable.
  NOTIFY_EMAIL: '',
  // Send the applicant a warm confirmation email. Run sendTestConfirmation() to preview it.
  SEND_CONFIRMATION: true,
  // Send the Discord welcome email WELCOME_DELAY_MINUTES after applying.
  SEND_WELCOME: true,
  WELCOME_DELAY_MINUTES: 120,
  DISCORD_URL: 'https://discord.gg/9wRRmEeMp',
  // Reminders before and on the event days (see REMINDERS below).
  SEND_REMINDERS: true,
  // Link people use to join on the day (kickoff and presentations). Leave empty until the
  // format is decided; the event-day emails then say the details are on their way.
  JOIN_URL: '',
  // Tab that holds scheduled follow-up emails. Created automatically.
  QUEUE_SHEET_NAME: 'Email queue',
  // The RSVP page on the site. Each email links here with the person's email and a signature.
  RSVP_URL: 'https://hackathon-26.melsoft.org/rsvp.html',
  // Spreadsheet the RSVPs go to. Leave empty and setup() creates "Colleague Zero · RSVPs" in
  // your Drive (its ID is kept in Script properties); or paste the ID of a sheet you made.
  RSVP_SHEET_ID: '',
  // Where applicants' replies go. Leave empty to use the account the script runs as.
  REPLY_TO: '',
  // Tab in the applications sheet that attendance sign-ins go to. Created automatically.
  ATTENDANCE_SHEET_NAME: 'Attendance',
  // true lets people sign in for either day on any date (for testing). Set false for the event.
  ATTENDANCE_ANY_DAY: false,
  // Tab in the applications sheet that feedback goes to. Created automatically.
  FEEDBACK_SHEET_NAME: 'Feedback',
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
  { key: 'inperson',          header: 'Can attend (old form only)' },
  { key: 'media_consent',     header: 'Media consent' },
  { key: 'marketing_consent', header: 'Marketing consent' },
  { key: 'duplicate',         header: 'Duplicate email' },
  { key: 'source',            header: 'Source page' },
];

// "Type" says which email a row is (welcome, rsvp, …). Rows from before it existed are welcomes.
const QUEUE_COLUMNS = ['Email', 'Name', 'Due at', 'Sent at', 'Status', 'Attempts', 'Type'];
const MAX_SEND_ATTEMPTS = 3;

const ROLES = ['Developer', 'Designer', 'Educator', 'Student', 'Product or business', 'Other'];
const EXPERIENCE = ['New to agents', 'Built a few prototypes', 'Shipped agents in production'];

// One row per person in the RSVP sheet; a changed answer overwrites their row.
const RSVP_COLUMNS = [
  { key: 'updated',  header: 'Updated at' },
  { key: 'email',    header: 'Email' },
  { key: 'name',     header: 'Name' },
  { key: 'coming',   header: 'Coming?' },
  { key: 'team',     header: 'Team name' },
  { key: 'people',   header: 'People coming (incl. them)' },
  { key: 'note',     header: 'Note' },
  { key: 'verified', header: 'Used their email link' },
  { key: 'first',    header: 'First answered at' },
];

// One row per person per day in the Attendance tab.
const ATTENDANCE_COLUMNS = [
  { key: 'time',    header: 'Signed in at' },
  { key: 'day',     header: 'Day' },
  { key: 'name',    header: 'Full name' },
  { key: 'email',   header: 'Email' },
  { key: 'phone',   header: 'Cell number' },
  { key: 'org',     header: 'Company or school' },
  { key: 'team',    header: 'Team name' },
  { key: 'applied', header: 'Applied?' },
  { key: 'rsvp',    header: 'RSVP' },
];

// One row per feedback answer. Ratings are 1–5; only overall is required.
const FEEDBACK_RATINGS = ['overall', 'challenge', 'food', 'vibe', 'organisation', 'venue'];
const FEEDBACK_COLUMNS = [
  { key: 'time',         header: 'Submitted at' },
  { key: 'overall',      header: 'Overall (1–5)' },
  { key: 'challenge',    header: 'Challenge (1–5)' },
  { key: 'food',         header: 'Food (1–5)' },
  { key: 'vibe',         header: 'Vibe (1–5)' },
  { key: 'organisation', header: 'How it was run (1–5)' },
  { key: 'venue',        header: 'Venue (1–5)' },
  { key: 'best',         header: 'Enjoyed most' },
  { key: 'better',       header: 'Do better next time' },
  { key: 'again',        header: 'Come to the next one?' },
];

// The days people can sign in for (date in SAST, as yyyy-MM-dd).
const ATTENDANCE_DAYS = {
  '1': { label: 'Day 1', date: '2026-11-13' },
  '2': { label: 'Day 2', date: '2026-11-14' },
};

// The answers on the RSVP page (the form posts the key).
const RSVP_ANSWERS = { yes: 'Yes', no: 'No' };

/** Run once from the editor: creates the header rows, the queue tab, the RSVP sheet and the follow-up timer. */
function setup() {
  const sheet = getSheet_();
  getQueueSheet_();
  getAttendanceSheet_();
  getFeedbackSheet_();
  const rsvp = getRsvpSheet_(true);
  rsvpToken_('setup'); // creates the signing secret before any link goes out
  setupFollowUps();
  Logger.log('Ready: ' + sheet.getParent().getUrl());
  Logger.log('RSVPs: ' + rsvp.getParent().getUrl());
}

/**
 * Run once from the editor: creates the Attendance and Feedback tabs (with their header rows)
 * in the applications sheet now, instead of waiting for the first response. Safe to run again.
 */
function createResponseTabs() {
  const attendance = getAttendanceSheet_();
  const feedback = getFeedbackSheet_();
  Logger.log('Attendance tab: ' + attendance.getParent().getUrl() + '#gid=' + attendance.getSheetId());
  Logger.log('Feedback tab: ' + feedback.getParent().getUrl() + '#gid=' + feedback.getSheetId());
}

/** Installs (or reinstalls) the 10-minute timer that sends due follow-up emails. */
function setupFollowUps() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processFollowUps')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processFollowUps').timeBased().everyMinutes(10).create();
  Logger.log('Follow-up timer installed: processFollowUps every 10 minutes.');
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.rsvp) return json_(rsvpLookup_(p));
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

  if (data.form === 'rsvp') return json_(handleRsvp_(data));
  if (data.form === 'attendance') return json_(handleAttendance_(data));
  if (data.form === 'feedback') return json_(handleFeedback_(data));

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
      inperson:          data.inperson ? yesNo_(data.inperson) : '',
      media_consent:     yesNo_(data.media_consent),
      marketing_consent: yesNo_(data.marketing_consent),
      duplicate:         emailExists_(sheet, email) ? 'Yes' : '',
      source:            clean_(data.source),
    };

    sheet.appendRow(COLUMNS.map(c => safeCell_(record[c.key])));

    // The application is saved; a queue problem shouldn't fail it.
    try { queueWelcome_(record); } catch (err) { console.error('Queueing welcome failed: ' + err); }
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

function getQueueSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.QUEUE_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.QUEUE_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(QUEUE_COLUMNS);
    sheet.getRange(1, 1, 1, QUEUE_COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getRange(1, QUEUE_COLUMNS.length).getValue() !== QUEUE_COLUMNS[QUEUE_COLUMNS.length - 1]) {
    // A queue tab made before the "Type" column existed: add the header.
    sheet.getRange(1, QUEUE_COLUMNS.length).setValue(QUEUE_COLUMNS[QUEUE_COLUMNS.length - 1]).setFontWeight('bold');
  }
  return sheet;
}

/** "email|type" for every row already in the queue. Rows without a type are welcomes. */
function queuedKeys_(queue) {
  const lastRow = queue.getLastRow();
  if (lastRow < 2) return new Set();
  return new Set(queue.getRange(2, 1, lastRow - 1, QUEUE_COLUMNS.length).getValues()
    .map(r => String(r[0]).replace(/^'/, '').trim().toLowerCase() + '|' + (r[6] || 'welcome')));
}

/** Adds one scheduled email to the queue unless that person already has one of that type. */
function queueEmail_(queue, keys, email, name, type, due) {
  const key = email + '|' + type;
  if (keys.has(key)) return;
  queue.appendRow([safeCell_(email), safeCell_(name), due, '', 'Queued', 0, type]);
  keys.add(key);
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

// --- Email 1: confirmation, sent straight away ------------------------------

function confirm_(r) {
  if (!CONFIG.SEND_CONFIRMATION) return;
  sendToApplicant_(r.email, confirmationEmail_(r));
}

/** Run from the editor to send yourself the confirmation email with sample details. */
function sendTestConfirmation() {
  const me = Session.getActiveUser().getEmail();
  confirm_({
    name: 'Thandi Mokoena',
    email: me,
    team: 'Rubric Rangers',
    size: 3,
    idea: 'A marking colleague that grades essays against our rubric and checks its own marks.',
  });
  Logger.log('Test confirmation sent to ' + me);
}

// --- Email 2: Discord welcome, sent WELCOME_DELAY_MINUTES later --------------

/** Queues the welcome (2 hours later) and every reminder still ahead. Call inside the script lock. */
function queueWelcome_(r) {
  const queue = getQueueSheet_();
  const keys = queuedKeys_(queue);
  if (CONFIG.SEND_WELCOME) {
    const due = new Date(r.timestamp.getTime() + CONFIG.WELCOME_DELAY_MINUTES * 60 * 1000);
    queueEmail_(queue, keys, r.email, r.name, 'welcome', due);
  }
  queueReminders_(queue, keys, r.email, r.name);
}

// --- Emails 3–8: reminders on fixed dates (all times SAST) ---------------------

const REMINDERS = [
  { type: 'countdown', at: new Date('2026-11-05T12:00:00+02:00') }, // 8 days before: are you excited?
  { type: 'rsvp',    at: new Date('2026-11-06T09:00:00+02:00') }, // 7 days before: are you coming?
  { type: 'idea',    at: new Date('2026-11-09T09:00:00+02:00') }, // 4 days before; applications close
  { type: 'remind2', at: new Date('2026-11-11T09:00:00+02:00') }, // 2 days before
  { type: 'day1',    at: new Date('2026-11-13T08:00:00+02:00') }, // morning of day 1 (kickoff 09:30)
  { type: 'day2',    at: new Date('2026-11-14T08:00:00+02:00') }, // morning of day 2 (submissions 11:00)
];
// A reminder that couldn't go out on time (e.g. the timer was off) is skipped after this long.
const REMINDER_GRACE_HOURS = 12;
// Emails dropped from the schedule; rows still queued for them are cancelled, not sent.
// (remind7, "one week to go", gave its Fri 6 Nov slot to the RSVP email.)
const RETIRED_EMAILS = ['remind7'];

/** Queues every reminder that is still ahead, or due within the grace period, for one person. */
function queueReminders_(queue, keys, email, name) {
  if (!CONFIG.SEND_REMINDERS) return;
  const now = Date.now();
  REMINDERS.filter(r => r.at.getTime() + REMINDER_GRACE_HOURS * 3600 * 1000 > now)
    .forEach(r => queueEmail_(queue, keys, email, name, r.type, r.at));
}

/** Makes sure everyone who has applied has their upcoming reminders queued (once each). */
function syncReminders_(queue) {
  if (!CONFIG.SEND_REMINDERS) return;
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const keys = queuedKeys_(queue);
  const col = key => COLUMNS.findIndex(c => c.key === key);
  sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues().forEach(row => {
    const email = String(row[col('email')]).replace(/^'/, '').trim().toLowerCase();
    if (email) queueReminders_(queue, keys, email, String(row[col('name')]).replace(/^'/, '').trim());
  });
}

function emailFor_(type, person) {
  switch (type) {
    case 'welcome': return welcomeEmail_(person);
    case 'countdown': return reminderEmail_(person, 'countdown');
    case 'rsvp':    return reminderEmail_(person, 'rsvp');
    case 'idea':    return reminderEmail_(person, 'idea');
    case 'remind2': return reminderEmail_(person, 'remind2');
    case 'day1':    return reminderEmail_(person, 'day1');
    case 'day2':    return reminderEmail_(person, 'day2');
    default: throw new Error('Unknown email type: ' + type);
  }
}

/**
 * Timer handler (every 10 minutes): queues reminders for anyone who doesn't have them yet,
 * then sends every queued email that is due.
 */
function processFollowUps() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return; // another run is busy; the next one will catch up
  try {
    const queue = getQueueSheet_();
    syncReminders_(queue);
    const lastRow = queue.getLastRow();
    if (lastRow < 2) return;

    const rows = queue.getRange(2, 1, lastRow - 1, QUEUE_COLUMNS.length).getValues();
    const now = new Date();
    rows.forEach((row, i) => {
      const [email, name, due, sentAt, , attempts, typeCell] = row;
      const type = typeCell || 'welcome';
      if (sentAt || !email || !(due instanceof Date) || due > now) return;
      if (Number(attempts) >= MAX_SEND_ATTEMPTS) return;
      const rowNum = i + 2;
      const to = String(email).replace(/^'/, '');
      const skip = RETIRED_EMAILS.includes(type) ? 'Cancelled'
        : type !== 'welcome' && now - due > REMINDER_GRACE_HOURS * 3600 * 1000 ? 'Skipped (too late)'
        : !['welcome', 'rsvp'].includes(type) && (rsvpAnswers_().get(to.toLowerCase()) || {}).coming === 'No' ? 'Skipped (not coming)'
        : '';
      if (skip) {
        if (String(row[4]) !== skip) queue.getRange(rowNum, 5).setValue(skip);
        return;
      }
      if (MailApp.getRemainingDailyQuota() < 1) return; // try again after the quota resets

      try {
        sendToApplicant_(to, emailFor_(type, { name: String(name).replace(/^'/, ''), email: to }));
        queue.getRange(rowNum, 4, 1, 3).setValues([[new Date(), 'Sent', Number(attempts) + 1]]);
      } catch (err) {
        queue.getRange(rowNum, 5, 1, 2).setValues([['Error: ' + String((err && err.message) || err).slice(0, 200), Number(attempts) + 1]]);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

// --- One-off catch-up for people who applied before these emails existed ------

/** Run first: logs who runBackfill() would email, without sending anything. */
function previewBackfill() {
  const people = backfillCandidates_();
  Logger.log(people.length + ' applicant(s) would get the confirmation now and the welcome 2 hours later:');
  people.forEach(p => Logger.log('  ' + p.name + ' <' + p.email + '>'));
}

/**
 * Sends each earlier applicant the confirmation now and queues their welcome for 2 hours
 * later, exactly as if they had just applied. Anyone already in the "Email queue" tab is
 * skipped, so it is safe to run again (e.g. after hitting the daily email limit).
 */
function runBackfill() {
  const people = backfillCandidates_();
  let done = 0;
  for (const p of people) {
    if (MailApp.getRemainingDailyQuota() < 2) {
      Logger.log('Daily email limit reached after ' + done + '. Run runBackfill() again tomorrow to continue.');
      break;
    }
    confirm_(p);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { queueWelcome_(Object.assign({}, p, { timestamp: new Date() })); } finally { lock.releaseLock(); }
    done++;
  }
  Logger.log('Backfill: emailed ' + done + ' of ' + people.length + '.');
}

/** Earlier applicants (one per email address) who haven't had their welcome queued yet. */
function backfillCandidates_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const col = key => COLUMNS.findIndex(c => c.key === key);
  const unquote = v => String(v).replace(/^'/, '').trim();

  // Reminder rows don't count: the timer queues those for everyone.
  const queued = queuedKeys_(getQueueSheet_());

  const seen = new Set();
  const people = [];
  sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues().forEach(row => {
    const email = unquote(row[col('email')]).toLowerCase();
    if (!email || queued.has(email + '|welcome') || seen.has(email)) return;
    seen.add(email);
    people.push({
      name: unquote(row[col('name')]),
      email: email,
      team: unquote(row[col('team')]),
      size: Number(row[col('size')]) || 1,
      idea: unquote(row[col('idea')]),
    });
  });
  return people;
}

/** Run from the editor to send yourself the Discord welcome email. */
function sendTestWelcome() {
  const me = Session.getActiveUser().getEmail();
  sendToApplicant_(me, welcomeEmail_({ name: 'Thandi Mokoena', email: me }));
  Logger.log('Test welcome sent to ' + me);
}

// --- RSVPs ----------------------------------------------------------------------

/** Signature for a person's RSVP link, so the page can trust which applicant it is. */
function rsvpToken_(email) {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('RSVP_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('RSVP_SECRET', secret);
  }
  const sig = Utilities.computeHmacSha256Signature(String(email).toLowerCase(), secret);
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '').slice(0, 22);
}

function rsvpLink_(email) {
  return CONFIG.RSVP_URL + '?e=' + encodeURIComponent(email) + '&t=' + rsvpToken_(email);
}

/**
 * The RSVP sheet (first tab of its own spreadsheet). With create, makes the spreadsheet if
 * there isn't one yet; without, returns null instead. Create only under the script lock.
 */
function getRsvpSheet_(create) {
  const props = PropertiesService.getScriptProperties();
  const id = CONFIG.RSVP_SHEET_ID || props.getProperty('RSVP_SHEET_ID');
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else if (create) {
    ss = SpreadsheetApp.create('Colleague Zero · RSVPs');
    props.setProperty('RSVP_SHEET_ID', ss.getId());
    ss.getSheets()[0].setName('RSVPs');
  } else {
    return null;
  }
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(RSVP_COLUMNS.map(c => c.header));
    sheet.getRange(1, 1, 1, RSVP_COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** The latest application for an email: { name, team, size }, or null. */
function findApplicant_(email) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const col = key => COLUMNS.findIndex(c => c.key === key);
  const unquote = v => String(v).replace(/^'/, '').trim();
  const rows = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (unquote(rows[i][col('email')]).toLowerCase() !== email) continue;
    return { name: unquote(rows[i][col('name')]), team: unquote(rows[i][col('team')]), size: Number(rows[i][col('size')]) || 1 };
  }
  return null;
}

let rsvpCache_ = null;

/** Email → that person's RSVP row (keyed by RSVP_COLUMNS), read once per run. */
function rsvpAnswers_() {
  if (rsvpCache_) return rsvpCache_;
  rsvpCache_ = new Map();
  const sheet = getRsvpSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return rsvpCache_;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, RSVP_COLUMNS.length).getValues().forEach((row, i) => {
    const rec = { row: i + 2 };
    RSVP_COLUMNS.forEach((c, j) => { rec[c.key] = typeof row[j] === 'string' ? row[j].replace(/^'/, '') : row[j]; });
    rsvpCache_.set(String(rec.email).toLowerCase(), rec);
  });
  return rsvpCache_;
}

/** For the RSVP page (GET ?rsvp=1&e=…&t=…): who this is and any answer they already gave. */
function rsvpLookup_(p) {
  const email = clean_(p.e).toLowerCase();
  if (!email || clean_(p.t) !== rsvpToken_(email)) return { ok: false };
  const person = findApplicant_(email);
  if (!person) return { ok: false };
  const prev = rsvpAnswers_().get(email);
  const answer = prev ? Object.keys(RSVP_ANSWERS).find(k => RSVP_ANSWERS[k] === prev.coming) || '' : '';
  return {
    ok: true,
    first: firstName_(person.name),
    team: prev ? String(prev.team) : person.team,
    people: prev && prev.people ? Number(prev.people) : person.size,
    answer: answer,
  };
}

/** Saves an RSVP from the page, or overwrites that person's earlier answer. */
function handleRsvp_(d) {
  const email = clean_(d.email).toLowerCase();
  const answer = RSVP_ANSWERS[clean_(d.answer)];
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'Enter the email address you applied with.' };
  if (!answer) return { ok: false, error: 'Let us know whether you’re coming.' };
  const people = answer === 'Yes' ? Number(d.people) : '';
  if (answer === 'Yes' && !(people >= 1 && people <= 5)) return { ok: false, error: 'Choose how many people are coming, from 1 to 5.' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return { ok: false, error: 'The server is busy. Please try again in a moment.' };
  }
  try {
    const person = findApplicant_(email);
    if (!person) {
      return { ok: false, error: 'We couldn’t find an application for ' + email + '. Use the email address you applied with, or call us on ' + EVENT.phone + '.' };
    }
    const sheet = getRsvpSheet_(true);
    rsvpCache_ = null; // read the sheet afresh now that we hold the lock
    const prev = rsvpAnswers_().get(email);
    const now = new Date();
    const rec = {
      updated:  now,
      email:    email,
      name:     person.name,
      coming:   answer,
      team:     answer === 'Yes' ? clean_(d.team) : '',
      people:   people,
      note:     clean_(d.note),
      verified: clean_(d.t) === rsvpToken_(email) ? 'Yes' : 'No',
      first:    prev && prev.first ? prev.first : now,
    };
    const values = RSVP_COLUMNS.map(c => safeCell_(rec[c.key]));
    if (prev) sheet.getRange(prev.row, 1, 1, values.length).setValues([values]);
    else sheet.appendRow(values);
    return { ok: true, first: firstName_(person.name), coming: answer };
  } finally {
    lock.releaseLock();
  }
}

/** Run from the editor: logs your own RSVP link, to try the page. */
function testRsvpLink() {
  Logger.log(rsvpLink_(Session.getActiveUser().getEmail()));
}

// --- Attendance ---------------------------------------------------------------

function getAttendanceSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.ATTENDANCE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ATTENDANCE_COLUMNS.map(c => c.header));
    sheet.getRange(1, 1, 1, ATTENDANCE_COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Signs someone in for a day from attend.html. Anyone can sign in (team-mates may not have
 * applied themselves); the Applied? and RSVP columns show who we already knew about.
 */
function handleAttendance_(d) {
  const day = ATTENDANCE_DAYS[clean_(d.day)];
  const name = clean_(d.name);
  const email = clean_(d.email).toLowerCase();
  if (!day) return { ok: false, error: 'This sign-in link is missing its day. Scan the QR code again.' };
  if (!name) return { ok: false, error: 'Enter your full name.' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'Enter a valid email address.' };
  if (clean_(d.phone).replace(/\D/g, '').length < 9) return { ok: false, error: 'Enter a valid cell number.' };
  if (!clean_(d.org)) return { ok: false, error: 'Enter your company or school.' };

  const today = Utilities.formatDate(new Date(), 'Africa/Johannesburg', 'yyyy-MM-dd');
  if (!CONFIG.ATTENDANCE_ANY_DAY && today !== day.date) {
    return { ok: false, error: today < day.date
      ? 'Sign-in for ' + day.label + ' opens on the day. See you then!'
      : 'Sign-in for ' + day.label + ' has closed. If you were there, tell a Melsoft team member.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return { ok: false, error: 'The server is busy. Please try again in a moment.' };
  }
  try {
    const sheet = getAttendanceSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const rows = sheet.getRange(2, 1, lastRow - 1, ATTENDANCE_COLUMNS.length).getValues();
      const dayCol = ATTENDANCE_COLUMNS.findIndex(c => c.key === 'day');
      const emailCol = ATTENDANCE_COLUMNS.findIndex(c => c.key === 'email');
      const prev = rows.find(r => r[dayCol] === day.label && String(r[emailCol]).replace(/^'/, '').toLowerCase() === email);
      if (prev) {
        const at = Utilities.formatDate(new Date(prev[0]), 'Africa/Johannesburg', 'HH:mm');
        return { ok: true, already: true, first: firstName_(name), at: at, day: day.label };
      }
    }
    const person = findApplicant_(email);
    const rsvp = rsvpAnswers_().get(email);
    const now = new Date();
    const rec = {
      time:    now,
      day:     day.label,
      name:    name,
      email:   email,
      phone:   clean_(d.phone),
      org:     clean_(d.org),
      team:    clean_(d.team),
      applied: person ? 'Yes' : 'No',
      rsvp:    rsvp ? rsvp.coming : '',
    };
    sheet.appendRow(ATTENDANCE_COLUMNS.map(c => safeCell_(rec[c.key])));
    return { ok: true, first: firstName_(name), at: Utilities.formatDate(now, 'Africa/Johannesburg', 'HH:mm'), day: day.label };
  } finally {
    lock.releaseLock();
  }
}

// --- Feedback -----------------------------------------------------------------

function getFeedbackSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.FEEDBACK_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.FEEDBACK_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(FEEDBACK_COLUMNS.map(c => c.header));
    sheet.getRange(1, 1, 1, FEEDBACK_COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Saves one feedback answer from feedback.html. */
function handleFeedback_(d) {
  const rec = { time: new Date() };
  for (const key of FEEDBACK_RATINGS) {
    const v = clean_(d[key]);
    if (v && !['1', '2', '3', '4', '5'].includes(v)) return { ok: false, error: 'Ratings go from 1 to 5.' };
    rec[key] = v ? Number(v) : '';
  }
  if (!rec.overall) return { ok: false, error: 'Give Colleague Zero an overall rating from 1 to 5.' };
  rec.best = clean_(d.best);
  rec.better = clean_(d.better);
  rec.again = ['Yes', 'Maybe', 'No'].includes(clean_(d.again)) ? clean_(d.again) : '';

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return { ok: false, error: 'The server is busy. Please try again in a moment.' };
  }
  try {
    getFeedbackSheet_().appendRow(FEEDBACK_COLUMNS.map(c => safeCell_(rec[c.key])));
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// --- Email content ------------------------------------------------------------

// The event, as stated on the landing page. Times are SAST; online or in person is still open,
// so nothing here names a place.
const EVENT = {
  kickoff: new Date('2026-11-13T09:30:00+02:00'),
  site: 'https://hackathon-26.melsoft.org/',
  // Calendar entries for both days (UTC; SAST is UTC+2).
  sessions: [
    {
      uid: 'colleague-zero-2026-day1@melsoft.org',
      title: 'Colleague Zero · Day 1',
      start: '20261113T073000Z', end: '20261113T220000Z',
      details: 'Opening at 09:30 SAST, idea pitches at 12:30, mentor sessions from 14:30, progress check-in survey at 20:30, Loom progress reports by midnight.',
    },
    {
      uid: 'colleague-zero-2026-day2@melsoft.org',
      title: 'Colleague Zero · Day 2',
      start: '20261114T070000Z', end: '20261114T120000Z',
      details: 'Check-in at 09:00 SAST, projects submitted at 11:00, final presentations at 12:00, winner announced at 14:00.',
    },
  ],
  days: [
    {
      day: 'Friday 13 November',
      items: [['09:30', 'Opening and welcome'], ['12:30', 'Idea pitches'], ['14:30', 'Development and mentor sessions'], ['20:30', 'Progress check-in survey'], ['00:00', 'Loom progress reports']],
    },
    {
      day: 'Saturday 14 November',
      items: [['09:00', 'Arrival and check-in'], ['11:00', 'Project submission'], ['12:00', 'Final presentations'], ['14:00', 'Winner announced: R10 000']],
    },
  ],
  phone: '010 158 4346',
};

function calendarUrl_(s) {
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE' +
    '&text=' + encodeURIComponent(s.title) +
    '&dates=' + s.start + '/' + s.end +
    '&details=' + encodeURIComponent(s.details + ' ' + EVENT.site);
}

/** A calendar invite (.ics) with both days, attached to the confirmation email. */
function calendarIcs_() {
  const esc = s => String(s).replace(/\\/g, '\\\\').replace(/([,;])/g, '\\$1');
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+/, '');
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Melsoft Academy//Colleague Zero//EN',
    'CALSCALE:GREGORIAN', 'METHOD:PUBLISH'];
  EVENT.sessions.forEach(s => lines.push(
    'BEGIN:VEVENT', 'UID:' + s.uid, 'DTSTAMP:' + stamp, 'DTSTART:' + s.start, 'DTEND:' + s.end,
    'SUMMARY:' + esc(s.title), 'DESCRIPTION:' + esc(s.details + ' ' + EVENT.site),
    'URL:' + EVENT.site, 'END:VEVENT'));
  lines.push('END:VCALENDAR');
  // Calendar files wrap lines longer than 75 bytes: continuation lines start with a space.
  const fold = l => l.length <= 60 ? l : l.match(/.{1,60}/g).join('\r\n ');
  return Utilities.newBlob(lines.map(fold).join('\r\n'), 'text/calendar', 'colleague-zero.ics');
}

function sendToApplicant_(to, email) {
  const options = { to: to, subject: email.subject, name: 'Melsoft Academy', body: email.text, htmlBody: email.html };
  if (email.attachments) options.attachments = email.attachments;
  if (CONFIG.REPLY_TO) options.replyTo = CONFIG.REPLY_TO;
  MailApp.sendEmail(options);
}

function firstName_(name) {
  return String(name).trim().split(/\s+/)[0] || 'there';
}

/** Agenda for the given days, as plain text and as email HTML. */
function agendaText_(days) {
  return days.map(d => d.day + '\n' + d.items.map(i => '  ' + i[0] + '  ' + i[1]).join('\n')).join('\n\n');
}

function agendaHtml_(title, days) {
  const esc = escapeHtml_;
  const rows = days.map(d =>
    '<tr><td style="padding:18px 0 6px;font:700 15px/1.4 Inter,Arial,sans-serif;color:#0a1733">' + esc(d.day) + '</td></tr>' +
    d.items.map(i =>
      '<tr><td style="padding:4px 0;font:15px/1.5 Inter,Arial,sans-serif;color:#34405e">' +
      '<span style="display:inline-block;width:56px;font-weight:700;color:#5e0743">' + esc(i[0]) + '</span>' + esc(i[1]) +
      '</td></tr>').join('')
  ).join('');
  return '<tr><td style="padding:16px 36px 8px">' +
    '<div style="border-top:1px solid #e3e8f2;padding-top:20px;font:700 12px/1 Inter,Arial,sans-serif;letter-spacing:.14em;text-transform:uppercase;color:#5e0743">' + esc(title) + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + rows + '</table>' +
    '</td></tr>';
}

/** "Add Day 1 / Add Day 2" calendar buttons, side by side. */
function calendarButtons_() {
  return '<tr><td style="padding:24px 36px 8px">' +
    EVENT.sessions.map((s, i) =>
      '<a href="' + escapeHtml_(calendarUrl_(s)) + '" style="display:inline-block;margin:0 8px 8px 0;padding:14px 22px;border-radius:999px;' +
      navyBg_('#0a1733') + 'color:#ffffff;font:700 15px/1 Inter,Arial,sans-serif;text-decoration:none">' +
      keepWhite_('Add Day ' + (i + 1) + ' to calendar', 'inline-block') + '</a>').join('') +
    '</td></tr>';
}

function confirmationEmail_(r) {
  const first = firstName_(r.name);
  const entry = Number(r.size) > 1
    ? (r.team ? 'your team, ' + r.team + ', with ' : 'a team of ') + r.size + ' people'
    : 'you as an individual';
  const subject = 'Thanks for applying to Colleague Zero, ' + first;

  // Plain-text version for mail apps that don't show HTML.
  const text = [
    'Hi ' + first + ',',
    '',
    'Thank you for applying to Colleague Zero, our 24-hour hackathon on multi-agent workflows. ' +
      "We're so glad you want to build a digital colleague with us.",
    '',
    "We've entered " + entry + '.' +
      (r.idea ? ' Your idea: "' + r.idea + '"' + (/[.!?]$/.test(r.idea) ? '' : '.') + ' You can change it on the day.' : ''),
    '',
    "What happens next: we'll email you with next steps before kickoff at 11:00 on Friday 13 November.",
    '',
    'THE WEEKEND AT A GLANCE (times SAST)',
    '',
    agendaText_(EVENT.days),
    '',
    'A calendar invite for both days is attached. Or add them to Google Calendar:',
    '  Day 1: ' + calendarUrl_(EVENT.sessions[0]),
    '  Day 2: ' + calendarUrl_(EVENT.sessions[1]),
    '',
    'Questions? Call the Melsoft Academy team on ' + EVENT.phone + '.',
    '',
    'See you soon,',
    'The Melsoft Academy team',
  ].join('\n');

  const esc = escapeHtml_;
  const html = emailLayout_({
    preview: 'Your application is in. Here’s the weekend at a glance.',
    eyebrow: 'Application received',
    heading: 'Thanks for applying, ' + esc(first) + '.',
    sections: [
      section_(
        '<p style="margin:0 0 16px">Thank you for applying to <strong style="color:#0a1733">Colleague Zero</strong>, our 24-hour hackathon on multi-agent workflows. ' +
        'We’re so glad you want to build a digital colleague with us.</p>' +
        '<p style="margin:0 0 16px">We’ve entered ' + esc(entry) + '.</p>' +
        (r.idea
          ? '<div style="margin:0 0 16px;padding:16px 18px;border-left:3px solid #5e0743;background:#f7e9f2;border-radius:0 12px 12px 0">' +
            '<div style="font:700 11px/1 Inter,Arial,sans-serif;letter-spacing:.14em;text-transform:uppercase;color:#5e0743">Your idea</div>' +
            '<div style="margin-top:8px;color:#0a1733">' + esc(r.idea) + '</div>' +
            '<div style="margin-top:6px;font-size:13px;color:#646c85">You can change it on the day.</div></div>'
          : '') +
        '<p style="margin:0">We’ll email you with next steps before kickoff at <strong style="color:#0a1733">11:00 on Friday 13 November</strong>.</p>',
        '32px 36px 8px'),
      agendaHtml_('The weekend at a glance', EVENT.days),
      calendarButtons_(),
      section_('<p style="margin:0;font-size:14px;color:#646c85">A calendar invite for both days is attached to this email.</p>', '4px 36px 0'),
      signOff_(),
    ],
  });

  return { subject: subject, text: text, html: html, attachments: [calendarIcs_()] };
}

// --- Reminders ----------------------------------------------------------------

const JOIN_PENDING = 'Keep an eye on your inbox and Discord: we’ll send you the details to join before it starts.';

/** How to join on the day: a button once CONFIG.JOIN_URL is set, otherwise a heads-up. */
function joinHtml_(label) {
  return CONFIG.JOIN_URL
    ? button_(CONFIG.JOIN_URL, label, '#0a1733')
    : section_('<p style="margin:0">' + JOIN_PENDING + '</p>', '16px 36px 0');
}

function joinText_(label) {
  return CONFIG.JOIN_URL ? label + ': ' + CONFIG.JOIN_URL : JOIN_PENDING.replace(/’/g, "'");
}

/** A numbered list of [title, text] steps, as plain text and as email HTML. */
function stepsText_(steps) {
  return steps.map((s, i) => (i + 1) + '. ' + s[0] + ' ' + s[1]).join('\n');
}

function stepsHtml_(steps) {
  const esc = escapeHtml_;
  return '<tr><td style="padding:16px 36px 0"><table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
    steps.map((s, i) =>
      '<tr><td valign="top" style="width:40px;padding:14px 0;' + (i ? 'border-top:1px solid #e3e8f2;' : '') +
      'font:700 15px/1.55 Inter,Arial,sans-serif;color:#5e0743">0' + (i + 1) + '</td>' +
      '<td style="padding:14px 0;' + (i ? 'border-top:1px solid #e3e8f2;' : '') + 'font:15px/1.55 Inter,Arial,sans-serif;color:#34405e">' +
      '<strong style="color:#0a1733">' + esc(s[0]) + '</strong><br>' + esc(s[1]) + '</td></tr>').join('') +
    '</table></td></tr>';
}

/** A text link to a page on the site, e.g. "See how projects are judged →". */
function pageLinkHtml_(label, url) {
  return section_('<p style="margin:0"><a href="' + escapeHtml_(url) + '" style="color:#5e0743;font-weight:700;text-decoration:none">' +
    escapeHtml_(label) + ' &rarr;</a></p>', '16px 36px 0');
}

/**
 * The RSVP ask in an email, as { text, html: [sections] }. "ask" is the RSVP email's button;
 * "nudge" asks again if they haven't answered, or shows their answer with a link to change it.
 */
function rsvpBlock_(email, mode) {
  const link = rsvpLink_(email);
  const button = button_(link, 'Confirm your spot', '#5e0743');
  const prev = mode === 'nudge' ? rsvpAnswers_().get(String(email).toLowerCase()) : null;
  if (mode === 'ask') return { text: 'Confirm your spot: ' + link, html: [button] };
  if (!prev) {
    return {
      text: "We still need your RSVP. Let us know if you're coming: " + link,
      html: [section_('<p style="margin:0"><strong style="color:#0a1733">We still need your RSVP.</strong> ' +
        'Let us know if you’re coming so we can plan the weekend.</p>', '16px 36px 0'), button],
    };
  }
  const said = prev.coming === 'Yes' ? 'You’re confirmed. See you on Friday.' : 'You told us you can’t make it.';
  return {
    text: said.replace(/’/g, "'") + ' Plans changed? Update your RSVP: ' + link,
    html: [section_('<div style="padding:14px 18px;border-left:3px solid #5e0743;background:#f7e9f2;border-radius:0 12px 12px 0">' +
      '<strong style="color:#0a1733">' + escapeHtml_(said) + '</strong> Plans changed? ' +
      '<a href="' + escapeHtml_(link) + '" style="color:#5e0743;font-weight:700;text-decoration:none">Update your RSVP &rarr;</a></div>', '16px 36px 0')],
  };
}

function reminderEmail_(r, kind) {
  const first = firstName_(r.name);
  const esc = escapeHtml_;
  const discordHtml = button_(CONFIG.DISCORD_URL, 'Join us on Discord', '#5865F2');
  const discordText = 'Join us on Discord: ' + CONFIG.DISCORD_URL;
  const signText = ['', 'Questions? Call the Melsoft Academy team on ' + EVENT.phone + '.', '', 'See you soon,', 'The Melsoft Academy team'];
  const detailsSoon = 'We’ll send you everything you need for the day, including how to join, closer to kickoff.';
  let e;

  if (kind === 'countdown') {
    e = {
      subject: 'Are you excited, ' + first + '? Colleague Zero is eight days away',
      preview: 'Eight days to go. Three ways to start getting ready today.',
      eyebrow: 'Eight days to go',
      heading: 'Are you excited, ' + esc(first) + '? We are.',
      intro: 'In eight days, on <strong style="color:#0a1733">Friday 13 November at 11:00 SAST</strong>, Colleague Zero kicks off: ' +
        '24 hours to build a digital colleague that does real work for education. The builders who prepare now get the most out of it.',
      introText: 'In eight days, on Friday 13 November at 11:00 SAST, Colleague Zero kicks off: 24 hours to build a digital colleague that does real work for education. The builders who prepare now get the most out of it.',
      steps: [
        ['Get hands-on with your tools.', 'Build something small with an AI tool this week, even a simple agent. Every hour you play now is time saved on the day.'],
        ['Look for the busywork.', 'Notice where teachers, students and school staff lose time. Your best idea might be hiding in plain sight.'],
        ['Clear your weekend.', 'Block out Friday 13 and Saturday 14 November so you can give it your full focus.'],
      ],
      more: 'Come with your curiosity and your best energy. We can’t wait to see what you build.',
    };
  } else if (kind === 'rsvp') {
    e = {
      subject: 'Are you coming, ' + first + '? Confirm your spot for Colleague Zero',
      preview: 'One week to go. Tell us if you’re coming: it takes less than a minute.',
      eyebrow: 'One week to go · RSVP',
      heading: 'Will we see you there, ' + esc(first) + '?',
      intro: 'Colleague Zero kicks off one week from today, on <strong style="color:#0a1733">Friday 13 November at 11:00 SAST</strong>. ' +
        'To plan the weekend well we need to know exactly who’s coming, so please confirm your spot. It takes less than a minute.',
      introText: 'Colleague Zero kicks off one week from today, on Friday 13 November at 11:00 SAST. To plan the weekend well we need to know exactly who\'s coming, so please confirm your spot. It takes less than a minute.',
      rsvp: 'ask',
      more: 'Can’t make it any more? Please use the same button to tell us, so we can plan. We’d love your answer by Wednesday 11 November.',
      days: EVENT.days, calendar: true,
    };
  } else if (kind === 'idea') {
    e = {
      subject: 'What will your digital colleague do, ' + first + '?',
      preview: 'Four days to go. Three questions that make an idea stand out.',
      eyebrow: 'Four days to go',
      heading: 'Sharpen your idea, ' + esc(first) + '.',
      intro: 'Kickoff is this Friday. The strongest projects start with a clear answer to three questions.',
      introText: 'Kickoff is this Friday. The strongest projects start with a clear answer to three questions.',
      steps: [
        ['Who is it for?', 'A teacher, a student, a lecturer or a school admin. Pick one person and picture their week.'],
        ['What does it take off their plate?', 'A whole task from start to finish, not a single answer to a single question.'],
        ['How does it know it got it right?', 'The judges look for agents that check their own work and fix their mistakes.'],
      ],
      rsvp: 'nudge',
      more: 'Applications close today. If a friend or classmate would love this, send them to ' + EVENT.site + ' before tonight.',
      link: ['See how projects are judged', EVENT.site + 'judges.html'],
    };
  } else if (kind === 'remind2') {
    e = {
      subject: 'Two days to go, ' + first + '! Your checklist',
      preview: 'Kickoff is Friday at 11:00 SAST. Here’s what to have ready.',
      eyebrow: 'Two days to go',
      heading: 'Two days to go, ' + esc(first) + '.',
      intro: 'Kickoff is this <strong style="color:#0a1733">Friday, 13 November, at 11:00 SAST</strong>. ' +
        'A little prep now means more time building on the day.',
      introText: 'Kickoff is this Friday, 13 November, at 11:00 SAST. A little prep now means more time building on the day.',
      steps: [
        ['Your laptop and charger.', 'You’ll be building for 24 hours.'],
        ['Your AI tools.', 'Accounts and API keys for the tools you plan to use, set up and tested before Friday.'],
        ['A GitHub account.', 'You’ll submit your project as a repository with a README.'],
        ['A plan for your team.', 'Everyone knows the problem you’re solving and who’s doing what.'],
      ],
      rsvp: 'nudge',
      more: detailsSoon,
      link: ['See everything you need', EVENT.site + 'needed.html'],
      days: EVENT.days, calendar: true,
    };
  } else if (kind === 'day1') {
    e = {
      subject: 'Today’s the day: Colleague Zero kicks off at 09:30',
      preview: 'Kickoff is at 09:30 SAST. Here’s today’s plan.',
      eyebrow: 'Day 1 · Friday 13 November',
      heading: 'Today’s the day, ' + esc(first) + '!',
      intro: 'Colleague Zero kicks off at <strong style="color:#0a1733">09:30 SAST</strong>.',
      introText: 'Colleague Zero kicks off at 09:30 SAST.',
      more: 'Pitch your idea at 12:30, use the mentor sessions from 14:30, complete the progress check-in survey at 20:30 and send your Loom report by midnight.',
      days: [EVENT.days[0]], join: 'Join the kickoff',
    };
  } else if (kind === 'day2') {
    e = {
      subject: 'Final day: submissions close at 11:00',
      preview: 'Submit by 11:00, present at 12:00, winner at 14:00.',
      eyebrow: 'Day 2 · Saturday 14 November',
      heading: 'Final day, ' + esc(first) + '. Let’s see it working.',
      intro: 'Submit your project by <strong style="color:#0a1733">11:00 SAST</strong>, ' +
        'final presentations start at <strong style="color:#0a1733">12:00</strong> and the winner is announced at <strong style="color:#0a1733">14:00</strong>.',
      introText: 'Submit your project by 11:00 SAST, final presentations start at 12:00 and the winner is announced at 14:00.',
      more: 'You’ve built something real. We can’t wait to see it.',
      days: [EVENT.days[1]], join: 'Join the presentations',
    };
  } else {
    throw new Error('Unknown reminder: ' + kind);
  }

  const rsvp = e.rsvp ? rsvpBlock_(r.email, e.rsvp) : null;
  const text = ['Hi ' + first + ',', '', e.introText, '']
    .concat(rsvp ? [rsvp.text, ''] : [])
    .concat(e.join ? [joinText_(e.join), ''] : [])
    .concat(e.steps ? [stepsText_(e.steps).replace(/’/g, "'"), ''] : [])
    .concat([e.more.replace(/’/g, "'"), ''])
    .concat(e.link ? [e.link[0] + ': ' + e.link[1], ''] : [])
    .concat(e.days ? [agendaText_(e.days), ''] : [])
    .concat(e.calendar ? ['Add it to your calendar:', '  Day 1: ' + calendarUrl_(EVENT.sessions[0]), '  Day 2: ' + calendarUrl_(EVENT.sessions[1]), ''] : [])
    .concat([discordText])
    .concat(signText)
    .join('\n');

  const html = emailLayout_({
    preview: e.preview,
    eyebrow: e.eyebrow,
    heading: e.heading,
    sections: [
      section_('<p style="margin:0">' + e.intro + '</p>', '32px 36px 0'),
    ].concat(rsvp ? rsvp.html : [])
      .concat(e.join ? [joinHtml_(e.join)] : [])
      .concat(e.steps ? [stepsHtml_(e.steps)] : [])
      .concat([section_('<p style="margin:0">' + esc(e.more) + '</p>', '16px 36px 8px')])
      .concat(e.link ? [pageLinkHtml_(e.link[0], e.link[1])] : [])
      .concat(e.days ? [agendaHtml_(e.days.length > 1 ? 'Both days at a glance' : 'Today', e.days)] : [])
      .concat(e.calendar ? [calendarButtons_()] : []).concat([
      discordHtml,
      signOff_(),
    ]),
  });

  return { subject: e.subject, text: text, html: html };
}

/** Run from the editor to send yourself all six reminders (sample details). */
function sendTestReminders() {
  const me = Session.getActiveUser().getEmail();
  REMINDERS.forEach(r => sendToApplicant_(me, reminderEmail_({ name: 'Thandi Mokoena', email: me }, r.type)));
  Logger.log('Test reminders sent to ' + me);
}

function welcomeEmail_(r) {
  const first = firstName_(r.name);
  const days = Math.ceil((EVENT.kickoff.getTime() - Date.now()) / 86400000);
  const countdown = days > 1 ? 'Just ' + days + ' days to go until kickoff.'
    : days === 1 ? 'Kickoff is tomorrow!'
    : 'Kickoff is almost here!';
  const subject = 'We’re so excited you applied for Colleague Zero, ' + first + '!';

  const text = [
    'Hi ' + first + ',',
    '',
    "Now that your application is in, we just have to say it: we are so excited for you.",
    '',
    'Colleague Zero is 24 hours of building something that has never existed before: a digital colleague ' +
      'that plans, works, checks itself and gets real work done for education. We love that you want to be part of it.',
    '',
    countdown + ' The best way to get ready is to join the Melsoft Academy community on Discord. ' +
      "Say hi, meet the other builders, find teammates if you're still looking, and catch every announcement as it lands.",
    '',
    'Join us on Discord: ' + CONFIG.DISCORD_URL,
    '',
    "Come with your curiosity, your ideas and your best energy. We can't wait to see what you build.",
    '',
    'See you on Discord,',
    'The Melsoft Academy team',
  ].join('\n');

  const html = emailLayout_({
    preview: 'We can’t wait to see what you build. Come and say hi on Discord.',
    eyebrow: countdown,
    heading: 'We’re so excited for you, ' + escapeHtml_(first) + '!',
    sections: [
      section_(
        '<p style="margin:0 0 16px">Now that your application is in, we just have to say it: ' +
        '<strong style="color:#0a1733">we are so excited for you.</strong></p>' +
        '<p style="margin:0 0 16px">Colleague Zero is 24 hours of building something that has never existed before: a digital colleague ' +
        'that plans, works, checks itself and gets real work done for education. We love that you want to be part of it.</p>' +
        '<p style="margin:0">The best way to get ready is to join the <strong style="color:#0a1733">Melsoft Academy community on Discord</strong>. ' +
        'Say hi, meet the other builders, find teammates if you’re still looking, and catch every announcement as it lands.</p>',
        '32px 36px 8px'),
      button_(CONFIG.DISCORD_URL, 'Join us on Discord', '#5865F2'),
      section_(
        '<p style="margin:0 0 16px">Come with your curiosity, your ideas and your best energy. We can’t wait to see what you build.</p>' +
        '<p style="margin:0">See you on Discord,<br><strong style="color:#0a1733">The Melsoft Academy team</strong></p>',
        '24px 36px 32px'),
    ],
  });

  return { subject: subject, text: text, html: html };
}

/**
 * Keep the light design when the reader's phone is in dark mode.
 * - color-scheme "light only": Apple Mail (iPhone/Mac) and Outlook stop recolouring the email.
 * - Gmail's apps ignore that and repaint colours themselves, but leave background images alone,
 *   so the navy areas are also painted as a one-colour gradient. White text on navy is wrapped
 *   in the "screen"/"difference" blend pair, which cancels Gmail's text inversion
 *   (the `u + .body` selector only matches inside Gmail).
 */
const EMAIL_HEAD =
  '<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
  '<meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light only">' +
  '<style>:root{color-scheme:light only;supported-color-schemes:light only}' +
  'u + .body .gm-screen{background:#000;mix-blend-mode:screen}' +
  'u + .body .gm-diff{background:#000;mix-blend-mode:difference}</style></head>';

/** Navy that Gmail's dark mode can't recolour. */
function navyBg_(colour) {
  return 'background-color:' + colour + ';background-image:linear-gradient(' + colour + ',' + colour + ');';
}

/** White text that stays white on a locked dark background in Gmail's dark mode. */
function keepWhite_(html, display) {
  const d = display || 'block';
  return '<div class="gm-screen" style="display:' + d + '"><div class="gm-diff" style="display:' + d + '">' + html + '</div></div>';
}

/** Shared email shell: navy header band, white card, footer. Values in opts must already be escaped. */
function emailLayout_(opts) {
  return '<!doctype html><html>' + EMAIL_HEAD + '<body class="body" style="margin:0;padding:0;background:#eef1f7">' +
    // Preview line shown next to the subject in most inboxes.
    '<div style="display:none;max-height:0;overflow:hidden">' + opts.preview + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef1f7;padding:32px 12px">' +
    '<tr><td align="center">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#ffffff;border-radius:20px;overflow:hidden">' +
    '<tr><td style="' + navyBg_('#0a1733') + 'padding:32px 36px">' +
    '<div style="font:700 12px/1 Inter,Arial,sans-serif;letter-spacing:.14em;text-transform:uppercase;color:#eab0d4">' + opts.eyebrow + '</div>' +
    keepWhite_('<div style="margin-top:12px;font:800 30px/1.1 Inter,Arial,sans-serif;letter-spacing:-.02em;color:#ffffff">' + opts.heading + '</div>') +
    '</td></tr>' +
    opts.sections.join('') +
    '</table>' +
    '<div style="max-width:600px;padding:16px 8px;font:12px/1.5 Inter,Arial,sans-serif;color:#8c9bbb">' +
    'You’re receiving this because you applied for Colleague Zero. Melsoft Academy (PTY) LTD, 173 Oxford Road, Rosebank, Sandton, 2196.</div>' +
    '</td></tr></table></body></html>';
}

function section_(innerHtml, padding) {
  return '<tr><td style="padding:' + padding + ';font:16px/1.65 Inter,Arial,sans-serif;color:#34405e">' + innerHtml + '</td></tr>';
}

function button_(url, label, colour) {
  return '<tr><td style="padding:24px 36px 8px">' +
    '<a href="' + escapeHtml_(url) + '" style="display:inline-block;padding:14px 24px;border-radius:999px;' + navyBg_(colour) + 'color:#ffffff;' +
    'font:700 15px/1 Inter,Arial,sans-serif;text-decoration:none">' + keepWhite_(escapeHtml_(label), 'inline-block') + '</a>' +
    '</td></tr>';
}

function signOff_() {
  return section_(
    '<p style="margin:0 0 16px">Questions? Call the Melsoft Academy team on ' +
    '<a href="tel:+27101584346" style="color:#5e0743;font-weight:700;text-decoration:none">' + EVENT.phone + '</a>.</p>' +
    '<p style="margin:0">See you soon,<br><strong style="color:#0a1733">The Melsoft Academy team</strong></p>',
    '24px 36px 32px');
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ---------------------------------------------------------------------------

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
