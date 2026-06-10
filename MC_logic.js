// ============================================================
//  CAMPAIGN LAUNCHER — MailChimp_Sync.gs
//  Syncs filtered GSheet contacts to Mailchimp via API
// ============================================================


// ============================================================
//  CONFIG_MC — edit this section before deploying
// ============================================================
const CONFIG_MC = {
  MC_API_KEY:  '',       // e.g. 'abc123def456-us1'
  MC_LIST_ID:  '', // MC > Audience > Settings > Audience ID

  // Leave null if this script is bound directly to the sheet (normal case).
  // Fill in the Spreadsheet ID only if the script lives in a different GSheet file.
  SPREADSHEET_ID: '',                 // e.g. '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms'

  SHEET_NAME: 'Leads',           // exact tab name

  // ── Contact column headers (must match sheet exactly, case-sensitive) ──
  COL_EMAIL:   'Email Address',
  COL_NAME:    'Contact Name',        // single full-name column → maps to FNAME in MC
  COL_COMPANY: 'Business Name',     // set to null if there is no company column

  // ── Columns to show as filters in the popup ──
  // Use only columns with consistent, limited values (dropdown-style)
  FILTERABLE_COLS: ['Status', 'Equipment', ],

  // Values with count ≤ this are flagged amber as possible typos in the UI
  LOW_COUNT_FLAG: 3,

  LOG_SHEET_NAME: 'MailChimp Sync Log',
};


// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Campaign Launcher')
    .addItem('Open launcher...', 'openLauncher')
    .addToUi();
}


// ============================================================
//  OPEN SIDEBAR DIALOG
// ============================================================
function openLauncher() {
  const html = HtmlService.createHtmlOutputFromFile('MC_pop')
    .setWidth(500)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Campaign Launcher');
}


// ============================================================
//  WEB APP ENTRY POINT — for MC Test Panel
//  Deploy as Web App: Execute as "Me", Access "Only myself"
//  Then open the /exec URL in browser
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('MC_tester')
    .setTitle('MC Test Panel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}


// ============================================================
//  GET FILTER OPTIONS
//  Returns: { options: { "Col": [{value, count}] }, lowCountFlag }
// ============================================================
function getFilterOptions() {
  const sheet   = _getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  const result = {};

  CONFIG_MC.FILTERABLE_COLS.forEach(colName => {
    const colIdx = headers.indexOf(colName);
    if (colIdx === -1) return;

    const counts = {};
    rows.forEach(row => {
      const val = String(row[colIdx]).trim();
      if (!val || val === 'undefined' || val === 'null') return;
      counts[val] = (counts[val] || 0) + 1;
    });

    result[colName] = Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .map(([value, count]) => ({ value, count }));
  });

  return { options: result, lowCountFlag: CONFIG_MC.LOW_COUNT_FLAG };
}


// ============================================================
//  PREVIEW COUNT
//  Returns: { total, withEmail, noEmail }
// ============================================================
function previewContacts(filters) {
  const all = _getMatchedContacts(filters);
  const withEmail = all.filter(c => c.email).length;
  return { total: all.length, withEmail, noEmail: all.length - withEmail };
}


// ============================================================
//  SYNC TO MAILCHIMP
//  Returns: { synced, skipped, errors[] }
// ============================================================
function syncToMailchimp(filters, campaignTag) {
  const contacts = _getMatchedContacts(filters).filter(c => c.email);
  if (contacts.length === 0) return { synced: 0, skipped: 0, errors: [] };

  const dc      = CONFIG_MC.MC_API_KEY.split('-').pop();
  const baseUrl = `https://${dc}.api.mailchimp.com/3.0/lists/${CONFIG_MC.MC_LIST_ID}/members`;
  const auth    = 'Basic ' + Utilities.base64Encode('anystring:' + CONFIG_MC.MC_API_KEY);
  const tags    = campaignTag
    ? campaignTag.split(',').map(t => t.trim()).filter(Boolean)
    : [];

  let synced = 0, skipped = 0;
  const errors = [];

  contacts.forEach(contact => {
    try {
      const emailHash = _md5(contact.email);
      const payload   = {
        email_address: contact.email,
        status_if_new: 'subscribed',
        merge_fields:  {},
      };

      if (contact.name)    payload.merge_fields.FNAME   = contact.name;
      if (contact.company) payload.merge_fields.COMPANY = contact.company;

      const res  = UrlFetchApp.fetch(`${baseUrl}/${emailHash}`, {
        method:             'PUT',
        headers:            { 'Authorization': auth, 'Content-Type': 'application/json' },
        payload:            JSON.stringify(payload),
        muteHttpExceptions: true,
      });
      const code = res.getResponseCode();

      if (code === 200 || code === 201) {
        if (tags.length > 0) _applyTags(emailHash, tags, auth, dc);
        synced++;
      } else {
        const body = JSON.parse(res.getContentText());
        errors.push(`${contact.email}: ${body.detail || 'HTTP ' + code}`);
        skipped++;
      }
    } catch (e) {
      errors.push(`${contact.email}: ${e.message}`);
      skipped++;
    }
  });

  _writeLog(filters, campaignTag, contacts.length, synced, skipped, errors);
  return { synced, skipped, errors };
}


// ============================================================
//  TEST PANEL FUNCTIONS (called from TestPanel.html)
// ============================================================

/** Verify API key — returns account name + plan */
function testGetAccount() {
  const { auth, dc } = _mcAuth();
  const res  = UrlFetchApp.fetch(`https://${dc}.api.mailchimp.com/3.0/`, {
    headers: { 'Authorization': auth }, muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error(body.detail || 'Auth failed');
  return {
    account_name: body.account_name,
    email:        body.email,
    plan:         body.account_industry || '—',
  };
}

/** List all audiences — use this to find your MC_LIST_ID */
function testListAudiences() {
  const { auth, dc } = _mcAuth();
  const res  = UrlFetchApp.fetch(
    `https://${dc}.api.mailchimp.com/3.0/lists?count=20&fields=lists.id,lists.name,lists.stats.member_count`,
    { headers: { 'Authorization': auth }, muteHttpExceptions: true }
  );
  const body = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error(body.detail || 'Failed to fetch audiences');
  return body.lists.map(l => ({
    id:      l.id,
    name:    l.name,
    members: l.stats.member_count,
  }));
}

/** Look up a single contact by email */
function testFindContact(email) {
  const { auth, dc } = _mcAuth();
  const hash = _md5(email.trim().toLowerCase());
  const res  = UrlFetchApp.fetch(
    `https://${dc}.api.mailchimp.com/3.0/lists/${CONFIG_MC.MC_LIST_ID}/members/${hash}` +
    `?fields=email_address,status,merge_fields,tags`,
    { headers: { 'Authorization': auth }, muteHttpExceptions: true }
  );
  const body = JSON.parse(res.getContentText());
  if (res.getResponseCode() === 404) return { found: false };
  if (res.getResponseCode() !== 200) throw new Error(body.detail || 'Lookup failed');
  return {
    found:    true,
    email:    body.email_address,
    status:   body.status,
    name:     body.merge_fields?.FNAME || '—',
    company:  body.merge_fields?.COMPANY || '—',
    tags:     (body.tags || []).map(t => t.name),
  };
}

/** Add a test contact to verify sync works end-to-end */
function testAddContact(email, name) {
  const { auth, dc } = _mcAuth();
  const hash    = _md5(email.trim().toLowerCase());
  const payload = {
    email_address: email.trim().toLowerCase(),
    status_if_new: 'subscribed',
    merge_fields:  { FNAME: name || 'Test Contact' },
  };
  const res  = UrlFetchApp.fetch(
    `https://${dc}.api.mailchimp.com/3.0/lists/${CONFIG_MC.MC_LIST_ID}/members/${hash}`,
    {
      method:             'PUT',
      headers:            { 'Authorization': auth, 'Content-Type': 'application/json' },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true,
    }
  );
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText());
  if (code !== 200 && code !== 201) throw new Error(body.detail || 'HTTP ' + code);
  return { ok: true, status: body.status, email: body.email_address };
}

// Quick Test Call
function runTests() {
  console.log('=== Account ===');
  console.log(JSON.stringify(testGetAccount()));
  
  console.log('=== Audiences ===');
  console.log(JSON.stringify(testListAudiences()));
}


// ============================================================
//  INTERNAL — helpers
// ============================================================

function _mcAuth() {
  const dc   = CONFIG_MC.MC_API_KEY.split('-').pop();
  const auth = 'Basic ' + Utilities.base64Encode('anystring:' + CONFIG_MC.MC_API_KEY);
  return { auth, dc };
}

function _getSheet() {
  const ss = CONFIG_MC.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG_MC.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_MC.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG_MC.SHEET_NAME}" not found. Check SHEET_NAME in CONFIG_MC.`);
  return sheet;
}

function _getMatchedContacts(filters) {
  const sheet   = _getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  const emailIdx   = headers.indexOf(CONFIG_MC.COL_EMAIL);
  const nameIdx    = CONFIG_MC.COL_NAME    ? headers.indexOf(CONFIG_MC.COL_NAME)    : -1;
  const companyIdx = CONFIG_MC.COL_COMPANY ? headers.indexOf(CONFIG_MC.COL_COMPANY) : -1;

  const seen    = new Set();
  const matched = [];

  rows.forEach(row => {
    const passes = Object.entries(filters).every(([colName, selectedVals]) => {
      if (!selectedVals || selectedVals.length === 0) return true;
      const idx = headers.indexOf(colName);
      if (idx === -1) return true;
      return selectedVals.includes(String(row[idx]).trim());
    });

    if (!passes) return;

    const rawEmail   = emailIdx >= 0 ? String(row[emailIdx]).trim().toLowerCase() : '';
    const validEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(rawEmail);
    const email      = validEmail ? rawEmail : null;

    if (email && seen.has(email)) return;
    if (email) seen.add(email);

    matched.push({
      email,
      name:    nameIdx    >= 0 ? String(row[nameIdx]).trim()    : '',
      company: companyIdx >= 0 ? String(row[companyIdx]).trim() : '',
    });
  });

  return matched;
}

function _applyTags(emailHash, tags, auth, dc) {
  const url = `https://${dc}.api.mailchimp.com/3.0/lists/${CONFIG_MC.MC_LIST_ID}/members/${emailHash}/tags`;
  UrlFetchApp.fetch(url, {
    method:             'POST',
    headers:            { 'Authorization': auth, 'Content-Type': 'application/json' },
    payload:            JSON.stringify({ tags: tags.map(name => ({ name, status: 'active' })) }),
    muteHttpExceptions: true,
  });
}

function _md5(str) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function _ensureLogSheet() {
  const ss  = CONFIG_MC.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG_MC.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(CONFIG_MC.LOG_SHEET_NAME);
  if (!log) {
    log = ss.insertSheet(CONFIG_MC.LOG_SHEET_NAME);
    log.appendRow(['Timestamp', 'Campaign Tag', 'Filters Applied', 'Total Matched', 'Synced OK', 'Skipped', 'Errors']);
    log.setFrozenRows(1);
    log.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return log;
}

function _writeLog(filters, campaignTag, total, synced, skipped, errors) {
  const log = _ensureLogSheet();
  log.appendRow([
    new Date(),
    campaignTag || '(none)',
    JSON.stringify(filters),
    total,
    synced,
    skipped,
    errors.length > 0 ? errors.join(' | ') : '',
  ]);
}
