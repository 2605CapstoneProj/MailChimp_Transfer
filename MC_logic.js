// ============================================================
//  CAMPAIGN LAUNCHER — MC_logic.gs
//  Syncs filtered GSheet contacts to Mailchimp via API
// ============================================================


// ============================================================
//  CONFIG — edit this section before deploying
// ============================================================
const CONFIG = {
  MC_API_KEY:    'YOUR_MC_API_KEY',         // e.g. 'abc123def456-us1'
  MC_LIST_ID:    'YOUR_AUDIENCE_LIST_ID',   // MC > Audience > Settings > Audience ID

  SHEET_NAME:    'Active Leads',            // tab name to read contacts from

  // Column header names — must match sheet exactly (case-sensitive)
  COL_EMAIL:     'Email Address',
  COL_FNAME:     'First Name',
  COL_LNAME:     'Last Name',
  COL_COMPANY:   'Company',

  // Columns to expose as filters in the popup
  // Only use "clean" columns with consistent dropdown-style values
  FILTERABLE_COLS: ['Machine Type', 'Priority', 'State'],

  // Low count threshold — values with count <= this are flagged as possible typos
  LOW_COUNT_FLAG: 3,

  LOG_SHEET_NAME: 'Sync Log',
};


// ============================================================
//  MENU — adds "Campaign Launcher" to the GSheet toolbar
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Campaign Launcher')
    .addItem('Open launcher...', 'openLauncher')
    .addToUi();
}


// ============================================================
//  OPEN DIALOG
// ============================================================
function openLauncher() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(500)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Campaign Launcher');
}


// ============================================================
//  GET FILTER OPTIONS
//  Called from Dialog on load.
//  Returns: { "Machine Type": [{value, count}, ...], ... }
// ============================================================
function getFilterOptions() {
  const sheet = _getSheet();
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  const result = {};

  CONFIG.FILTERABLE_COLS.forEach(colName => {
    const colIdx = headers.indexOf(colName);
    if (colIdx === -1) return; // column not found — skip silently

    const counts = {};
    rows.forEach(row => {
      const val = String(row[colIdx]).trim();
      if (!val || val === 'undefined' || val === 'null') return;
      counts[val] = (counts[val] || 0) + 1;
    });

    // Sort by count descending so most common values appear first
    result[colName] = Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .map(([value, count]) => ({ value, count }));
  });

  return { options: result, lowCountFlag: CONFIG.LOW_COUNT_FLAG };
}


// ============================================================
//  PREVIEW COUNT
//  Called from Dialog when filter selection changes.
//  filters: { "Machine Type": ["A","B"], "Priority": ["1"] }
//  Returns: { total, withEmail, noEmail }
// ============================================================
function previewContacts(filters) {
  const all = _getMatchedContacts(filters);
  const withEmail = all.filter(c => c.email).length;
  return {
    total:     all.length,
    withEmail: withEmail,
    noEmail:   all.length - withEmail,
  };
}


// ============================================================
//  SYNC TO MAILCHIMP
//  Called from Dialog on sync button click.
//  Returns: { synced, skipped, errors[] }
// ============================================================
function syncToMailchimp(filters, campaignTag) {
  const contacts = _getMatchedContacts(filters).filter(c => c.email);
  if (contacts.length === 0) return { synced: 0, skipped: 0, errors: [] };

  const dc      = CONFIG.MC_API_KEY.split('-').pop();         // e.g. 'us1'
  const baseUrl = `https://${dc}.api.mailchimp.com/3.0/lists/${CONFIG.MC_LIST_ID}/members`;
  const auth    = 'Basic ' + Utilities.base64Encode('anystring:' + CONFIG.MC_API_KEY);
  const tags    = campaignTag
    ? campaignTag.split(',').map(t => t.trim()).filter(Boolean)
    : [];

  let synced = 0, skipped = 0;
  const errors = [];

  contacts.forEach(contact => {
    try {
      const emailHash = _md5(contact.email);

      const payload = {
        email_address: contact.email,
        status_if_new: 'subscribed',  // only sets status for brand-new contacts
        merge_fields:  {},
      };
      if (contact.fname)   payload.merge_fields.FNAME   = contact.fname;
      if (contact.lname)   payload.merge_fields.LNAME   = contact.lname;
      if (contact.company) payload.merge_fields.COMPANY = contact.company;

      const res = UrlFetchApp.fetch(`${baseUrl}/${emailHash}`, {
        method:            'PUT',
        headers:           { 'Authorization': auth, 'Content-Type': 'application/json' },
        payload:           JSON.stringify(payload),
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
//  INTERNAL — apply tags to a contact (separate MC endpoint)
// ============================================================
function _applyTags(emailHash, tags, auth, dc) {
  const url = `https://${dc}.api.mailchimp.com/3.0/lists/${CONFIG.MC_LIST_ID}/members/${emailHash}/tags`;
  UrlFetchApp.fetch(url, {
    method:            'POST',
    headers:           { 'Authorization': auth, 'Content-Type': 'application/json' },
    payload:           JSON.stringify({ tags: tags.map(name => ({ name, status: 'active' })) }),
    muteHttpExceptions: true,
  });
}


// ============================================================
//  INTERNAL — get matched + deduped contacts from sheet
//  Filter logic: AND across columns, OR within a column
// ============================================================
function _getMatchedContacts(filters) {
  const sheet   = _getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  const emailIdx   = headers.indexOf(CONFIG.COL_EMAIL);
  const fnameIdx   = headers.indexOf(CONFIG.COL_FNAME);
  const lnameIdx   = headers.indexOf(CONFIG.COL_LNAME);
  const companyIdx = headers.indexOf(CONFIG.COL_COMPANY);

  const seen    = new Set();
  const matched = [];

  rows.forEach(row => {
    // Must pass ALL filter groups (AND logic across columns)
    const passes = Object.entries(filters).every(([colName, selectedVals]) => {
      if (!selectedVals || selectedVals.length === 0) return true; // no selection = no restriction
      const idx = headers.indexOf(colName);
      if (idx === -1) return true;
      return selectedVals.includes(String(row[idx]).trim());
    });

    if (!passes) return;

    const rawEmail = emailIdx >= 0 ? String(row[emailIdx]).trim().toLowerCase() : '';
    const validEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(rawEmail);

    // Still include in count even without valid email (for preview noEmail count)
    // but mark email as null so sync step skips them
    const email = validEmail ? rawEmail : null;

    if (email && seen.has(email)) return; // dedupe by email
    if (email) seen.add(email);

    matched.push({
      email,
      fname:   fnameIdx   >= 0 ? String(row[fnameIdx]).trim()   : '',
      lname:   lnameIdx   >= 0 ? String(row[lnameIdx]).trim()   : '',
      company: companyIdx >= 0 ? String(row[companyIdx]).trim() : '',
    });
  });

  return matched;
}


// ============================================================
//  INTERNAL — get the configured sheet (throws if missing)
// ============================================================
function _getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" not found. Check SHEET_NAME in CONFIG.`);
  return sheet;
}


// ============================================================
//  INTERNAL — MD5 hash (required by MC member endpoint)
// ============================================================
function _md5(str) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}


// ============================================================
//  LOG — writes one row per sync run to a Sync Log tab
// ============================================================
function _ensureLogSheet() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   log = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!log) {
    log = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
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
