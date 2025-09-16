/***** ROUTING that ignores column G (ROUTED_AT) and writes 6 columns to every destination *****/
const DROP_OFF_SHEET = 'DROP_OFF';
const CONFIG_SHEET   = 'CONFIG';
const LOG_SHEET      = 'LOG';

const ORDERS_HEADER  = ['PART','LOC','CUSTM','PRICE','DATE','DESCR']; // A–F
const ACCOUNT_HEADER = ['PART','LOC','CUSTM','PRICE'];                 // A–D

/** Installable on-edit handler — handles multi-row pastes */
function routeDropOffOnEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== DROP_OFF_SHEET) return;

    // Only react if the edit overlaps columns A–F (1..6)
    const c1 = e.range.getColumn();
    const cN = c1 + e.range.getNumColumns() - 1;
    if (cN < 1 || c1 > 6) return;

    const startRow = e.range.getRow();
    const numRows  = e.range.getNumRows();

    for (let r = 0; r < numRows; r++) {
      routeRowSafe_(startRow + r);
    }
  } catch (err) {
    console.error('routeDropOffOnEdit error:', err);
  }
}

/** Routes a single DROP_OFF row with locking; duplicates allowed across rows */
function routeRowSafe_(row) {
  withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    const drop = ss.getSheetByName(DROP_OFF_SHEET);
    if (!drop) return;

    // Read A–F only
    const vals = drop.getRange(row, 1, 1, 6).getValues()[0];
    const [part, loc, custm, price, date, descr] = vals;

    // Require ALL A–F present before routing (supports manual entry)
    if (![part, loc, custm, price, date, descr].every(v => v !== '' && v !== null)) return;

    // Only skip if THIS ROW was already routed
    if (rowWasLogged_(row)) return;

    // Resolve destination from CONFIG
    const dest = lookupDestination_(custm); // {name, isDefault}
    if (!dest || !dest.name) return;

    const target = getOrCreateSheet_(dest.name);

    // Headers: ORDERS gets full A–F; all other tabs only enforce A–D (but we still write A–F)
    const isOrders = String(dest.name).trim().toUpperCase() === 'ORDERS' || dest.isDefault === true;
    if (isOrders) {
      ensureHeaders_(target, ORDERS_HEADER);
    } else {
      ensureHeaders_(target, ACCOUNT_HEADER);
    }

    // Find next empty row considering 6 columns and write A–F to every destination
    const next = nextRowByFirstNCols_(target, 6);
    target.getRange(next, 1, 1, 6).setValues([[part, loc, custm, price, date, descr]]);

    // LOG every routing event (duplicates allowed across rows)
    appendLog_(buildKey_(vals), dest.name, row);

    // Optional toast
    SpreadsheetApp.getActive().toast(`Routed row ${row} → ${dest.name}`, 'Drop-off', 3);
  });
}

/** Find the first row whose first N columns (starting at A) are all empty */
function nextRowByFirstNCols_(sheet, N) {
  const last = sheet.getLastRow();
  if (last < 2) return 2; // headers only → first data row

  const width = N;
  const height = Math.max(0, last - 1);
  const rng = sheet.getRange(2, 1, height, width);
  const values = rng.getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const rowHasData = values[i].some(v => v !== '' && v !== null);
    if (rowHasData) return 2 + i + 1; // first empty row after the last data row in A..N
  }
  return 2; // no data in A..N
}

/** CONFIG lookup: returns {name, isDefault} */
function lookupDestination_(custm) {
  const ss = SpreadsheetApp.getActive();
  const cfg = ss.getSheetByName(CONFIG_SHEET);
  if (!cfg) return null;
  const last = cfg.getLastRow();
  if (last < 2) return null;

  const rows = cfg.getRange(2, 1, last - 1, 3).getValues(); // A:C
  let defaultTarget = null;

  for (const [_, target, isDefault] of rows) {
    if (String(isDefault).toUpperCase() === 'TRUE' && target) defaultTarget = target;
  }
  for (const [code, target] of rows) {
    if (!code) continue;
    if (String(code).trim().toUpperCase() === String(custm).trim().toUpperCase()) {
      return { name: (target || defaultTarget || 'ORDERS').toString(), isDefault: false };
    }
  }
  return { name: (defaultTarget || 'ORDERS').toString(), isDefault: true };
}

/** Get or create tab; tolerant of case/whitespace differences */
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const wanted = String(name).trim();
  let sh = ss.getSheetByName(wanted);
  if (!sh) {
    const match = ss.getSheets().map(s => s.getName())
      .find(n => n.trim().toUpperCase() === wanted.toUpperCase());
    sh = match ? ss.getSheetByName(match) : ss.insertSheet(wanted);
  }
  return sh;
}

/** Ensure the first headerWidth columns in row 1 match the expected header */
function ensureHeaders_(sheet, header) {
  const width = header.length;
  const current = sheet.getRange(1, 1, 1, width).getValues()[0].map(v => String(v).toUpperCase());
  const wanted  = header.map(v => v.toUpperCase());
  if (current.join('|') !== wanted.join('|')) {
    sheet.getRange(1, 1, 1, width).setValues([header]);
  }
}

/************ LOG helpers (duplicates allowed) ************/
function buildKey_(arrAtoF) {
  return arrAtoF.map(v => String(v).trim()).join('||').toUpperCase();
}
function rowWasLogged_(row) {
  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName(LOG_SHEET);
  if (!log) return false;
  const last = log.getLastRow();
  if (last < 2) return false;
  const rows = log.getRange(2, 4, last - 1, 1).getValues().flat(); // column D = ROW
  return rows.some(v => Number(v) === Number(row));
}
function appendLog_(key, destName, row) {
  const ss = SpreadsheetApp.getActive();
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.getRange(1, 1, 1, 4).setValues([['KEY','WHEN','DEST','ROW']]);
  }
  log.appendRow([key, new Date(), destName, row]);
}

/** Prevent concurrent double-appends when multiple users paste at once */
function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.tryLock(5000);
  try { fn(); } finally { lock.releaseLock(); }
}
