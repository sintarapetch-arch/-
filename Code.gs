/**
 * Google Apps Script for Pakorn Technical Supply Ltd., Part. (หจก. ปกรณ์ เทคนิคอล ซัพพลาย)
 * Engineering Task Management System API - Fail-Safe CORS & JSONP Supported
 *
 * v6: No preview/formula columns in the sheet at all.
 *     The three attachment columns hold readable, clickable links (rich text):
 *
 *       Site_Photos      📷 รูปที่ 1        <- blue link, opens the photo in Drive
 *       Document_Files   📄 report.pdf · 1.2 MB
 *       Video_Files      🎬 site-clip.mp4 · 4 MB
 *
 *     Visual previewing happens in the web app; the sheet just links out.
 *     The URLs live in each line's link metadata, so the app reads them back
 *     while the sheet stays human-readable.
 *
 * v7: Concurrency and identity fixes.
 *     - Every write runs inside a LockService lock, so two people saving at the
 *       same moment can no longer append onto each other or delete a row while
 *       another request is still counting rows. Drive uploads stay outside the
 *       lock so one big photo batch does not block the whole team.
 *     - Job_ID is now "highest existing number + 1" instead of "row count + 1".
 *       The old formula repeated an ID after any deletion, and because UPDATE
 *       and DELETE stop at the first match, every later edit silently landed on
 *       the wrong job. CREATE also re-issues an ID if the one sent already exists.
 *     - Text starting with "=" is stored as text instead of becoming a formula.
 *
 * Deployment Instructions:
 * 1. Open your Google Sheet.
 * 2. Click Extensions > Apps Script.
 * 3. Replace all code in Code.gs with this script, then press Ctrl+S.
 * 4. Select the function "authorizeDriveAccess" and click Run once (grants Drive).
 * 5. Deploy > Manage deployments > ✏️ > Version: "New version" > Deploy.
 */

const SHEET_NAME = "Engineering_Tasks";
const DRIVE_ROOT_FOLDER_NAME = "PTS_Engineering_Attachments";
const MIN_ROW_HEIGHT = 21;
const MAX_ROW_HEIGHT = 320;

/**
 * ⚡ รันฟังก์ชันนี้ใน Apps Script เพื่อสร้าง/อัปเดตหัวคอลัมน์ใน Google Sheet ทันที
 */
function setupSheetHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  CacheService.getScriptCache().remove(HEADER_OK_KEY);
  const headers = ensureHeaders(sheet);
  invalidateCache();
  SpreadsheetApp.flush();
  Logger.log("✅ อัปเดตหัวตารางเรียบร้อย: " + headers.join(", "));
  return "✅ อัปเดตหัวตารางเรียบร้อย (" + headers.length + " คอลัมน์)";
}

/**
 * ⚡ รันฟังก์ชันนี้ใน Apps Script เพื่อแปลงวันที่ทุกแถวใน Google Sheet ให้เป็นมาตรฐาน วัน/เดือน/ปี (dd/MM/yyyy) ทันที
 */
function formatAllDatesInSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return "❌ ไม่พบ Sheet " + SHEET_NAME;

  const headers = getHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "ℹ️ ไม่มีข้อมูลในตาราง";

  const dateColIndexes = [];
  DATE_COLUMNS.forEach(function (colName) {
    const idx = headers.indexOf(colName);
    if (idx !== -1) dateColIndexes.push(idx);
  });

  if (dateColIndexes.length === 0) return "❌ ไม่พบคอลัมน์วันที่";

  const range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  const values = range.getValues();
  let count = 0;

  for (let r = 0; r < values.length; r++) {
    dateColIndexes.forEach(function (c) {
      const orig = values[r][c];
      if (orig !== null && orig !== undefined && orig !== "") {
        const formatted = formatDateForSheet(orig);
        if (formatted && formatted !== orig) {
          values[r][c] = formatted;
          count++;
        }
      }
    });
  }

  range.setValues(values);

  // ตั้งค่ารูปแบบ NumberFormat ของคอลัมน์วันที่ให้แสดงผลเป็น dd/MM/yyyy กึ่งกลาง
  DATE_COLUMNS.forEach(function (name) {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;
    sheet.getRange(2, col, Math.max(lastRow - 1, 1))
      .setNumberFormat("@")
      .setHorizontalAlignment("center");
  });

  invalidateCache();
  SpreadsheetApp.flush();
  Logger.log("✅ แปลงวันที่ในชีตเป็น dd/MM/yyyy เรียบร้อย " + count + " จุด");
  return "✅ แปลงวันที่ในชีตเป็น dd/MM/yyyy เรียบร้อย (" + count + " จุด)";
}

// Short-lived read cache. Every write clears it, so the app still sees its own
// changes instantly - the cache only absorbs repeated polling from the browser.
const DATA_CACHE_KEY = "pts_tasks_payload";
const DATA_CACHE_SECONDS = 15;
const HEADER_OK_KEY = "pts_headers_ok";
const CACHE_MAX_BYTES = 90000; // CacheService rejects values over ~100 KB

// Revision counter for near-realtime sync. Bumped on every change (including
// manual edits in the sheet via onEdit). Clients poll it with ?ping=1, which
// touches no spreadsheet data at all and is therefore very cheap.
const REV_KEY = "pts_revision";

// Drive folder ids, remembered so a save never has to search Drive by name.
const ROOT_FOLDER_ID_KEY = "pts_root_folder_id";
const ROOT_SHARED_KEY = "pts_root_shared";
const FOLDER_CACHE_SECONDS = 21600; // 6 hours

const HEADERS = [
  "Job_ID",
  "Project_Name",
  "Sub_Department",
  "Technician_In_Charge",
  "Task_Detail",
  "Status",
  "JSA_Completed",
  "Priority",
  "Target_Date",
  "Site_Location",
  "Notes_Issues",
  "Updated_At",
  "Site_Photos",
  "PO_Approval_Date",
  "Contract_Expiry_Date",
  "Completion_Date",
  "Document_Files",
  "Video_Files",
  "Site_Contact_Phone",
  "Site_Map_Url",
  "Delivery_Doc"
];

/**
 * Columns stored as rich text instead of plain values: the cell shows a friendly
 * label per file and carries the Drive URL as a real link.
 */
const LINK_COLUMNS = {
  Site_Photos: { icon: "📷", folder: "Photos", kind: "photo" },
  Document_Files: { icon: "📄", folder: "Documents", kind: "file" },
  Video_Files: { icon: "🎬", folder: "Videos", kind: "file" },
  Delivery_Doc: { icon: "📝", folder: "Delivery_Docs", kind: "file" }
};
const LINK_COLUMN_NAMES = Object.keys(LINK_COLUMNS);

// Preview/helper columns created by v5. They are removed automatically now that
// previewing happens in the app instead of the sheet.
const OBSOLETE_HEADERS = [
  "Photo_Preview",
  "Photo_Links",
  "Video_Preview",
  "Video_Links",
  "Document_Links"
];

const SIZE_SEPARATOR = " · ";

const DATE_COLUMNS = [
  "Target_Date",
  "PO_Approval_Date",
  "Contract_Expiry_Date",
  "Completion_Date"
];

function formatDateForSheet(val) {
  if (val === undefined || val === null) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Bangkok", "dd/MM/yyyy");
  }
  const s = String(val).trim();
  if (!s || s === "-" || s === "null" || s === "undefined") return "";

  // Match YYYY-MM-DD
  const isoMatch = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (isoMatch) {
    let year = parseInt(isoMatch[1], 10);
    if (year > 2400) year -= 543;
    const month = String(isoMatch[2]).padStart(2, '0');
    const day = String(isoMatch[3]).padStart(2, '0');
    return `${day}/${month}/${year}`;
  }

  // Match DD/MM/YYYY or DD-MM-YYYY
  const dmyMatch = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dmyMatch) {
    const day = String(dmyMatch[1]).padStart(2, '0');
    const month = String(dmyMatch[2]).padStart(2, '0');
    let year = parseInt(dmyMatch[3], 10);
    if (year > 2400) year -= 543;
    return `${day}/${month}/${year}`;
  }

  return s;
}

/* ---------------------------------------------------------------------------
 * SHEET SETUP
 * ------------------------------------------------------------------------- */

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    ensureHeaders(sheet);
    return sheet;
  }
  ensureHeadersFast(sheet);
  return sheet;
}

/**
 * PERFORMANCE: verifies all required headers exist and runs ensureHeaders if any are missing.
 */
function ensureHeadersFast(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); });

  const complete = HEADERS.every(function (h) { return headers.indexOf(h) !== -1; });
  const clean = OBSOLETE_HEADERS.every(function (h) { return headers.indexOf(h) === -1; });

  if (complete && clean) {
    return;
  }

  ensureHeaders(sheet);
}

/**
 * Adds any missing column, drops the obsolete preview columns, and applies the
 * wrapping needed for the multi-line link cells.
 */
function ensureHeaders(sheet) {
  let lastCol = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); });

  // Remove the v5 preview columns (right to left so indexes stay valid)
  for (let i = headers.length - 1; i >= 0; i--) {
    if (OBSOLETE_HEADERS.indexOf(headers[i]) !== -1) {
      sheet.deleteColumn(i + 1);
      headers.splice(i, 1);
    }
  }

  while (headers.length > 0 && headers[headers.length - 1] === "") {
    headers.pop();
  }

  let added = false;
  HEADERS.forEach(function (h) {
    if (headers.indexOf(h) === -1) {
      headers.push(h);
      added = true;
    }
  });

  if (added) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#991B1B")
    .setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  LINK_COLUMN_NAMES.forEach(function (name) {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;
    const width = sheet.getColumnWidth(col);
    if (width < 200 || width > 400) sheet.setColumnWidth(col, 240);
    sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1))
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment("top");
  });

  DATE_COLUMNS.forEach(function (name) {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;
    sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1))
      .setNumberFormat("@")
      .setHorizontalAlignment("center");
  });

  return headers;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
}

/* ---------------------------------------------------------------------------
 * WEB APP ENTRY POINTS
 * ------------------------------------------------------------------------- */

function doGet(e) {
  try {
    // Lightweight change check - no spreadsheet access, just a counter
    if (e && e.parameter && e.parameter.ping === "1") {
      return jsonResponse({ status: "success", rev: getRevision() }, e);
    }

    // Serve the cached payload when the browser is just polling for changes
    const wantsFresh = e && e.parameter && e.parameter.fresh === "1";
    const cache = CacheService.getScriptCache();
    if (!wantsFresh) {
      const hit = cache.get(DATA_CACHE_KEY);
      if (hit) return rawJsonResponse(hit, e);
    }

    const sheet = getOrCreateSheet();
    const range = sheet.getDataRange();
    const data = range.getValues();

    if (data.length <= 1) {
      return jsonResponse({ status: "success", data: [] }, e);
    }

    const headers = data[0].map(function (h) { return String(h).trim(); });
    const rows = data.slice(1);

    // Attachment URLs live in the cells' link metadata, not in their text
    const linkCols = {};
    LINK_COLUMN_NAMES.forEach(function (name) {
      const idx = headers.indexOf(name);
      if (idx !== -1) linkCols[idx] = name;
    });
    const richValues = Object.keys(linkCols).length > 0 ? range.getRichTextValues() : null;

    const result = rows.map(function (row, rowIdx) {
      const item = {};
      headers.forEach(function (header, index) {
        if (!header) return;

        if (linkCols[index]) {
          const rich = richValues ? richValues[rowIdx + 1][index] : null;
          const config = LINK_COLUMNS[linkCols[index]];
          item[header] = JSON.stringify(
            config.kind === "photo"
              ? extractLinkedUrls(rich, row[index])
              : extractLinkedFiles(rich, row[index], config)
          );
          return;
        }

        let val = row[index];
        if (val instanceof Date) {
          // Timestamps keep their time; plain date fields standardized to dd/MM/yyyy
          val = Utilities.formatDate(val, "Asia/Bangkok",
            header === "Updated_At" ? "dd/MM/yyyy HH:mm" : "dd/MM/yyyy");
        } else if (DATE_COLUMNS.indexOf(header) !== -1 && val) {
          val = formatDateForSheet(val);
        }
        item[header] = val !== undefined && val !== null ? String(val) : "";
      });
      return item;
    }).filter(function (item) {
      return item.Job_ID && String(item.Job_ID).trim() !== "";
    });

    const payload = JSON.stringify({ status: "success", rev: getRevision(), data: result });
    if (payload.length < CACHE_MAX_BYTES) {
      cache.put(DATA_CACHE_KEY, payload, DATA_CACHE_SECONDS);
    }
    return rawJsonResponse(payload, e);

  } catch (error) {
    return jsonResponse({ status: "error", message: error.toString() }, e);
  }
}

function doPost(e) {
  try {
    let requestData;
    if (e && e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.payload) {
      requestData = JSON.parse(e.parameter.payload);
    } else {
      requestData = e ? e.parameter : {};
    }

    const action = requestData.action || "CREATE";
    const sheet = getOrCreateSheet();

    if (action === "CREATE") {
      return handleCreate(sheet, requestData.data, e);
    } else if (action === "UPDATE") {
      return handleUpdate(sheet, requestData.data, e);
    } else if (action === "DELETE") {
      return handleDelete(sheet, requestData.job_id, e);
    } else {
      return jsonResponse({ status: "error", message: "Invalid action type: " + action }, e);
    }

  } catch (error) {
    return jsonResponse({ status: "error", message: error.toString() }, e);
  }
}

/** Single choke point for "something changed": drop the cache and bump the revision. */
function invalidateCache() {
  CacheService.getScriptCache().remove(DATA_CACHE_KEY);
  bumpRevision();
}

function getRevision() {
  try {
    return Number(PropertiesService.getScriptProperties().getProperty(REV_KEY)) || 0;
  } catch (err) {
    return 0;
  }
}

function bumpRevision() {
  try {
    const props = PropertiesService.getScriptProperties();
    const next = (Number(props.getProperty(REV_KEY)) || 0) + 1;
    props.setProperty(REV_KEY, String(next));
    return next;
  } catch (err) {
    return 0;
  }
}

/**
 * Simple trigger: fires when a person edits the sheet by hand. Without this,
 * manual edits would stay invisible to the apps until the next full refresh.
 * Wrapped in try/catch because simple triggers run with reduced authorization.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() !== SHEET_NAME) return;
    CacheService.getScriptCache().remove(DATA_CACHE_KEY);
    bumpRevision();
  } catch (err) {
    // Never let a trigger failure block the user's edit
  }
}

/* ---------------------------------------------------------------------------
 * CRUD HANDLERS
 * ------------------------------------------------------------------------- */

/**
 * Every write runs inside this lock.
 *
 * appendRow / deleteRow / "read the sheet, find the row, write it back" are all
 * read-modify-write sequences. Two technicians pressing Save within the same
 * second used to be able to append onto the same row or delete a row while the
 * other request was still counting rows. Apps Script gives no transactions, so
 * a script lock is the only thing standing between us and a shifted sheet.
 *
 * Slow work (Drive uploads) is deliberately kept OUTSIDE this lock.
 */
function withSheetLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("ระบบกำลังบันทึกข้อมูลของผู้ใช้อื่นอยู่ กรุณากดบันทึกอีกครั้ง");
  }
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (err) { /* already released */ }
  }
}

/** All Job_IDs currently in the sheet, in row order. */
function readJobIds(sheet, headers) {
  const cols = headers || getHeaders(sheet);
  const col = cols.indexOf("Job_ID") + 1;
  const lastRow = sheet.getLastRow();
  if (col <= 0 || lastRow < 2) return [];
  return sheet.getRange(2, col, lastRow - 1, 1).getValues().map(function (r) {
    return String(r[0]).trim();
  });
}

/**
 * Highest existing number + 1.
 *
 * The old formula was "row count + 1", which repeats an ID as soon as anything
 * has been deleted. A duplicate Job_ID is silently destructive here: UPDATE and
 * DELETE both stop at the first match, so every later edit lands on the wrong
 * job and the delete button removes somebody else's row.
 */
function nextJobId(existingIds) {
  let max = 0;
  existingIds.forEach(function (id) {
    const m = /(\d+)\s*$/.exec(id);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return "PTS-ENG-" + String(max + 1).padStart(3, "0");
}

function findRowByJobId(sheet, jobId, col) {
  const lastRow = sheet.getLastRow();
  if (col <= 0 || lastRow < 2) return -1;
  const target = String(jobId).trim();
  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === target) return i + 2;
  }
  return -1;
}

/**
 * The attachment fields as they now stand, with every Base64 payload already
 * replaced by its Drive URL. The app sends big attachment sets in several small
 * requests and feeds this back into the next one, so the files it has already
 * pushed are never uploaded - or re-sent - a second time.
 */
function attachmentSnapshot(task) {
  const out = {};
  LINK_COLUMN_NAMES.forEach(function (name) {
    if (task.hasOwnProperty(name)) out[name] = task[name];
  });
  return out;
}

/** Trusts the row we just appended, but re-finds it if anything shifted. */
function resolveRow(sheet, expectedRow, jobId, headers) {
  const col = headers.indexOf("Job_ID") + 1;
  if (col <= 0) return -1;
  if (expectedRow >= 2 && expectedRow <= sheet.getLastRow() &&
      String(sheet.getRange(expectedRow, col).getValue()).trim() === String(jobId).trim()) {
    return expectedRow;
  }
  return findRowByJobId(sheet, jobId, col);
}

/**
 * Sheets turns any text starting with "=" into a live formula, so a site note
 * like "=สรุปงานวันนี้" would be stored as #NAME? and the original wording lost.
 * A leading apostrophe forces plain text; getValue() still returns it without.
 */
function sheetSafeValue(header, value) {
  if (typeof header !== "string" && value === undefined) {
    value = header;
    header = "";
  }
  if (header && DATE_COLUMNS.indexOf(header) !== -1) {
    return formatDateForSheet(value);
  }
  if (typeof value !== "string") return value;
  return value.charAt(0) === "=" ? "'" + value : value;
}

function handleCreate(sheet, task, e) {
  task = task || {};
  // NOT toLocaleString("th-TH") - that emits Buddhist years (2569) which Sheets
  // then stores as a real date in year 2569.
  const now = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm");

  const defaults = {
    Sub_Department: "งานโครงการ",
    Status: "วางแผน / เตรียมอุปกรณ์",
    JSA_Completed: "No",
    Priority: "Medium",
    Site_Photos: "[]",
    Document_Files: "[]",
    Video_Files: "[]",
    Site_Contact_Phone: "",
    Site_Map_Url: "",
    Delivery_Doc: "[]"
  };
  Object.keys(defaults).forEach(function (key) {
    if (!task[key]) task[key] = defaults[key];
  });
  task.Updated_At = now;

  const headers = getHeaders(sheet);

  // Claim the row - and with it the Job_ID - before the slow Drive upload runs,
  // so a second person saving meanwhile cannot be handed the same ID.
  const claim = withSheetLock(function () {
    const existing = readJobIds(sheet, headers);
    let jobId = String(task.Job_ID || "").trim();
    if (!jobId || existing.indexOf(jobId) !== -1) jobId = nextJobId(existing);
    task.Job_ID = jobId;

    const row = headers.map(function (header) {
      // Link columns are written afterwards as rich text
      if (LINK_COLUMNS[header]) return "";
      return task.hasOwnProperty(header) ? sheetSafeValue(header, task[header]) : "";
    });

    sheet.appendRow(row);
    return { jobId: jobId, rowIndex: sheet.getLastRow() };
  });

  const warnings = [];
  const attachments = uploadTaskAttachments(task, warnings);

  withSheetLock(function () {
    const rowIndex = resolveRow(sheet, claim.rowIndex, claim.jobId, headers);
    if (rowIndex > 0) writeAttachmentCells(sheet, headers, rowIndex, attachments);
  });

  invalidateCache();

  return jsonResponse({
    status: "success",
    message: "Task created successfully",
    job_id: claim.jobId,
    rev: getRevision(),
    attachments: attachmentSnapshot(task),
    warnings: warnings
  }, e);
}

function handleUpdate(sheet, task, e) {
  task = task || {};
  const jobId = String(task.Job_ID || "").trim();
  if (!jobId) {
    return jsonResponse({ status: "error", message: "Missing Job_ID" }, e);
  }

  task.Updated_At = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm");

  const warnings = [];
  const attachments = uploadTaskAttachments(task, warnings); // slow - stays outside the lock

  const found = withSheetLock(function () {
    const headers = getHeaders(sheet);
    const jobIdIndex = headers.indexOf("Job_ID");
    if (jobIdIndex === -1) throw new Error("Job_ID column not found.");

    const foundRow = findRowByJobId(sheet, jobId, jobIdIndex + 1);
    if (foundRow === -1) return false;

    // PERFORMANCE: batched setValues instead of ~18 individual setValue calls.
    // Link columns are skipped - writing their plain text back would strip the
    // hyperlinks - and are rewritten as rich text right after.
    const existingRow = sheet.getRange(foundRow, 1, 1, headers.length).getValues()[0];
    let run = [];
    let runStart = -1;

    const flushRun = function () {
      if (runStart !== -1 && run.length > 0) {
        sheet.getRange(foundRow, runStart + 1, 1, run.length).setValues([run]);
      }
      run = [];
      runStart = -1;
    };

    headers.forEach(function (header, colIndex) {
      if (LINK_COLUMNS[header]) {
        flushRun();
        return;
      }
      if (runStart === -1) runStart = colIndex;
      run.push(task.hasOwnProperty(header)
        ? sheetSafeValue(header, task[header])
        : existingRow[colIndex]);
    });
    flushRun();

    writeAttachmentCells(sheet, headers, foundRow, attachments);
    return true;
  });

  if (!found) {
    return jsonResponse({ status: "error", message: "Task not found with Job_ID: " + jobId }, e);
  }

  invalidateCache();

  return jsonResponse({
    status: "success",
    message: "Task updated successfully",
    job_id: jobId,
    rev: getRevision(),
    attachments: attachmentSnapshot(task),
    warnings: warnings
  }, e);
}

function handleDelete(sheet, jobId, e) {
  const deleted = withSheetLock(function () {
    const headers = getHeaders(sheet);
    const jobIdIndex = headers.indexOf("Job_ID");
    if (jobIdIndex === -1) throw new Error("Job_ID column not found.");

    const row = findRowByJobId(sheet, jobId, jobIdIndex + 1);
    if (row === -1) return false;
    sheet.deleteRow(row);
    return true;
  });

  if (!deleted) {
    return jsonResponse({ status: "error", message: "Task not found to delete." }, e);
  }

  invalidateCache();
  return jsonResponse({ status: "success", message: "Task deleted successfully", rev: getRevision() }, e);
}

/* ---------------------------------------------------------------------------
 * CLICKABLE ATTACHMENT CELLS
 * ------------------------------------------------------------------------- */

/**
 * Writes all three attachment columns as rich text, one clickable line per file.
 */
function writeAttachmentCells(sheet, headers, rowIndex, attachments) {
  // Drive unavailable: leave whatever is already in the sheet alone.
  if (attachments.driveBlocked) return;

  // SPEED: a save that touches no attachment field should cost zero Sheets calls
  // here. It used to still read and compare the row height every time.
  const names = LINK_COLUMN_NAMES.filter(function (name) {
    return attachments.entries[name];
  });
  if (names.length === 0) return;

  let maxLines = 1;
  names.forEach(function (name) {
    const entries = attachments.entries[name];
    maxLines = Math.max(maxLines, entries.length);
    writeLinkedListCell(sheet, headers, rowIndex, name, entries);
  });

  if (maxLines <= 1) return;   // single-line rows already fit the default height

  const wanted = Math.min(maxLines * MIN_ROW_HEIGHT + 8, MAX_ROW_HEIGHT);
  if (sheet.getRowHeight(rowIndex) < wanted) {
    sheet.setRowHeight(rowIndex, wanted);
  }
}

/**
 * Puts several real hyperlinks into ONE cell, one per line.
 *
 * =HYPERLINK() only supports a single link per cell, which is why the files
 * could not be listed individually before. A RichTextValue carries many link
 * ranges, so every line renders as its own blue clickable link.
 */
function writeLinkedListCell(sheet, headers, rowIndex, header, entries) {
  const col = headers.indexOf(header) + 1;
  if (col <= 0) return;

  const cell = sheet.getRange(rowIndex, col);
  if (!entries || entries.length === 0) {
    cell.clearContent();
    return;
  }

  let text = "";
  const ranges = [];
  entries.forEach(function (entry, i) {
    if (i > 0) text += "\n";
    const start = text.length;
    text += entry.label;
    if (entry.url) ranges.push({ start: start, end: text.length, url: entry.url });
  });

  let builder = SpreadsheetApp.newRichTextValue().setText(text);
  ranges.forEach(function (r) {
    builder = builder.setLinkUrl(r.start, r.end, r.url);
  });

  // Wrapping and alignment are applied to the whole column by ensureHeaders, so
  // repeating them per cell was two extra Sheets calls per attachment column.
  cell.setRichTextValue(builder.build());
}

/** Photos: the app only needs the URL of each image. */
function extractLinkedUrls(richTextValue, fallbackValue) {
  const urls = [];
  if (richTextValue) {
    richTextValue.getRuns().forEach(function (run) {
      const url = run.getLinkUrl();
      if (url && urls.indexOf(url) === -1) urls.push(url);
    });
    if (urls.length > 0) return urls;
  }
  // Rows written by an older version still hold a raw JSON array
  return parseJsonArray(fallbackValue).filter(function (u) {
    return typeof u === "string" && u;
  });
}

/** Documents & videos: rebuild {name, size, url} from the label and the link. */
function extractLinkedFiles(richTextValue, fallbackValue, config) {
  if (richTextValue) {
    const out = [];
    richTextValue.getRuns().forEach(function (run) {
      const url = run.getLinkUrl();
      if (!url) return;

      let label = String(run.getText()).trim();
      if (config && config.icon && label.indexOf(config.icon) === 0) {
        label = label.substring(config.icon.length).trim();
      }

      let name = label;
      let size = "";
      const sep = label.lastIndexOf(SIZE_SEPARATOR);
      if (sep !== -1) {
        name = label.substring(0, sep);
        size = label.substring(sep + SIZE_SEPARATOR.length);
      }

      const id = driveIdFromUrl(url);
      const item = { name: name, size: size, url: url, fileId: id };
      if (config && config.kind === "file" && config.icon === "🎬" && id) {
        item.thumb = driveImageUrl(id);
      }
      out.push(item);
    });
    if (out.length > 0) return out;
  }
  return parseJsonArray(fallbackValue);
}

/* ---------------------------------------------------------------------------
 * DRIVE UPLOAD PIPELINE
 * ------------------------------------------------------------------------- */

/**
 * Replaces every Base64 data URL inside the task with a Google Drive URL and
 * builds the display entries for the sheet.
 *
 * Items that are already URLs (previously uploaded) are left untouched, so the
 * same task can be saved repeatedly without duplicating files.
 */
function uploadTaskAttachments(task, warnings) {
  const jobId = String(task.Job_ID || "UNSORTED").trim() || "UNSORTED";
  const stamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd-HHmmss");
  const result = { entries: {}, driveBlocked: false };

  // SPEED: a text-only edit does not need Drive at all, but this probe used to
  // run on every single save - one wasted Drive round-trip per keystroke-sized
  // change. Only pay for it when there is really something to upload.
  const driveError = hasPendingUploads(task) ? checkDriveAccess() : "";
  if (driveError) {
    result.driveBlocked = true;
    warnings.push(driveError);
    // Leave every attachment field untouched so nothing already saved is lost.
    LINK_COLUMN_NAMES.forEach(function (name) { delete task[name]; });
    return result;
  }

  // ---- Site photos: array of URL strings -------------------------------
  if (task.hasOwnProperty("Site_Photos")) {
    const photos = parseJsonArray(task.Site_Photos);
    const outPhotos = [];
    const entries = [];
    let folder = null;
    let failed = 0;

    photos.forEach(function (entry, i) {
      const src = typeof entry === "string" ? entry : (entry && (entry.url || entry.dataUrl)) || "";
      if (!src) return;

      if (src.indexOf("data:") !== 0) {
        outPhotos.push(src);
        return;
      }

      try {
        if (!folder) folder = getJobFolder(jobId, "Photos");
        const file = uploadDataUrl(folder, src, jobId + "_photo_" + stamp + "_" + (i + 1), warnings);
        outPhotos.push(driveViewUrl(file.getId()));
      } catch (err) {
        failed++;
        warnings.push("อัปโหลดรูปภาพลำดับที่ " + (i + 1) + " ไม่สำเร็จ: " + err);
      }
    });

    if (failed > 0) {
      // Keep the previously saved list rather than writing a partial one
      delete task.Site_Photos;
    } else {
      outPhotos.forEach(function (url, i) {
        entries.push({ label: "📷 รูปที่ " + (i + 1), url: viewUrlFor(url) });
      });
      task.Site_Photos = JSON.stringify(outPhotos);
      result.entries.Site_Photos = entries;
    }
  }

  // ---- Documents, videos & delivery docs: array of {name,size,type,dataUrl} ------------
  [
    { key: "Document_Files", folder: "Documents", icon: "📄" },
    { key: "Video_Files", folder: "Videos", icon: "🎬", isVideo: true },
    { key: "Delivery_Doc", folder: "Delivery_Docs", icon: "📝" }
  ].forEach(function (group) {
    if (!task.hasOwnProperty(group.key)) return;

    const items = parseJsonArray(task[group.key]);
    const out = [];
    let folder = null;
    let failed = 0;

    items.forEach(function (item, i) {
      if (!item) return;

      if (typeof item === "string") {
        out.push({ name: "ไฟล์ " + (i + 1), url: item, fileId: driveIdFromUrl(item) });
        return;
      }

      // Already uploaded previously
      if (item.url && String(item.url).indexOf("data:") !== 0) {
        const keptId = item.fileId || driveIdFromUrl(item.url);
        out.push({
          name: item.name,
          size: item.size,
          type: item.type,
          url: item.url,
          fileId: keptId,
          thumb: item.thumb || (group.isVideo && keptId ? driveImageUrl(keptId) : "")
        });
        return;
      }

      if (!item.dataUrl || String(item.dataUrl).indexOf("data:") !== 0) return;

      try {
        if (!folder) folder = getJobFolder(jobId, group.folder);
        const safeName = sanitizeFileName(item.name || (jobId + "_" + group.folder + "_" + (i + 1)));
        const file = uploadDataUrl(folder, item.dataUrl, safeName, warnings);
        out.push({
          name: item.name || safeName,
          size: item.size || "",
          type: item.type || file.getMimeType(),
          url: file.getUrl(),
          fileId: file.getId(),
          // Drive renders a poster frame for videos at the thumbnail endpoint
          thumb: group.isVideo ? driveImageUrl(file.getId()) : ""
        });
      } catch (err) {
        failed++;
        warnings.push("อัปโหลดไฟล์ " + (item.name || i + 1) + " ไม่สำเร็จ: " + err);
      }
    });

    // Never overwrite the cell with a partial list - that would delete files.
    if (failed > 0) {
      delete task[group.key];
      return;
    }

    task[group.key] = JSON.stringify(out);
    result.entries[group.key] = out.map(function (f, i) {
      const name = f.name || "ไฟล์ " + (i + 1);
      return {
        label: group.icon + " " + name + (f.size ? SIZE_SEPARATOR + f.size : ""),
        url: f.url
      };
    });
  });

  return result;
}

/** True only if some attachment field still carries a Base64 payload to upload. */
function hasPendingUploads(task) {
  let found = false;
  LINK_COLUMN_NAMES.forEach(function (name) {
    if (found || !task.hasOwnProperty(name)) return;
    parseJsonArray(task[name]).forEach(function (entry) {
      if (found) return;
      const src = typeof entry === "string" ? entry : (entry && (entry.dataUrl || entry.url)) || "";
      if (String(src).indexOf("data:") === 0) found = true;
    });
  });
  return found;
}

/**
 * Returns "" when Drive is usable, otherwise a single human-readable instruction.
 * Called once per save so the user gets one clear message instead of one per file.
 */
function checkDriveAccess() {
  try {
    DriveApp.getRootFolder().getId();
    return "";
  } catch (err) {
    return "ยังไม่ได้อนุญาตสิทธิ์ Google Drive ให้สคริปต์ จึงยังอัปโหลดไฟล์แนบไม่ได้ " +
      "(ข้อมูลอื่นบันทึกแล้ว และไฟล์แนบเดิมไม่ถูกลบ)\n" +
      "วิธีแก้: เปิด Apps Script > เลือกฟังก์ชัน authorizeDriveAccess > กด Run > กด Allow " +
      "แล้วกลับมา Deploy > Manage deployments > แก้ไข > Version: New version > Deploy";
  }
}

function uploadDataUrl(folder, dataUrl, baseName, warnings) {
  const commaAt = dataUrl.indexOf(",");
  if (commaAt === -1) throw new Error("Invalid data URL");

  const meta = dataUrl.substring(5, commaAt);          // e.g. "image/jpeg;base64"
  const payload = dataUrl.substring(commaAt + 1);
  const mimeType = meta.split(";")[0] || "application/octet-stream";

  const bytes = Utilities.base64Decode(payload);
  let fileName = baseName;
  if (fileName.indexOf(".") === -1) {
    fileName += "." + extensionForMime(mimeType);
  }

  const file = folder.createFile(Utilities.newBlob(bytes, mimeType, fileName));

  // Normally skipped: the parent folder is already shared, and the file inherits
  // that. Only fall back to per-file sharing when the folder could not be shared.
  if (!rootIsShared()) {
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {
      warnings.push("ไม่สามารถตั้งค่าแชร์ไฟล์ " + fileName + " เป็นสาธารณะได้ (นโยบายองค์กร) - รูปอาจไม่แสดงในแอป");
    }
  }

  return file;
}

/** Opens the file in Drive - used for the clickable links. */
function driveViewUrl(fileId) {
  return "https://drive.google.com/file/d/" + fileId + "/view";
}

/** Renders inside <img> tags - used by the web app for thumbnails. */
function driveImageUrl(fileId) {
  return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1600";
}

function viewUrlFor(url) {
  const id = driveIdFromUrl(url);
  return id ? driveViewUrl(id) : url;
}

function driveIdFromUrl(url) {
  if (!url) return "";
  let m = /\/file\/d\/([a-zA-Z0-9_-]+)/.exec(url);
  if (m) return m[1];
  m = /[?&]id=([a-zA-Z0-9_-]+)/.exec(url);
  return m ? m[1] : "";
}

/**
 * SPEED: folder lookups used to be a name search on every call - three searches
 * of the whole Drive root per save, plus one per job folder. The ids are stable,
 * so remember them and go straight to getFolderById.
 */
function getRootFolder() {
  const props = PropertiesService.getScriptProperties();
  const cachedId = props.getProperty(ROOT_FOLDER_ID_KEY);
  if (cachedId) {
    try {
      return DriveApp.getFolderById(cachedId);
    } catch (err) {
      // Folder was deleted or moved to trash - fall through and find/create it
    }
  }

  const it = DriveApp.getFoldersByName(DRIVE_ROOT_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(DRIVE_ROOT_FOLDER_NAME);

  // Share the CONTAINER once instead of every file. Files inherit link access
  // from their parent, and setSharing() was by far the most expensive part of a
  // save: one extra Drive round-trip per attachment, every single time.
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props.setProperty(ROOT_SHARED_KEY, "1");
  } catch (err) {
    // Workspace policy may forbid link sharing - uploadDataUrl then falls back
    // to sharing each file individually and warns if that fails too.
    props.deleteProperty(ROOT_SHARED_KEY);
  }

  props.setProperty(ROOT_FOLDER_ID_KEY, folder.getId());
  return folder;
}

function rootIsShared() {
  try {
    return PropertiesService.getScriptProperties().getProperty(ROOT_SHARED_KEY) === "1";
  } catch (err) {
    return false;
  }
}

function getSubFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getJobFolder(jobId, category) {
  const cache = CacheService.getScriptCache();
  const key = "pts_folder_" + jobId + "_" + category;
  const cachedId = cache.get(key);
  if (cachedId) {
    try {
      return DriveApp.getFolderById(cachedId);
    } catch (err) {
      // Recreate below
    }
  }

  const folder = getSubFolder(getSubFolder(getRootFolder(), jobId), category);
  cache.put(key, folder.getId(), FOLDER_CACHE_SECONDS);
  return folder;
}

/* ---------------------------------------------------------------------------
 * HELPERS
 * ------------------------------------------------------------------------- */

function parseJsonArray(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value;
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function sanitizeFileName(name) {
  return String(name).replace(/[\\\/:*?"<>|]/g, "_").substring(0, 120);
}

function extensionForMime(mimeType) {
  const map = {
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
    "image/png": "png",
    "image/webp": "webp",
    "image/gif": "gif",
    "application/pdf": "pdf",
    "video/mp4": "mp4",
    "video/webm": "webm",
    "video/quicktime": "mov"
  };
  if (map[mimeType]) return map[mimeType];
  const tail = mimeType.split("/")[1];
  return tail ? tail.split("+")[0] : "bin";
}

function jsonResponse(obj, e) {
  return rawJsonResponse(JSON.stringify(obj), e);
}

/** Same as jsonResponse but skips re-serialising an already-built JSON string. */
function rawJsonResponse(payload, e) {
  const callback = e && e.parameter ? e.parameter.callback : null;
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + payload + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------------------------------------------------------------------
 * ONE-OFF MAINTENANCE (run manually from the Apps Script editor)
 * ------------------------------------------------------------------------- */

/**
 * RUN THIS FIRST after pasting the script.
 *
 * Apps Script only asks for the Google Drive permission the first time code that
 * actually touches Drive is executed. Running this triggers that consent screen,
 * which is what fixes the "คุณไม่ได้รับอนุญาตให้เรียกใช้ DriveApp..." error.
 *
 * Afterwards redeploy: Deploy > Manage deployments > ✏️ > Version: "New version".
 */
function authorizeDriveAccess() {
  const folder = getRootFolder();
  const sheet = getOrCreateSheet();
  const msg = "OK - Drive พร้อมใช้งานแล้ว\n" +
    "โฟลเดอร์เก็บไฟล์: " + folder.getUrl() + "\n" +
    "ชีต: " + sheet.getName() + "\n\n" +
    "ขั้นตอนต่อไป: Deploy > Manage deployments > ✏️ > Version: New version > Deploy";
  Logger.log(msg);
  return msg;
}

/**
 * Converts every existing row to the clickable-link format, uploads any leftover
 * Base64 attachments to Drive, and removes the old preview columns.
 * Files already on Drive are skipped, so this is safe to run more than once.
 */
function rebuildAllRows() {
  const sheet = getOrCreateSheet();
  const headers = getHeaders(sheet);
  const lastRow = sheet.getLastRow();
  const warnings = [];

  const linkCols = {};
  LINK_COLUMN_NAMES.forEach(function (name) {
    const idx = headers.indexOf(name);
    if (idx !== -1) linkCols[idx] = name;
  });

  for (let row = 2; row <= lastRow; row++) {
    const range = sheet.getRange(row, 1, 1, headers.length);
    const values = range.getValues()[0];
    const rich = range.getRichTextValues()[0];

    const task = {};
    headers.forEach(function (h, i) {
      if (!h) return;
      if (linkCols[i]) {
        const config = LINK_COLUMNS[linkCols[i]];
        task[h] = JSON.stringify(
          config.kind === "photo"
            ? extractLinkedUrls(rich[i], values[i])
            : extractLinkedFiles(rich[i], values[i], config)
        );
        return;
      }
      task[h] = values[i];
    });
    if (!task.Job_ID) continue;

    const attachments = uploadTaskAttachments(task, warnings);
    writeAttachmentCells(sheet, headers, row, attachments);
  }

  invalidateCache();
  const msg = "ซ่อมข้อมูลเรียบร้อย " + (lastRow - 1) + " แถว\n" +
    (warnings.length ? "คำเตือน:\n- " + warnings.join("\n- ") : "ไม่มีข้อผิดพลาด");
  Logger.log(msg);
  return msg;
}

/**
 * ONE-OFF REPAIR for sheets written by v6 or earlier.
 *
 * The old ID formula could hand the same Job_ID to many rows. While duplicates
 * exist the app is effectively read-only for all but the first copy: UPDATE and
 * DELETE stop at the first match, so edits to any later row silently land on the
 * first one instead.
 *
 * This keeps the first row of each ID untouched and renumbers the rest from the
 * highest number in use. Nothing is deleted. Safe to run more than once - it
 * reports "ไม่พบรหัสซ้ำ" when there is nothing to do.
 */
function repairDuplicateJobIds() {
  const sheet = getOrCreateSheet();
  const headers = getHeaders(sheet);
  const col = headers.indexOf("Job_ID") + 1;
  if (col <= 0) return "ไม่พบคอลัมน์ Job_ID";

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "ยังไม่มีข้อมูลในชีต";

  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  const ids = values.map(function (r) { return String(r[0]).trim(); });

  let max = 0;
  ids.forEach(function (id) {
    const m = /(\d+)\s*$/.exec(id);
    if (m) max = Math.max(max, Number(m[1]));
  });

  const seen = {};
  const changes = [];
  const out = ids.map(function (id) {
    if (!id) return [id];
    if (!seen[id]) {
      seen[id] = true;
      return [id];
    }
    max++;
    const fresh = "PTS-ENG-" + String(max).padStart(3, "0");
    changes.push(id + " -> " + fresh);
    return [fresh];
  });

  if (changes.length === 0) {
    const clean = "ไม่พบรหัสซ้ำ (" + ids.length + " แถว)";
    Logger.log(clean);
    return clean;
  }

  sheet.getRange(2, col, out.length, 1).setValues(out);
  invalidateCache();

  const msg = "แก้รหัสซ้ำแล้ว " + changes.length + " แถว จากทั้งหมด " + ids.length + " แถว\n- " +
    changes.join("\n- ");
  Logger.log(msg);
  return msg;
}

/**
 * รันฟังก์ชันนี้ใน Apps Script เพื่อสร้าง/อัปเดตหัวคอลัมน์ใน Google Sheet ทันที
 */
function setupSheetHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  CacheService.getScriptCache().remove(HEADER_OK_KEY);
  const headers = ensureHeaders(sheet);
  invalidateCache();
  SpreadsheetApp.flush();
  Logger.log("✅ อัปเดตหัวตารางเรียบร้อย: " + headers.join(", "));
  return "✅ อัปเดตหัวตารางเรียบร้อย (" + headers.length + " คอลัมน์)";
}

/** Optional: hides nothing, just re-applies column widths and wrapping. */
function tidySheetLayout() {
  const sheet = getOrCreateSheet();
  CacheService.getScriptCache().remove(HEADER_OK_KEY);
  ensureHeaders(sheet);
  invalidateCache();
  return "จัดหน้าตาชีตเรียบร้อย";
}
