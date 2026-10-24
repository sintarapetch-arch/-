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
 * v9: To-do list, built to match the sales app (PTS Sales).
 *     A new tab "Tasks" holds tasks with an assignee, a due date and an
 *     optional time, priority and type - the same column names and the same
 *     Open/Done status the sales app uses, so the two are directly
 *     comparable. Job_ID takes the place of its Opportunity_ID and is
 *     optional. Read with ?tasks=1, written with TASK_SAVE / TASK_DONE /
 *     TASK_DELETE. sendTaskReminderEmail() is an optional daily digest that
 *     stays dormant until someone sets an address and adds a trigger.
 *     The job sheet and the attachment flow are untouched.
 *
 * v8: Option lists managed from inside the app.
 *     Sub-departments, technician names and customer names live in a new tab
 *     "App_Settings" (Key | Value, one row per entry), read with ?settings=1
 *     and written with action SAVE_SETTINGS. The tab is created on the first
 *     save. Nothing about the task sheet or the attachment flow changed.
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
        if (formatted) {
          values[r][c] = "'" + formatted;
          count++;
        }
      }
    });
  }

  range.setValues(values);

  // จัดข้อความให้อยู่กึ่งกลาง
  DATE_COLUMNS.forEach(function (name) {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;
    try {
      sheet.getRange(2, col, Math.max(lastRow - 1, 1))
        .setHorizontalAlignment("center");
    } catch (err) {}
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
    try {
      sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1))
        .setHorizontalAlignment("center");
    } catch (err) {}
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

    // User-managed option lists (sub-departments, technicians, customers)
    if (e && e.parameter && e.parameter.settings === "1") {
      return jsonResponse({ status: "success", settings: readSettings() }, e);
    }

    // To-do list: who has to do what, by when
    if (e && e.parameter && e.parameter.tasks === "1") {
      return jsonResponse({ status: "success", tasks: readJobTasks(), rev: getRevision() }, e);
    }

    // Web Push: is it set up, and is this device registered
    if (e && e.parameter && e.parameter.push === "1") {
      return jsonResponse({ status: "success", push: pushStatusPayload(e.parameter.endpoint || "") }, e);
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
    } else if (action === "SAVE_SETTINGS") {
      return handleSaveSettings(requestData.settings, e);
    } else if (action === "TASK_SAVE") {
      return handleSaveJobTask(requestData.task, e);
    } else if (action === "TASK_DONE") {
      return handleSetJobTaskDone(requestData, e);
    } else if (action === "TASK_DELETE") {
      return handleDeleteJobTask(requestData.task_id, e);
    } else if (action === "PUSH_SUB_SAVE") {
      return handleSavePushSub(requestData, e);
    } else if (action === "PUSH_SUB_DELETE") {
      return handleDeletePushSub(requestData, e);
    } else if (action === "PUSH_TEST") {
      return handleTestPush(requestData, e);
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
    const d = formatDateForSheet(value);
    return d ? "'" + d : "";
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

/* ---------------------------------------------------------------------------
 * APP SETTINGS - option lists the team manages from inside the app
 *
 * Sub-departments, technician names and customer names used to be hard-coded
 * in index.html, so adding one meant a code change. They now live in their
 * own tab, "App_Settings", one row per entry (Key | Value). The tab is created
 * on first save; the task sheet is never touched by any of this.
 * ------------------------------------------------------------------------- */
const SETTINGS_SHEET_NAME = "App_Settings";
const SETTINGS_KEYS = ["sub_departments", "technicians", "customers", "task_types"];
const SETTINGS_MAX_ITEMS = 300;
const SETTINGS_MAX_LEN = 120;

function getSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([["Key", "Value"]])
      .setFontWeight("bold").setBackground("#991B1B").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 170);
    sheet.setColumnWidth(2, 340);
  }
  return sheet;
}

function readSettings() {
  const out = {};
  SETTINGS_KEYS.forEach(function (k) { out[k] = []; });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return out;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  rows.forEach(function (r) {
    const key = String(r[0] || "").trim();
    const val = String(r[1] || "").trim();
    if (!key || !val || !out.hasOwnProperty(key)) return;
    if (out[key].indexOf(val) === -1) out[key].push(val);
  });
  return out;
}

function cleanSettingsList(list) {
  const seen = {};
  const out = [];
  (Array.isArray(list) ? list : []).forEach(function (v) {
    const s = String(v || "").replace(/\s+/g, " ").trim().slice(0, SETTINGS_MAX_LEN);
    if (!s || seen[s]) return;
    seen[s] = true;
    if (out.length < SETTINGS_MAX_ITEMS) out.push(s);
  });
  return out;
}

function handleSaveSettings(settings, e) {
  const incoming = settings || {};
  const keys = SETTINGS_KEYS.filter(function (k) { return incoming.hasOwnProperty(k); });
  if (keys.length === 0) {
    return jsonResponse({ status: "error", message: "No settings given" }, e);
  }

  withSheetLock(function () {
    const sheet = getSettingsSheet();
    const current = readSettings();
    keys.forEach(function (k) { current[k] = cleanSettingsList(incoming[k]); });

    // The tab is tiny, so it is simply rewritten - rows stay grouped by key
    const rows = [];
    SETTINGS_KEYS.forEach(function (k) {
      current[k].forEach(function (v) { rows.push([k, sheetSafeValue("", v)]); });
    });
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  });

  return jsonResponse({ status: "success", settings: readSettings() }, e);
}

/* ---------------------------------------------------------------------------
 * TASKS - the to-do list
 *
 * Deliberately built to match the sales app (PTS Sales), because the same
 * system was asked for in both: identical column names, Open/Done status,
 * yyyy-MM-dd dates with an optional HH:mm time, Priority and Type. Anyone
 * who knows one app's task list already knows this one, and the two sheets
 * can be compared side by side.
 *
 * Job_ID takes the place of the sales app's Opportunity_ID. It is optional:
 * filled in, the task belongs to that job; blank, it is a standalone reminder.
 *
 * Whether a task is due or overdue is worked out in the browser from
 * Due_Date + Due_Time. The server only stores and returns the rows, so no
 * timezone logic is duplicated across the two sides.
 *
 * The job sheet is never touched by any of this - deleting a job leaves its
 * tasks behind on purpose, so nothing disappears silently.
 * ------------------------------------------------------------------------- */
const TASKS_SHEET_NAME = "Tasks";
const TASK_HEADERS = [
  "Task_ID", "Title", "Detail", "Assignee", "Due_Date", "Due_Time", "Job_ID",
  "Status", "Created_By", "Created_At", "Updated_At", "Done_By", "Done_At",
  "Push_Stage", "Priority", "Type"
];
const TASK_OPEN = "Open";
const TASK_DONE = "Done";
const TASK_PRIORITIES = ["Low", "Normal", "High", "Urgent"];
const TASK_TITLE_MAX = 200;
const TASK_DETAIL_MAX = 1000;
// A finished task stays in the sheet for good; it just stops being sent to the
// browser once it is this old, so an old list never slows the app down.
const TASK_DONE_KEEP_DAYS = 60;

function getTasksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TASKS_SHEET_NAME);
  if (sheet) {
    // A tab from an older version: append whatever columns were added later
    const have = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(String);
    const missing = TASK_HEADERS.filter(function (h) { return have.indexOf(h) === -1; });
    if (missing.length > 0) {
      sheet.getRange(1, sheet.getLastColumn() + 1, 1, missing.length).setValues([missing]);
    }
    return sheet;
  }
  sheet = ss.insertSheet(TASKS_SHEET_NAME);
  sheet.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS])
    .setFontWeight("bold").setBackground("#991B1B").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(TASK_HEADERS.indexOf("Title") + 1, 300);
  sheet.setColumnWidth(TASK_HEADERS.indexOf("Detail") + 1, 240);
  sheet.setColumnWidth(TASK_HEADERS.indexOf("Assignee") + 1, 150);
  // Time as text, or the sheet turns 09:00 into a 1899 date that drifts by minutes
  sheet.getRange(2, TASK_HEADERS.indexOf("Due_Time") + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat("@");
  return sheet;
}

function taskHeaderIndexes(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  const idx = {};
  TASK_HEADERS.forEach(function (h) { idx[h] = headers.indexOf(h); });
  return idx;
}

/** yyyy-MM-dd from a Date or from text already in that shape */
function taskDateOnly(value) {
  if (!value) return "";
  if (value instanceof Date) return Utilities.formatDate(value, "Asia/Bangkok", "yyyy-MM-dd");
  const text = String(value).replace(/^'/, "").trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  if (m) return m[1] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[3]).slice(-2);
  // dd/MM/yyyy as typed straight into the sheet by hand
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (m) return m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  return "";
}

/** HH:mm from text, a Date, or an ISO string - "" when there is no time */
function taskTime(value) {
  if (!value) return "";
  if (value instanceof Date) return Utilities.formatDate(value, "Asia/Bangkok", "HH:mm");
  const text = String(value).replace(/^'/, "").trim();
  const iso = /T(\d{2}):(\d{2})/.exec(text);
  if (iso) return iso[1] + ":" + iso[2];
  const plain = /^(\d{1,2}):(\d{2})/.exec(text);
  if (!plain) return "";
  const h = Number(plain[1]), mi = Number(plain[2]);
  if (h > 23 || mi > 59) return "";
  return ("0" + h).slice(-2) + ":" + plain[2];
}

function taskTimestamp(value) {
  if (!value) return "";
  if (value instanceof Date) return Utilities.formatDate(value, "Asia/Bangkok", "dd/MM/yyyy HH:mm");
  return String(value).replace(/^'/, "").trim();
}

function cleanTaskText(value, limit) {
  return String(value === undefined || value === null ? "" : value)
    .replace(/^'/, "").replace(/\s+/g, " ").trim().slice(0, limit || TASK_DETAIL_MAX);
}

/** Row -> the shape the browser works with */
function rowToTask(row, idx) {
  const get = function (h) {
    const col = idx[h];
    return col === -1 || col === undefined ? "" : row[col];
  };
  return {
    Task_ID: cleanTaskText(get("Task_ID"), 40),
    Title: cleanTaskText(get("Title"), TASK_TITLE_MAX),
    Detail: String(get("Detail") === undefined || get("Detail") === null ? "" : get("Detail")).replace(/^'/, "").trim().slice(0, TASK_DETAIL_MAX),
    Assignee: cleanTaskText(get("Assignee"), 120),
    Due_Date: taskDateOnly(get("Due_Date")),
    Due_Time: taskTime(get("Due_Time")),
    Job_ID: cleanTaskText(get("Job_ID"), 40),
    Status: cleanTaskText(get("Status"), 20) === TASK_DONE ? TASK_DONE : TASK_OPEN,
    Created_By: cleanTaskText(get("Created_By"), 120),
    Created_At: taskTimestamp(get("Created_At")),
    Updated_At: taskTimestamp(get("Updated_At")),
    Done_By: cleanTaskText(get("Done_By"), 120),
    Done_At: taskTimestamp(get("Done_At")),
    Priority: TASK_PRIORITIES.indexOf(cleanTaskText(get("Priority"), 20)) !== -1 ? cleanTaskText(get("Priority"), 20) : "Normal",
    Type: cleanTaskText(get("Type"), 60)
  };
}

function readAllTaskRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASKS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const idx = taskHeaderIndexes(sheet);
  if (idx.Task_ID === -1) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), 1)).getValues();
  const out = [];
  rows.forEach(function (row) {
    const task = rowToTask(row, idx);
    if (task.Task_ID) out.push(task);
  });
  return out;
}

/** Days from a dd/MM/yyyy HH:mm or yyyy-MM-dd timestamp until now */
function taskDaysSince(stamp) {
  const text = String(stamp || "").trim();
  if (!text) return null;
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(text);
  let d;
  if (m) d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  else {
    m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
    if (!m) return null;
    d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  if (isNaN(d.getTime())) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  d.setHours(0, 0, 0, 0);
  return Math.round((today - d) / 86400000);
}

/** What the browser gets: every open task, plus recently finished ones */
function readJobTasks() {
  return readAllTaskRows().filter(function (t) {
    if (t.Status !== TASK_DONE) return true;
    const age = taskDaysSince(t.Done_At || t.Updated_At);
    return age === null || age <= TASK_DONE_KEEP_DAYS;
  });
}

function nextTaskId(existingIds) {
  let max = 0;
  existingIds.forEach(function (id) {
    const m = /(\d+)\s*$/.exec(String(id).trim());
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return "TSK-" + ("000" + (max + 1)).slice(-4);
}

function handleSaveJobTask(task, e) {
  task = task || {};
  const title = cleanTaskText(task.Title, TASK_TITLE_MAX);
  if (!title) return jsonResponse({ status: "error", message: "กรุณากรอกชื่อ Task" }, e);

  const assignee = cleanTaskText(task.Assignee, 120);
  if (!assignee) return jsonResponse({ status: "error", message: "กรุณาระบุผู้รับผิดชอบ" }, e);

  const due = taskDateOnly(task.Due_Date);
  if (!due) return jsonResponse({ status: "error", message: "กรุณาระบุวันที่ต้องทำ" }, e);

  const priority = TASK_PRIORITIES.indexOf(cleanTaskText(task.Priority, 20)) !== -1
    ? cleanTaskText(task.Priority, 20) : "Normal";
  const now = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm");
  const actor = cleanTaskText(task.Actor, 120);

  const saved = withSheetLock(function () {
    const sheet = getTasksSheet();
    const idx = taskHeaderIndexes(sheet);
    const lastRow = sheet.getLastRow();
    const ids = lastRow > 1
      ? sheet.getRange(2, idx.Task_ID + 1, lastRow - 1, 1).getValues().map(function (r) { return String(r[0] || "").trim(); })
      : [];

    const wanted = cleanTaskText(task.Task_ID, 40);
    const at = wanted ? ids.indexOf(wanted) : -1;
    const rowIndex = at === -1 ? -1 : at + 2;
    const isNew = rowIndex === -1;

    // Editing keeps whatever the row already holds for the done/created fields
    const before = isNew ? null : rowToTask(
      sheet.getRange(rowIndex, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0], idx);

    const record = {
      Task_ID: isNew ? nextTaskId(ids) : wanted,
      Title: title,
      Detail: String(task.Detail || "").replace(/^'/, "").trim().slice(0, TASK_DETAIL_MAX),
      Assignee: assignee,
      Due_Date: due,
      Due_Time: taskTime(task.Due_Time),
      Job_ID: cleanTaskText(task.Job_ID, 40),
      Status: (before ? before.Status : TASK_OPEN),
      Created_By: isNew ? actor : (before.Created_By || actor),
      Created_At: isNew ? now : (before.Created_At || now),
      Updated_At: now,
      Done_By: before ? before.Done_By : "",
      Done_At: before ? before.Done_At : "",
      Push_Stage: "",
      Priority: priority,
      Type: cleanTaskText(task.Type, 60)
    };

    const values = TASK_HEADERS.map(function (h) {
      // Dates and times keep a leading apostrophe so Sheets stores them as
      // written instead of reinterpreting them
      if (h === "Due_Date" || h === "Due_Time") return record[h] ? "'" + record[h] : "";
      return sheetSafeValue("", record[h]);
    });

    if (isNew) sheet.appendRow(values);
    else sheet.getRange(rowIndex, 1, 1, TASK_HEADERS.length).setValues([values]);

    delete record.Push_Stage;
    return record;
  });

  invalidateCache();
  // Hand a newly assigned task straight to the person's devices; everything
  // else is left to the one-minute timer.
  if (typeof afterTaskWrite_ === "function") {
    const reassigned = actor && saved.Assignee !== actor;
    afterTaskWrite_(saved, reassigned ? actor : "");
  }
  return jsonResponse({ status: "success", task: saved, rev: getRevision() }, e);
}

/** Tick a task off, or put it back on the list */
function handleSetJobTaskDone(input, e) {
  input = input || {};
  const id = cleanTaskText(input.Task_ID, 40);
  if (!id) return jsonResponse({ status: "error", message: "Missing Task_ID" }, e);

  const done = input.Done === true || String(input.Done) === "true";
  const actor = cleanTaskText(input.Actor, 120);
  const now = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm");

  const saved = withSheetLock(function () {
    const sheet = getTasksSheet();
    const idx = taskHeaderIndexes(sheet);
    if (sheet.getLastRow() < 2) return null;

    const ids = sheet.getRange(2, idx.Task_ID + 1, sheet.getLastRow() - 1, 1).getValues()
      .map(function (r) { return String(r[0] || "").trim(); });
    const at = ids.indexOf(id);
    if (at === -1) return null;
    const rowIndex = at + 2;

    const current = rowToTask(sheet.getRange(rowIndex, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0], idx);
    current.Status = done ? TASK_DONE : TASK_OPEN;
    current.Done_By = done ? actor : "";
    current.Done_At = done ? now : "";
    current.Updated_At = now;

    const setCell = function (header, value) {
      const col = idx[header];
      if (col === -1 || col === undefined) return;
      sheet.getRange(rowIndex, col + 1).setValue(sheetSafeValue("", value));
    };
    setCell("Status", current.Status);
    setCell("Done_By", current.Done_By);
    setCell("Done_At", current.Done_At);
    setCell("Updated_At", current.Updated_At);
    setCell("Push_Stage", "");
    return current;
  });

  if (!saved) return jsonResponse({ status: "error", message: "ไม่พบ Task นี้ อาจถูกลบไปแล้ว" }, e);

  invalidateCache();
  if (typeof afterTaskWrite_ === "function") afterTaskWrite_(null, "");
  return jsonResponse({ status: "success", task: saved, rev: getRevision() }, e);
}

function handleDeleteJobTask(taskId, e) {
  const id = cleanTaskText(taskId, 40);
  if (!id) return jsonResponse({ status: "error", message: "Missing Task_ID" }, e);

  const removed = withSheetLock(function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASKS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return false;
    const idx = taskHeaderIndexes(sheet);
    if (idx.Task_ID === -1) return false;
    const ids = sheet.getRange(2, idx.Task_ID + 1, sheet.getLastRow() - 1, 1).getValues()
      .map(function (r) { return String(r[0] || "").trim(); });
    const at = ids.indexOf(id);
    if (at === -1) return false;
    sheet.deleteRow(at + 2);
    return true;
  });

  if (!removed) return jsonResponse({ status: "error", message: "ไม่พบ Task นี้ อาจถูกลบไปแล้ว" }, e);

  invalidateCache();
  return jsonResponse({ status: "success", task_id: id, rev: getRevision() }, e);
}

/**
 * OPTIONAL daily email digest - dormant until someone fills in an address.
 * To switch on: Apps Script > Triggers > Add Trigger > sendTaskReminderEmail,
 * Time-driven, Day timer, 7am-8am.
 */
const TASK_REMINDER_EMAIL = "";   // e.g. "pakorn@example.com"

function sendTaskReminderEmail() {
  if (!TASK_REMINDER_EMAIL) return "ยังไม่ได้ตั้งค่า TASK_REMINDER_EMAIL";

  const today = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
  const late = [];
  const due = [];
  readJobTasks().forEach(function (t) {
    if (t.Status === TASK_DONE || !t.Due_Date) return;
    if (t.Due_Date < today) late.push(t);
    else if (t.Due_Date === today) due.push(t);
  });
  if (late.length === 0 && due.length === 0) return "ไม่มีงานที่ต้องแจ้งเตือนวันนี้";

  const line = function (t) {
    return "• " + t.Title + " — " + t.Assignee +
      " (กำหนด " + t.Due_Date.split("-").reverse().join("/") + (t.Due_Time ? " " + t.Due_Time + " น." : "") +
      (t.Job_ID ? ", ใบงาน " + t.Job_ID : "") + ")";
  };
  const body = [
    late.length ? "⛔ เกินกำหนด " + late.length + " รายการ:\n" + late.map(line).join("\n") : "",
    due.length ? "⏰ ครบกำหนดวันนี้ " + due.length + " รายการ:\n" + due.map(line).join("\n") : ""
  ].filter(Boolean).join("\n\n");

  MailApp.sendEmail({
    to: TASK_REMINDER_EMAIL,
    subject: "[PTS] งานที่ต้องทำ — เกินกำหนด " + late.length + " / ครบกำหนดวันนี้ " + due.length,
    body: body + "\n\nเปิดระบบ: https://sintarapetch-arch.github.io/-/"
  });
  return "ส่งอีเมลแจ้งเตือนแล้ว";
}

/* ---------------------------------------------------------------------------
 * WEB PUSH - แจ้งเตือน Task ถึงเครื่อง แม้ปิดแอปอยู่
 *
 * ยกมาจากแอพงานขาย (PTS Sales) ให้ทำงานเหมือนกันทั้งสองแอพ
 *   1. หน้าเว็บขออนุญาตแจ้งเตือน แล้วลงทะเบียนเครื่อง (endpoint) เก็บในแท็บ Push_Subs
 *   2. ตัวตั้งเวลาเรียก pushDueTasks() ทุก 1 นาที ตรวจ Task ที่ครบกำหนด / เกินกำหนด
 *      (ส่วนใหญ่จบในเสี้ยววินาที เพราะจำไว้ว่ารอบหน้าที่มีงานต้องส่งคือกี่โมง)
 *      การมอบหมายงานให้คนอื่น ส่งทันทีตอนบันทึก ไม่ต้องรอตัวตั้งเวลา
 *   3. ส่งข้อความ (เข้ารหัส) ผ่านเซิร์ฟเวอร์แจ้งเตือนของ Google / Apple / Mozilla / Microsoft
 *      เครื่องปลายทางแสดงเองผ่าน sw.js โดยไม่ต้องเปิดแอป
 *
 * Apps Script ไม่มีเครื่องมือเข้ารหัสที่ Web Push บังคับ (ECDSA / ECDH บน P-256,
 * AES-128-GCM) โค้ดส่วนเข้ารหัสด้านล่างจึงเขียนตาม RFC 8291 / 8292 / 6979
 * และยกมาทั้งชุดจากแอพงานขายโดยไม่แก้ เพื่อให้ยังตรงกับที่ทดสอบไว้แล้ว
 *
 * ตั้งค่าครั้งเดียว: เลือกฟังก์ชัน setupPush แล้วกด Run (จะขออนุญาตสิทธิ์เพิ่ม — กดอนุญาต)
 * ------------------------------------------------------------------------- */
const PUSH_SHEET_NAME = "Push_Subs";
const PUSH_DAY_START = "08:00";   // งานไม่ระบุเวลา / เกินกำหนด เตือนช่วงกลางวันเท่านั้น
const PUSH_DAY_END = "21:00";
const PUSH_MAX_FAILS = 10;        // ส่งไม่ผ่านติดกันเกินนี้ = เครื่องเลิกใช้แล้ว ลบทิ้ง
const PUSH_IDLE_MS = 15 * 60000;  // ถึงไม่มีงานรอส่ง ก็ตรวจละเอียดทุก 15 นาที
const PUSH_SUB_HEADERS = ["Sub_ID", "Name", "Endpoint", "P256dh", "Auth", "User_Agent",
  "Created_At", "Updated_At", "Last_Push", "Last_Status", "Fail_Count", "Last_Error"];
const TH_MONTHS_ = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
const PUSH_APP_URL = "https://sintarapetch-arch.github.io/-/";

/* ---------------- ตั้งค่า (รันจาก Apps Script Editor) ---------------- */

/** ตั้งค่าครั้งเดียว — รันซ้ำได้ ไม่ทำให้เครื่องที่ลงทะเบียนไว้หลุด */
function setupPush() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("PTS_VAPID_PRIVATE") || !props.getProperty("PTS_VAPID_PUBLIC")) {
    const k = newKeyPair_();
    props.setProperty("PTS_VAPID_PRIVATE", b64u_(bigToBytes_(k.d, 32)));
    props.setProperty("PTS_VAPID_PUBLIC", b64u_(k.pub));
  }
  pushSubSheet_();
  getTasksSheet();
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === "pushDueTasks"; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("pushDueTasks").timeBased().everyMinutes(1).create();
  pushKick_();
  invalidateCache();
  const msg = "ตั้งค่าแจ้งเตือนแบบปิดแอปเรียบร้อย — ตรวจ Task ทุก 1 นาที";
  Logger.log(msg);
  return msg;
}

/** หยุดส่งแจ้งเตือนแบบปิดแอป — เครื่องที่ลงทะเบียนไว้ยังอยู่ รัน setupPush() เพื่อเปิดใหม่ */
function stopPush() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === "pushDueTasks"; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  return "หยุดแจ้งเตือนแบบปิดแอปแล้ว";
}

function pushKick_() { PropertiesService.getScriptProperties().setProperty("PTS_PUSH_NEXT", "0"); }
function vapidPublic_() { return PropertiesService.getScriptProperties().getProperty("PTS_VAPID_PUBLIC") || ""; }
function vapidPrivate_() { return bytesToBig_(unb64u_(PropertiesService.getScriptProperties().getProperty("PTS_VAPID_PRIVATE") || "")); }
function pushReady_() { return !!(vapidPublic_() && PropertiesService.getScriptProperties().getProperty("PTS_VAPID_PRIVATE")); }
/** Apple บังคับให้ผู้ติดต่อเป็น https:// หรือ mailto: */
function pushSubject_() { return PUSH_APP_URL; }

/* ---------------- แท็บเครื่องที่ลงทะเบียน ---------------- */

function pushSubSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PUSH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PUSH_SHEET_NAME);
    sheet.getRange(1, 1, 1, PUSH_SUB_HEADERS.length).setValues([PUSH_SUB_HEADERS])
      .setFontWeight("bold").setBackground("#991B1B").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readPushSubs_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PUSH_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PUSH_SUB_HEADERS.length).getValues();
  const out = [];
  rows.forEach(function (r) {
    const item = {};
    PUSH_SUB_HEADERS.forEach(function (h, i) { item[h] = r[i] === undefined || r[i] === null ? "" : r[i]; });
    if (String(item.Sub_ID).trim()) out.push(item);
  });
  return out;
}

/**
 * อ่าน-แก้-เขียนทั้งแท็บในล็อกเดียว: mutate ได้รายการทั้งหมด คืนรายการใหม่กลับมา
 * (แท็บนี้เล็กมาก - หนึ่งแถวต่อหนึ่งเครื่อง - จึงเขียนทับทั้งแท็บได้)
 */
function writePushSubs_(mutate) {
  return withSheetLock(function () {
    const sheet = pushSubSheet_();
    const next = mutate(readPushSubs_()) || [];
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, PUSH_SUB_HEADERS.length).clearContent();
    if (next.length > 0) {
      sheet.getRange(2, 1, next.length, PUSH_SUB_HEADERS.length).setValues(
        next.map(function (r) {
          return PUSH_SUB_HEADERS.map(function (h) { return r[h] === undefined || r[h] === null ? "" : r[h]; });
        }));
    }
    return next;
  });
}

/** รหัสเครื่อง = แฮชของ endpoint (endpoint ยาวและเป็นความลับ ไม่ใช้ตรงๆ เป็นกุญแจ) */
function subId_(endpoint) { return b64u_(sha256_(utf8_(String(endpoint)))).slice(0, 22); }

/** หน้าเว็บลงทะเบียนเครื่องนี้ */
function handleSavePushSub(input, e) {
  input = input || {};
  const endpoint = cleanTaskText(input.endpoint, 500);
  const p256dh = cleanTaskText(input.p256dh, 200);
  const auth = cleanTaskText(input.auth, 100);
  const name = cleanTaskText(input.name, 120);
  if (!endpoint || !p256dh || !auth) {
    return jsonResponse({ status: "error", message: "ข้อมูลการลงทะเบียนไม่ครบ" }, e);
  }

  const id = subId_(endpoint);
  const now = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm");
  writePushSubs_(function (rows) {
    const found = rows.filter(function (r) { return String(r.Sub_ID) === id; })[0];
    if (found) {
      found.Name = name;
      found.Endpoint = endpoint;
      found.P256dh = p256dh;
      found.Auth = auth;
      found.User_Agent = cleanTaskText(input.ua, 200);
      found.Updated_At = now;
      found.Fail_Count = 0;
      found.Last_Error = "";
      return rows;
    }
    rows.push({
      Sub_ID: id, Name: name, Endpoint: endpoint, P256dh: p256dh, Auth: auth,
      User_Agent: cleanTaskText(input.ua, 200), Created_At: now, Updated_At: now,
      Last_Push: "", Last_Status: "", Fail_Count: 0, Last_Error: ""
    });
    return rows;
  });
  pushKick_();
  return jsonResponse({ status: "success", sub_id: id }, e);
}

function handleDeletePushSub(input, e) {
  const id = subId_(cleanTaskText((input || {}).endpoint, 500));
  writePushSubs_(function (rows) {
    return rows.filter(function (r) { return String(r.Sub_ID) !== id; });
  });
  return jsonResponse({ status: "success" }, e);
}

/** สถานะการแจ้งเตือน — ให้หน้าเว็บรู้ว่าตั้งค่าแล้วหรือยัง และเอากุญแจไปลงทะเบียน */
function pushStatusPayload(endpoint) {
  const ready = pushReady_();
  const triggered = ScriptApp.getProjectTriggers()
    .some(function (t) { return t.getHandlerFunction() === "pushDueTasks"; });
  let registered = false;
  if (ready && endpoint) {
    const id = subId_(endpoint);
    registered = readPushSubs_().some(function (r) { return String(r.Sub_ID) === id; });
  }
  return { ready: ready, running: triggered, publicKey: vapidPublic_(), registered: registered };
}

/** ส่งแจ้งเตือนทดสอบไปเครื่องที่กด — หน่วงเวลาได้ เพื่อปิดแอปก่อนแล้วดูว่าเด้งจริง */
function handleTestPush(input, e) {
  input = input || {};
  if (!pushReady_()) return jsonResponse({ status: "error", message: "ยังไม่ได้ตั้งค่า Web Push (รัน setupPush ใน Apps Script)" }, e);
  const id = subId_(cleanTaskText(input.endpoint, 500));
  const sub = readPushSubs_().filter(function (r) { return String(r.Sub_ID) === id; })[0];
  if (!sub) return jsonResponse({ status: "error", message: "ยังไม่ได้ลงทะเบียนเครื่องนี้" }, e);

  const delay = Math.min(Math.max(Number(input.delay) || 0, 0), 30);
  if (delay > 0) Utilities.sleep(delay * 1000);

  const responses = sendToSubs_([sub], {
    items: [{ title: "PTS — ทดสอบแจ้งเตือน", body: "ถ้าเห็นข้อความนี้ แปลว่าใช้งานได้แล้วครับ", tag: "pts-test" }],
    badge: 0
  });
  const code = responses[0] ? responses[0].getResponseCode() : 0;
  recordPushResults_([sub], responses);
  return code >= 200 && code < 300
    ? jsonResponse({ status: "success", message: "ส่งแล้ว" }, e)
    : jsonResponse({ status: "error", message: "ส่งไม่สำเร็จ (รหัส " + code + ")" }, e);
}

/* ---------------- ข้อความ ---------------- */

const shortTitle_ = function (s) {
  s = cleanTaskText(s, 200);
  return s.length > 80 ? s.slice(0, 79) + "…" : s;
};

/** เวลาแบบ ms ของวัน d (yyyy-MM-dd) เวลา hm ตามเขตเวลาไทย */
function pushAt_(d, hm) {
  return new Date(d + "T" + hm + ":00" + Utilities.formatDate(new Date(), "Asia/Bangkok", "XXX")).getTime();
}
function pushNextDay_(d) {
  return Utilities.formatDate(new Date(pushAt_(d, "12:00") + 86400000), "Asia/Bangkok", "yyyy-MM-dd");
}

/** "วันนี้ 15:00 น." / "พรุ่งนี้" / "25 ต.ค." */
function dueLabel_(d, tm, today) {
  if (!d) return "ไม่ระบุวัน";
  const at = tm ? " " + tm + " น." : "";
  if (d === today) return "วันนี้" + at;
  if (d === pushNextDay_(today)) return "พรุ่งนี้" + at;
  const parts = d.split("-");
  return Number(parts[2]) + " " + TH_MONTHS_[Number(parts[1]) - 1] + at;
}

/** จำนวน Task ที่ต้องสนใจของคนนี้ — ใช้เป็นตัวเลขบนไอคอนแอป */
function alertCount_(rows, who, today) {
  return rows.filter(function (t) {
    return t.Status !== TASK_DONE && t.Assignee === who && t.Due_Date && t.Due_Date <= today;
  }).length;
}

/** ส่งชุดข้อความถึงทุกเครื่องของคนหนึ่ง */
function sendToSubs_(subs, payload) {
  if (!subs.length) return [];
  const requests = [];
  const usable = [];
  subs.forEach(function (s) {
    try { requests.push(pushRequest_(s, payload)); usable.push(s); }
    catch (err) { /* กุญแจของเครื่องนั้นเสีย - ข้ามไป */ }
  });
  if (!requests.length) return [];
  try { return UrlFetchApp.fetchAll(requests); }
  catch (err) { return []; }
}

/**
 * ส่งทันทีตอนบันทึก Task — ใช้กับการมอบหมายงานให้คนอื่นเท่านั้น
 * (งานครบกำหนด/เกินกำหนด เป็นหน้าที่ของตัวตั้งเวลา)
 */
function pushNow_(task, assignedBy) {
  if (!task || !assignedBy) return;
  const who = cleanTaskText(task.Assignee, 120);
  if (!who) return;
  const subs = readPushSubs_().filter(function (r) { return cleanTaskText(r.Name, 120) === who; });
  if (!subs.length) return;

  const today = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
  const rows = readAllTaskRows();
  const responses = sendToSubs_(subs, {
    items: [{
      title: "ได้รับมอบหมายงานใหม่",
      body: shortTitle_(task.Title) + "\nกำหนด " + dueLabel_(task.Due_Date, task.Due_Time, today) + " · โดย " + assignedBy,
      tag: "pts-task-" + task.Task_ID,
      taskId: task.Task_ID
    }],
    badge: alertCount_(rows, who, today)
  });
  recordPushResults_(subs, responses);
}

/** เรียกหลังบันทึก Task: ให้ตัวตั้งเวลาตรวจใหม่ + ส่งงานที่มอบหมายใหม่ทันที */
function afterTaskWrite_(task, assignedBy) {
  if (!pushReady_()) return;
  try {
    pushKick_();
    if (task && assignedBy) pushNow_(task, assignedBy);
  } catch (err) {
    // การแจ้งเตือนล้มเหลวต้องไม่ทำให้การบันทึก Task ล้มเหลวตาม
  }
}

/**
 * ขั้นที่ควรแจ้งของ Task หนึ่ง ณ เวลานี้
 * '' = ยังไม่ถึงเวลา · 'due' = ครบกำหนด · 'overdue' = เกินกำหนด
 */
function pushStage_(task, today, hm) {
  if (task.Status === TASK_DONE || !task.Due_Date) return "";
  if (task.Due_Date > today) return "";
  if (task.Due_Date < today) {
    // งานค้างข้ามวัน เตือนเฉพาะช่วงกลางวัน
    return (hm >= PUSH_DAY_START && hm <= PUSH_DAY_END) ? "overdue" : "";
  }
  if (task.Due_Time) return hm >= task.Due_Time ? "due" : "";
  return (hm >= PUSH_DAY_START && hm <= PUSH_DAY_END) ? "due" : "";
}

function pushItems_(list, today) {
  return list.slice(0, 4).map(function (t) {
    const overdue = t.Due_Date < today;
    return {
      title: overdue ? "เกินกำหนดแล้ว" : "ถึงกำหนดทำงานนี้",
      body: shortTitle_(t.Title) + "\n" + dueLabel_(t.Due_Date, t.Due_Time, today),
      tag: "pts-task-" + t.Task_ID,
      taskId: t.Task_ID,
      urgent: overdue
    };
  });
}

/**
 * รอบหน้าที่มีงานต้องส่งคือเมื่อไร — ให้ตัวตั้งเวลาส่วนใหญ่จบทันทีโดยไม่เปิดชีต
 */
function pushNextAt_(rows, now) {
  const today = Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd");
  let best = now.getTime() + PUSH_IDLE_MS;
  rows.forEach(function (t) {
    if (t.Status === TASK_DONE || !t.Due_Date) return;
    let at;
    if (t.Due_Date < today) at = pushAt_(today, PUSH_DAY_START);
    else if (t.Due_Date === today) at = t.Due_Time ? pushAt_(today, t.Due_Time) : pushAt_(today, PUSH_DAY_START);
    else at = t.Due_Time ? pushAt_(t.Due_Date, t.Due_Time) : pushAt_(t.Due_Date, PUSH_DAY_START);
    if (at > now.getTime() && at < best) best = at;
  });
  return best;
}

/** ตัวตั้งเวลาเรียกทุก 1 นาที */
function pushDueTasks() {
  if (!pushReady_()) return;
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const next = Number(props.getProperty("PTS_PUSH_NEXT") || "0");
  if (next && now.getTime() < next) return;   // ยังไม่ถึงรอบ - จบทันที ไม่เปิดชีต

  const today = Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd");
  const hm = Utilities.formatDate(now, "Asia/Bangkok", "HH:mm");
  const rows = readAllTaskRows();
  const subs = readPushSubs_();

  if (subs.length) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASKS_SHEET_NAME);
    const idx = sheet ? taskHeaderIndexes(sheet) : null;
    const stages = {};
    if (sheet && idx && idx.Push_Stage !== -1 && sheet.getLastRow() > 1) {
      const col = sheet.getRange(2, idx.Push_Stage + 1, sheet.getLastRow() - 1, 1).getValues();
      const ids = sheet.getRange(2, idx.Task_ID + 1, sheet.getLastRow() - 1, 1).getValues();
      ids.forEach(function (r, i) { stages[String(r[0]).trim()] = String(col[i][0] || "").trim(); });
    }

    // จัดกลุ่มตามผู้รับผิดชอบ แล้วส่งทีละคน (หนึ่งข้อความต่อเครื่อง)
    const byName = {};
    const marks = [];
    rows.forEach(function (t) {
      const stage = pushStage_(t, today, hm);
      if (!stage) return;
      const wanted = stage + ":" + today;
      if (stages[t.Task_ID] === wanted) return;   // แจ้งขั้นนี้ของวันนี้ไปแล้ว
      const who = cleanTaskText(t.Assignee, 120);
      if (!who) return;
      (byName[who] = byName[who] || []).push(t);
      marks.push({ id: t.Task_ID, value: wanted });
    });

    Object.keys(byName).forEach(function (who) {
      const mine = subs.filter(function (r) { return cleanTaskText(r.Name, 120) === who; });
      if (!mine.length) return;
      const list = byName[who].sort(function (a, b) {
        return (a.Due_Date + " " + (a.Due_Time || "24:00")).localeCompare(b.Due_Date + " " + (b.Due_Time || "24:00"));
      });
      const responses = sendToSubs_(mine, {
        items: pushItems_(list, today),
        badge: alertCount_(rows, who, today)
      });
      recordPushResults_(mine, responses);
    });

    if (marks.length) writePushStages_(marks);
  }

  props.setProperty("PTS_PUSH_NEXT", String(pushNextAt_(rows, now)));
}

/** จดว่าเตือน Task ถึงขั้นไหนแล้ว — เขียนเฉพาะคอลัมน์ Push_Stage ไม่แตะคอลัมน์อื่น */
function writePushStages_(marks) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASKS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;
  const idx = taskHeaderIndexes(sheet);
  if (idx.Push_Stage === -1 || idx.Task_ID === -1) return;
  const ids = sheet.getRange(2, idx.Task_ID + 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); });
  marks.forEach(function (m) {
    const at = ids.indexOf(m.id);
    if (at !== -1) sheet.getRange(at + 2, idx.Push_Stage + 1).setValue(m.value);
  });
}

/** บันทึกผลการส่ง — เครื่องที่ยกเลิกแล้ว (404/410) หรือพังติดกันหลายครั้ง ลบทิ้ง */
function recordPushResults_(subs, responses) {
  if (!subs.length || !responses.length) return;
  const result = {}, errText = {};
  subs.forEach(function (s, i) {
    const code = responses[i] ? responses[i].getResponseCode() : 0;
    result[s.Sub_ID] = code;
    errText[s.Sub_ID] = responses[i] && code >= 300 ? String(responses[i].getContentText() || "").slice(0, 200) : "";
  });
  const now = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm");
  writePushSubs_(function (rows) {
    return rows.filter(function (r) {
      if (!Object.prototype.hasOwnProperty.call(result, r.Sub_ID)) return true;
      const code = result[r.Sub_ID];
      if (code === 404 || code === 410) return false;   // เครื่องยกเลิกไปแล้ว
      r.Last_Push = now;
      r.Last_Status = code;
      r.Last_Error = errText[r.Sub_ID];
      r.Fail_Count = code >= 200 && code < 300 ? 0 : (Number(r.Fail_Count) || 0) + 1;
      return r.Fail_Count < PUSH_MAX_FAILS;
    });
  });
}

/** คำขอส่งแจ้งเตือน 1 เครื่อง (สำหรับ UrlFetchApp.fetchAll) */
function pushRequest_(sub, payload) {
  const endpoint = String(sub.Endpoint || "").trim();
  const aud = (endpoint.match(/^https:\/\/[^\/]+/) || [""])[0];
  if (!aud) throw new Error("endpoint ไม่ถูกต้อง");
  const body = encryptPush_(JSON.stringify(payload), String(sub.P256dh || "").trim(), String(sub.Auth || "").trim());
  return {
    url: endpoint,
    method: "post",
    muteHttpExceptions: true,
    contentType: "application/octet-stream",
    headers: {
      Authorization: "vapid t=" + vapidJwt_(aud) + ", k=" + vapidPublic_(),
      "Content-Encoding": "aes128gcm",
      TTL: "86400",
      Urgency: "high"
    },
    payload: Utilities.newBlob(s8_(body), "application/octet-stream")
  };
}

/** JWT ของ VAPID (ES256) — เก็บไว้ใช้ซ้ำ 6 ชั่วโมงต่อเซิร์ฟเวอร์ปลายทาง */
function vapidJwt_(aud) {
  const cache = CacheService.getScriptCache(), key = "vapid_" + aud;
  const hit = cache.get(key);
  if (hit) return hit;
  const enc = function (obj) { return b64u_(utf8_(JSON.stringify(obj))); };
  const unsigned = enc({ typ: "JWT", alg: "ES256" }) + "." +
    enc({ aud: aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: pushSubject_() });
  const jwt = unsigned + "." + b64u_(ecdsaSign_(utf8_(unsigned), vapidPrivate_()));
  cache.put(key, jwt, 6 * 3600);
  return jwt;
}

/* =========================================================================
 * เข้ารหัส — ไบต์ / ตัวเลขใหญ่
 * Apps Script ใช้ byte แบบ Java (-128..127) หน้าที่ของ s8_/u8_ คือแปลงไปมากับ 0..255
 * ========================================================================= */

const s8_ = a => Array.from(a, b => ((b & 255) > 127 ? (b & 255) - 256 : (b & 255)));
const u8_ = a => Array.from(a, b => b & 255);
const utf8_ = str => u8_(Utilities.newBlob(String(str)).getBytes());
const sha256_ = bytes => u8_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s8_(bytes)));
const hmac_ = (key, data) => u8_(Utilities.computeHmacSha256Signature(s8_(data), s8_(key)));
const b64u_ = bytes => Utilities.base64EncodeWebSafe(s8_(bytes)).replace(/=+$/, '');
function unb64u_(str) {
  let s = String(str || '').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  while (s.length % 4) s += '=';
  return s.length ? u8_(Utilities.base64DecodeWebSafe(s)) : [];
}
function bigToBytes_(x, len) {
  const out = new Array(len).fill(0);
  for (let i = len - 1; i >= 0; i--) { out[i] = Number(x & BigInt(255)); x >>= BigInt(8); }
  return out;
}
function bytesToBig_(bytes) {
  let x = BigInt(0);
  for (let i = 0; i < bytes.length; i++) x = (x << BigInt(8)) | BigInt(bytes[i] & 255);
  return x;
}
/** สุ่มไบต์ — Utilities.getUuid() มาจาก SecureRandom ของ Java (สุ่มแบบปลอดภัยสำหรับการเข้ารหัส) */
function randomBytes_(n) {
  let out = [];
  while (out.length < n) out = out.concat(sha256_(utf8_(Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid())));
  return out.slice(0, n);
}

/* ----- เส้นโค้ง P-256 (ใช้ทั้งลายเซ็น VAPID และ ECDH ตอนเข้ารหัสข้อความ) ----- */
const P256_ = {
  p: BigInt('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff'),
  n: BigInt('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551'),
  b: BigInt('0x5ac635d8aa3a93e7b3ebbd55769886bc651d06b0cc53b0f63bce3c3e27d2604b'),
  G: [BigInt('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'),
      BigInt('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')]
};
function mod_(a, m) { const r = a % m; return r < BigInt(0) ? r + m : r; }
function modPow_(b, e, m) {
  let r = BigInt(1); b = mod_(b, m);
  while (e > BigInt(0)) { if (e & BigInt(1)) r = (r * b) % m; b = (b * b) % m; e >>= BigInt(1); }
  return r;
}
const modInv_ = (a, m) => modPow_(a, m - BigInt(2), m);   // m เป็นจำนวนเฉพาะ

/** จุดบนเส้นโค้งอยู่ในพิกัด Jacobian [X, Y, Z] เพื่อเลี่ยงการหารทุกขั้น · null = จุดอนันต์ */
function pDouble_(P) {
  if (!P) return null;
  const p = P256_.p, X = P[0], Y = P[1], Z = P[2];
  if (Y === BigInt(0)) return null;
  const delta = Z * Z % p, gamma = Y * Y % p, beta = X * gamma % p;
  const alpha = BigInt(3) * mod_(X - delta, p) % p * ((X + delta) % p) % p;   // a = -3
  const X3 = mod_(alpha * alpha - BigInt(8) * beta, p);
  const Z3 = mod_((Y + Z) * (Y + Z) - gamma - delta, p);
  const Y3 = mod_(alpha * mod_(BigInt(4) * beta - X3, p) - BigInt(8) * gamma % p * gamma, p);
  return [X3, Y3, Z3];
}
function pAdd_(P, Q) {
  if (!P) return Q;
  if (!Q) return P;
  const p = P256_.p;
  const Z1Z1 = P[2] * P[2] % p, Z2Z2 = Q[2] * Q[2] % p;
  const U1 = P[0] * Z2Z2 % p, U2 = Q[0] * Z1Z1 % p;
  const S1 = P[1] * Q[2] % p * Z2Z2 % p, S2 = Q[1] * P[2] % p * Z1Z1 % p;
  const H = mod_(U2 - U1, p), r = mod_(S2 - S1, p);
  if (H === BigInt(0)) return r === BigInt(0) ? pDouble_(P) : null;
  const HH = H * H % p, HHH = H * HH % p, V = U1 * HH % p;
  const X3 = mod_(r * r - HHH - BigInt(2) * V, p);
  const Y3 = mod_(r * mod_(V - X3, p) - S1 * HHH, p);
  const Z3 = P[2] * Q[2] % p * H % p;
  return [X3, Y3, Z3];
}
/** k × จุด (affine [x, y]) → affine [x, y] */
function pMul_(k, A) {
  let R = null, Q = [A[0], A[1], BigInt(1)];
  while (k > BigInt(0)) { if (k & BigInt(1)) R = pAdd_(R, Q); Q = pDouble_(Q); k >>= BigInt(1); }
  if (!R) throw new Error('จุดอนันต์');
  const p = P256_.p, zi = modInv_(R[2], p), zi2 = zi * zi % p;
  return [R[0] * zi2 % p, R[1] * zi2 % p * zi % p];
}
function onCurve_(x, y) {
  const p = P256_.p;
  return mod_(y * y - (x * x % p * x - BigInt(3) * x + P256_.b), p) === BigInt(0);
}
function newKeyPair_() {
  let d;
  do { d = bytesToBig_(randomBytes_(32)); } while (d <= BigInt(0) || d >= P256_.n);
  const Q = pMul_(d, P256_.G);
  return { d: d, pub: [4].concat(bigToBytes_(Q[0], 32), bigToBytes_(Q[1], 32)) };
}

/** ลายเซ็น ECDSA P-256 + SHA-256 แบบ RFC 6979 (ค่า k คำนวณจากข้อความ ไม่ต้องพึ่งการสุ่ม) → r‖s 64 ไบต์ */
function ecdsaSign_(msg, d) {
  const n = P256_.n, h = sha256_(msg), e = bytesToBig_(h);
  const x = bigToBytes_(d, 32), h1 = bigToBytes_(mod_(e, n), 32);
  let V = new Array(32).fill(1), K = new Array(32).fill(0);
  K = hmac_(K, V.concat([0], x, h1)); V = hmac_(K, V);
  K = hmac_(K, V.concat([1], x, h1)); V = hmac_(K, V);
  for (;;) {
    V = hmac_(K, V);
    const k = bytesToBig_(V);
    if (k > BigInt(0) && k < n) {
      const r = mod_(pMul_(k, P256_.G)[0], n);
      if (r > BigInt(0)) {
        const s = mod_(modInv_(k, n) * (e + r * d), n);
        if (s > BigInt(0)) return bigToBytes_(r, 32).concat(bigToBytes_(s, 32));
      }
    }
    K = hmac_(K, V.concat([0])); V = hmac_(K, V);
  }
}

/* ----- AES-128-GCM (เข้ารหัสอย่างเดียว) ----- */
const AES_SBOX_ = (() => {
  const sbox = new Array(256), rotl = (v, s) => ((v << s) | (v >>> (8 - s))) & 255;
  let p = 1, q = 1;
  do {
    p = p ^ ((p << 1) & 255) ^ (p & 0x80 ? 0x1b : 0);   // p × 3
    q ^= q << 1; q ^= q << 2; q ^= q << 4; q &= 255;    // q ÷ 3
    if (q & 0x80) q ^= 0x09;
    sbox[p] = q ^ rotl(q, 1) ^ rotl(q, 2) ^ rotl(q, 3) ^ rotl(q, 4) ^ 0x63;
  } while (p !== 1);
  sbox[0] = 0x63;
  return sbox;
})();
const xtime_ = b => ((b << 1) ^ (b & 0x80 ? 0x1b : 0)) & 255;
function aesExpand_(key) {
  const w = key.slice();
  let rcon = 1;
  for (let i = 16; i < 176; i += 4) {
    let t = w.slice(i - 4, i);
    if (i % 16 === 0) {
      t = [AES_SBOX_[t[1]] ^ rcon, AES_SBOX_[t[2]], AES_SBOX_[t[3]], AES_SBOX_[t[0]]];
      rcon = ((rcon << 1) ^ (rcon & 0x80 ? 0x11b : 0)) & 255;
    }
    for (let j = 0; j < 4; j++) w.push(w[i - 16 + j] ^ t[j]);
  }
  return w;
}
function aesBlock_(w, input) {
  let s = input.map((b, i) => b ^ w[i]);
  for (let round = 1; round <= 10; round++) {
    const t = new Array(16);
    for (let c = 0; c < 4; c++) for (let r = 0; r < 4; r++) t[c * 4 + r] = AES_SBOX_[s[((c + r) % 4) * 4 + r]];
    if (round < 10) {
      for (let c = 0; c < 4; c++) {
        const a0 = t[c * 4], a1 = t[c * 4 + 1], a2 = t[c * 4 + 2], a3 = t[c * 4 + 3], all = a0 ^ a1 ^ a2 ^ a3;
        t[c * 4] = a0 ^ all ^ xtime_(a0 ^ a1);
        t[c * 4 + 1] = a1 ^ all ^ xtime_(a1 ^ a2);
        t[c * 4 + 2] = a2 ^ all ^ xtime_(a2 ^ a3);
        t[c * 4 + 3] = a3 ^ all ^ xtime_(a3 ^ a0);
      }
    }
    s = t.map((b, i) => b ^ w[round * 16 + i]);
  }
  return s;
}
function gmul_(X, Y) {
  const R = BigInt(0xe1) << BigInt(120);
  let Z = BigInt(0), V = Y;
  for (let i = 127; i >= 0; i--) {
    if ((X >> BigInt(i)) & BigInt(1)) Z ^= V;
    V = (V & BigInt(1)) ? (V >> BigInt(1)) ^ R : V >> BigInt(1);
  }
  return Z;
}
/** AES-128-GCM ไม่มี AAD → ciphertext ‖ tag 16 ไบต์ */
function aesGcmEncrypt_(key, iv, plain) {
  const w = aesExpand_(key);
  const H = bytesToBig_(aesBlock_(w, new Array(16).fill(0)));
  const j0 = iv.concat([0, 0, 0, 1]), ctr = j0.slice(), out = [];
  for (let i = 0; i < plain.length; i += 16) {
    for (let k = 15; k >= 12; k--) { ctr[k] = (ctr[k] + 1) & 255; if (ctr[k]) break; }
    const ks = aesBlock_(w, ctr);
    plain.slice(i, i + 16).forEach((b, j) => out.push(b ^ ks[j]));
  }
  let y = BigInt(0);
  for (let i = 0; i < out.length; i += 16) {
    const blk = out.slice(i, i + 16);
    while (blk.length < 16) blk.push(0);
    y = gmul_(y ^ bytesToBig_(blk), H);
  }
  y = gmul_(y ^ BigInt(out.length * 8), H);   // len(A)=0 ‖ len(C) เป็นบิต
  return out.concat(bigToBytes_(bytesToBig_(aesBlock_(w, j0)) ^ y, 16));
}

/* ----- เข้ารหัสข้อความแจ้งเตือน (RFC 8291, aes128gcm) ----- */
const hkdfExpand_ = (prk, info, len) => hmac_(prk, info.concat([1])).slice(0, len);
function encryptPush_(text, p256dhB64, authB64, fixed) {
  const uaPub = unb64u_(p256dhB64), authSecret = unb64u_(authB64);
  if (uaPub.length !== 65 || uaPub[0] !== 4) throw new Error('กุญแจของเครื่องไม่ถูกต้อง');
  const ua = [bytesToBig_(uaPub.slice(1, 33)), bytesToBig_(uaPub.slice(33))];
  if (!onCurve_(ua[0], ua[1])) throw new Error('กุญแจของเครื่องไม่อยู่บนเส้นโค้ง');
  const eph = (fixed && fixed.eph) || newKeyPair_();
  const salt = (fixed && fixed.salt) || randomBytes_(16);
  const ecdh = bigToBytes_(pMul_(eph.d, ua)[0], 32);
  const ikm = hkdfExpand_(hmac_(authSecret, ecdh), utf8_('WebPush: info').concat([0], uaPub, eph.pub), 32);
  const prk = hmac_(salt, ikm);
  const cek = hkdfExpand_(prk, utf8_('Content-Encoding: aes128gcm').concat([0]), 16);
  const nonce = hkdfExpand_(prk, utf8_('Content-Encoding: nonce').concat([0]), 12);
  const cipher = aesGcmEncrypt_(cek, nonce, utf8_(text).concat([2]));   // 0x02 = ระเบียนสุดท้าย
  return salt.concat([0, 0, 16, 0], [65], eph.pub, cipher);              // rs = 4096
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
