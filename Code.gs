// ══════════════════════════════════════════════════════════════════
//  Satija Paper – FMS Google Apps Script
// ══════════════════════════════════════════════════════════════════

const SHEET1_ID      = "1yq1zQy71b2X0DYU3FWeQKquWj4mBmxqF4067O-FW8RE";
const SHEET2_ID      = "1Cmxy8vlV-Q6nqezlOnazhW2wFMglarpk4D53xDzKyi4";
const MASTER_TAB     = "Master_FMS";
const SAMPLES_TAB    = "Samples";
const DRIVE_FOLDER_ID = "1uxvFTunw21NblGfgLUB8eSzr_ybviQvC";

const HANDOVER_OPTIONS = ["Mr. Mukesh", "Mr. Pawan", "Mr. Rishabh", "Mr. Sonu"];

// ── Master_FMS column map (1-based, 13 columns total) ────────────
const COL = {
  DOC_TYPE:        1,
  INVOICE_NO:      2,
  DATE:            3,
  PARTY:           4,
  ADDRESS:         5,
  ITEMS:           6,
  SALES_PERSON:    7,
  ASSIGNED_TO:     8,
  SIGN_DATE:       9,
  HANDOVER_BILL:   10,
  DOC_STATUS:      11,
  DOC_LINK:        12,
  ATTACHMENT_DATE: 13   // ← NEW: mandatory date when invoice is uploaded
};
const MASTER_COL_COUNT = 13;

// ── Samples sheet column map (1-based, 11 columns total) ─────────
// Columns 1-7 are backward-compatible with old rows.
// Columns 8-11 are new fields appended at end.
const SCOL = {
  TIMESTAMP:      1,
  PARTY_NAME:     2,
  LOCATION:       3,  // WMS Location
  CONTACT_PERSON: 4,
  PHONE:          5,
  SEND_NAME:      6,  // person name (if send mode = person)
  REMARKS:        7,
  SEND_MODE:      8,  // 'person' | 'courier'  (NEW)
  COURIER_DATE:   9,  // date when courier was sent  (NEW)
  COURIER_SENDER: 10, // name of courier sender       (NEW)
  FILE_LINK:      11  // Drive link to attached file  (NEW)
};
const SAMPLES_COL_COUNT = 11;


// ══════════════════════════════════════════════════════════════════
//  SETUP FUNCTIONS
// ══════════════════════════════════════════════════════════════════

function setupMasterHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let master = ss.getSheetByName(MASTER_TAB);
  if (!master) master = ss.insertSheet(MASTER_TAB);

  master.clear();
  const headers = [
    "Doc Type", "Invoice No.", "Date & Time", "Party / Firm Name",
    "Address / Location", "Merged Items & Details", "Sales Person",
    "Assigned To", "Party Sign Date", "Who to Handover Bill",
    "Doc Status", "Uploaded Doc Link", "Attachment Date"
  ];
  master.appendRow(headers);
  master.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold").setBackground("#cfe2f3");
  master.setFrozenRows(1);
  _applyHandoverValidation(master, 2, 1000);
  Logger.log("Master_FMS headers set (" + headers.length + " columns).");
}

// Safe migration: adds col-13 header without touching data
function addAttachmentDateColumn() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(MASTER_TAB);
  if (!master) { Logger.log("Master_FMS not found"); return; }

  // First, strip any columns beyond 13
  _cleanExtraColumnsOnSheet(master, MASTER_COL_COUNT);

  // Add/update col-13 header
  const cell = master.getRange(1, MASTER_COL_COUNT);
  if (!String(cell.getValue()).trim()) {
    cell.setValue("Attachment Date")
        .setFontWeight("bold").setBackground("#cfe2f3").setFontColor("#000000");
    Logger.log("'Attachment Date' header added at column 13.");
  } else {
    Logger.log("Column 13 header already set: " + cell.getValue());
  }
}

function setupSamplesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SAMPLES_TAB);
  if (!sheet) sheet = ss.insertSheet(SAMPLES_TAB);

  // Idempotent: always sets/refreshes the full header row
  const headers = [
    "Timestamp", "Party Name", "WMS Location",
    "Contact Person", "Phone",
    "Send By Name", "Remarks",
    "Send Mode", "Courier Date", "Courier Sender", "File Link"
  ];

  // Extend sheet columns if needed
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
       .setFontWeight("bold")
       .setBackground("#1e40af")
       .setFontColor("#ffffff");
  sheet.setFrozenRows(1);

  const widths = [140, 170, 140, 140, 120, 150, 230, 100, 115, 160, 200];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  Logger.log("Samples sheet ready with " + headers.length + " columns.");
  return sheet;
}

function _applyHandoverValidation(sheet, startRow, numRows) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(HANDOVER_OPTIONS, true)
    .setAllowInvalid(true).build();
  sheet.getRange(startRow, COL.HANDOVER_BILL, numRows, 1).setDataValidation(rule);
}


// ══════════════════════════════════════════════════════════════════
//  REAL-TIME AUTO-SYNC TRIGGER MANAGEMENT
// ══════════════════════════════════════════════════════════════════

// Run this ONCE from the Apps Script editor to enable real-time sync
function setupAutoSync() {
  // Remove any existing sync triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncBothSheetsSeparately') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Fire every 1 minute for near-real-time updates from Form 1 & Form 2
  ScriptApp.newTrigger('syncBothSheetsSeparately')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log("✅ Auto-sync enabled: Master_FMS updates every 1 minute from Form 1 & Form 2.");
}

// Run this to pause auto-sync
function removeAutoSync() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncBothSheetsSeparately') {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  Logger.log("Removed " + n + " auto-sync trigger(s).");
}


// ══════════════════════════════════════════════════════════════════
//  WEB APP ENTRY POINTS
// ══════════════════════════════════════════════════════════════════

function doGet(e) {
  const action = (e.parameter || {}).action;
  let out = { success: false, error: "Invalid action" };
  if (action === "fetchOrders")  out = { success: true, data: fetchOrdersFromMaster() };
  if (action === "fetchSamples") out = fetchSamplesData();
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let params = {};
  if (e.postData && e.postData.contents) {
    try { params = JSON.parse(e.postData.contents); } catch(_) { params = e.parameter || {}; }
  } else {
    params = e.parameter || {};
  }

  const action = params.action;
  let out = { success: false, error: "Invalid action" };

  if      (action === "fetchOrders")       out = { success: true, data: fetchOrdersFromMaster() };
  else if (action === "fetchSamples")      out = fetchSamplesData();
  else if (action === "submitSample")      out = submitSample(params);
  else if (action === "assignOrder")       out = updateAssignment(params.rowIndex, params.assignTo, params.byPost);
  else if (action === "uploadDoc")         out = processFileUpload(params);
  else if (action === "savePostedOrders")  out = savePostedOrders(params.orders);
  else if (action === "updateHandoverInfo") out = updateHandoverInfo(params);

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}


// ══════════════════════════════════════════════════════════════════
//  SAMPLE FUNCTIONS
// ══════════════════════════════════════════════════════════════════

function fetchSamplesData() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SAMPLES_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };

    const lastRow  = sheet.getLastRow();
    const lastCol  = Math.max(sheet.getLastColumn(), SAMPLES_COL_COUNT);
    const rows     = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    rows.reverse(); // newest first

    const data = rows.map(r => ({
      timestamp:     String(r[SCOL.TIMESTAMP      - 1] || ""),
      partyName:     String(r[SCOL.PARTY_NAME     - 1] || ""),
      location:      String(r[SCOL.LOCATION       - 1] || ""),
      contactPerson: String(r[SCOL.CONTACT_PERSON - 1] || ""),
      phone:         String(r[SCOL.PHONE          - 1] || ""),
      sendName:      String(r[SCOL.SEND_NAME      - 1] || ""),
      remarks:       String(r[SCOL.REMARKS        - 1] || ""),
      sendMode:      String(r[SCOL.SEND_MODE      - 1] || "person"),  // default 'person' for old rows
      courierDate:   String(r[SCOL.COURIER_DATE   - 1] || ""),
      courierSender: String(r[SCOL.COURIER_SENDER - 1] || ""),
      fileLink:      String(r[SCOL.FILE_LINK      - 1] || "")
    }));

    return { success: true, data };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function submitSample(params) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SAMPLES_TAB) || setupSamplesSheet();

    const now = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy HH:mm");

    // Upload attached file to Drive (same system as invoice uploads)
    let fileLink = "";
    if (params.fileData && params.fileName) {
      try {
        const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        const blob    = Utilities.newBlob(
          Utilities.base64Decode(params.fileData),
          params.mimeType || "application/octet-stream",
          params.fileName
        );
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileLink = file.getUrl();
      } catch (uploadErr) {
        Logger.log("Sample file upload error: " + uploadErr.message);
        // Non-fatal: continue saving the row without a file link
      }
    }

    // Append row in the new 11-column layout
    sheet.appendRow([
      params.timestamp     || now,       // 1  Timestamp
      params.partyName     || "",        // 2  Party Name
      params.location      || "",        // 3  WMS Location
      params.contactPerson || "",        // 4  Contact Person
      params.phone         || "",        // 5  Phone
      params.sendName      || "",        // 6  Send By Name
      params.remarks       || "",        // 7  Remarks
      params.sendMode      || "person",  // 8  Send Mode
      params.courierDate   || "",        // 9  Courier Date
      params.courierSender || "",        // 10 Courier Sender
      fileLink                           // 11 File Link
    ]);

    return { success: true, fileLink };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════════════
//  ORDERS FUNCTIONS
// ══════════════════════════════════════════════════════════════════

function fetchOrdersFromMaster() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_TAB);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const list = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    // Skip rows where invoice number looks like a date or is blank
    const inv = String(r[COL.INVOICE_NO - 1]).trim();
    if (!inv || r[COL.INVOICE_NO - 1] instanceof Date) continue;

    list.push({
      rowIndex:       i + 1,
      docType:        String(r[COL.DOC_TYPE        - 1] || ""),
      docNo:          inv,
      date: (function() {
        const v = r[COL.DATE - 1];
        if (v instanceof Date) return Utilities.formatDate(v, "GMT+5:30", "dd/MM/yyyy HH:mm");
        const s = String(v).trim();
        if (!s) return '';
        try { const d = new Date(s); if (!isNaN(d)) return Utilities.formatDate(d, "GMT+5:30", "dd/MM/yyyy HH:mm"); } catch(_) {}
        return s;
      })(),
      party:          String(r[COL.PARTY           - 1] || ""),
      address:        String(r[COL.ADDRESS         - 1] || ""),
      items:          String(r[COL.ITEMS           - 1] || ""),
      salesPerson:    String(r[COL.SALES_PERSON    - 1] || "").trim(),
      assignedTo:     String(r[COL.ASSIGNED_TO     - 1] || "Unassigned").trim(),
      partySignDate:  String(r[COL.SIGN_DATE       - 1] || "").trim(),
      handoverBill:   String(r[COL.HANDOVER_BILL   - 1] || "").trim(),
      status:         String(r[COL.DOC_STATUS      - 1] || "Pending").trim(),
      docLink:        String(r[COL.DOC_LINK        - 1] || "").trim(),
      attachmentDate: String(r[COL.ATTACHMENT_DATE - 1] || "").trim()  // NEW
    });
  }
  return list;
}

function updateAssignment(rowIndex, assignTo, byPost) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_TAB);
  sheet.getRange(rowIndex, COL.ASSIGNED_TO).setValue(assignTo);
  if (byPost === true || byPost === "true") {
    sheet.getRange(rowIndex, COL.DOC_STATUS).setValue("Uploaded");
  }
  return { success: true };
}

// Handles: attachmentDate, partySignDate, handoverBill — any combination
function updateHandoverInfo(params) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_TAB);
    const row   = parseInt(params.rowIndex);

    if (params.attachmentDate !== undefined && params.attachmentDate !== null) {
      sheet.getRange(row, COL.ATTACHMENT_DATE).setValue(params.attachmentDate);
    }
    if (params.partySignDate !== undefined && params.partySignDate !== null) {
      sheet.getRange(row, COL.SIGN_DATE).setValue(params.partySignDate);
    }
    if (params.handoverBill !== undefined && params.handoverBill !== null) {
      sheet.getRange(row, COL.HANDOVER_BILL).setValue(params.handoverBill);
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// Upload invoice file + save attachment date (NEW: params.attachmentDate)
function processFileUpload(params) {
  try {
    const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const blob    = Utilities.newBlob(
      Utilities.base64Decode(params.fileData),
      params.mimeType,
      params.fileName
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_TAB);
    const row   = parseInt(params.rowIndex);
    sheet.getRange(row, COL.DOC_STATUS).setValue("Uploaded");
    sheet.getRange(row, COL.DOC_LINK).setValue(file.getUrl());
    if (params.attachmentDate) {
      sheet.getRange(row, COL.ATTACHMENT_DATE).setValue(params.attachmentDate);
    }

    return { success: true, url: file.getUrl() };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


// ══════════════════════════════════════════════════════════════════
//  SYNC FROM FORM 1 (Sales Order) + FORM 2 (Direct Dispatch)
// ══════════════════════════════════════════════════════════════════

function syncBothSheetsSeparately() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let master = ss.getSheetByName(MASTER_TAB);
  if (!master) { setupMasterHeaders(); master = ss.getSheetByName(MASTER_TAB); }

  // Ensure column count is exactly MASTER_COL_COUNT; strip extras
  _cleanExtraColumnsOnSheet(master, MASTER_COL_COUNT);

  // Build index of existing rows  { "DocType_InvoiceNo" → rowNumber }
  const masterData   = master.getRange(1, 1, master.getLastRow(), MASTER_COL_COUNT).getValues();
  const existingMap  = {};
  for (let r = 1; r < masterData.length; r++) {
    const docType = String(masterData[r][0]).trim();
    const docNo   = String(masterData[r][1]).trim();
    if (docNo && !(masterData[r][1] instanceof Date)) {
      existingMap[`${docType}_${docNo}`] = r + 1;
    }
  }

  const s1Group = {};
  processSheet1Data(s1Group);
  const s2Group = {};
  processSheet2Data(s2Group);

  writeGroupToMaster(master, s1Group, existingMap, "Sales Order");
  writeGroupToMaster(master, s2Group, existingMap, "Direct Dispatch");

  Logger.log("Sync complete. Master rows: " + (master.getLastRow() - 1));
}

function processSheet1Data(group) {
  try {
    const sheet   = SpreadsheetApp.openById(SHEET1_ID).getSheetByName("FMS");
    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return;
    const data = sheet.getRange(7, 1, lastRow - 6, 57).getValues();

    data.forEach(r => {
      const invoiceNo = String(r[56]).trim();
      if (!invoiceNo) return;

      if (!group[invoiceNo]) {
        group[invoiceNo] = {
          docNo:       invoiceNo,
          date:        r[0],
          party:       r[2],
          address:     `${r[4]} (GST: ${r[3]})`,
          salesPerson: String(r[7] || "").trim(),
          items:       []
        };
      }
      const prod  = r[18];
      const qReam = Number(r[19]) || 0;
      const qBox  = Number(r[20]) || 0;
      if (prod) group[invoiceNo].items.push(`• ${prod} | Ream: ${qReam}, Box: ${qBox}`);
    });
  } catch (e) { Logger.log("Sheet1 error: " + e.message); }
}

function processSheet2Data(group) {
  try {
    const sheet   = SpreadsheetApp.openById(SHEET2_ID).getSheetByName("FMS");
    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return;
    const data = sheet.getRange(7, 1, lastRow - 6, 67).getValues();

    data.forEach(r => {
      const invoiceNo = String(r[66]).trim();
      if (!invoiceNo) return;

      if (!group[invoiceNo]) {
        group[invoiceNo] = {
          docNo:       invoiceNo,
          date:        r[0],
          party:       String(r[20] || "").trim(),
          address:     String(r[3]  || "").trim(),
          salesPerson: String(r[25] || "").trim(),
          items:       []
        };
      }
      const qReam = Number(r[4]) || 0;
      const qBox  = Number(r[5]) || 0;
      const qKg   = Number(r[6]) || 0;
      group[invoiceNo].items.push(`• Ream: ${qReam}, Box: ${qBox}, Kg: ${qKg}`);
    });
  } catch (e) { Logger.log("Sheet2 error: " + e.message); }
}

function writeGroupToMaster(masterSheet, groupObj, existingMap, docType) {
  Object.keys(groupObj).forEach(docNo => {
    const item      = groupObj[docNo];
    const key       = `${docType}_${docNo}`;
    const itemsText = item.items.join("\n");

    if (existingMap[key]) {
      // Update only data columns 1-7 (preserve assignment/status/upload cols 8-13)
      masterSheet.getRange(existingMap[key], COL.DOC_TYPE, 1, 7).setValues([[
        docType, item.docNo, item.date, item.party, item.address, itemsText, item.salesPerson
      ]]);
    } else {
      // New row — write exactly MASTER_COL_COUNT columns
      masterSheet.appendRow([
        docType, item.docNo, item.date, item.party, item.address, itemsText,
        item.salesPerson, "Unassigned", "", "", "Pending", "", ""
      ]);
      const newRow = masterSheet.getLastRow();
      existingMap[key] = newRow;
      _applyHandoverValidation(masterSheet, newRow, 1);
    }
  });
}


// ══════════════════════════════════════════════════════════════════
//  POSTED ORDERS
// ══════════════════════════════════════════════════════════════════

function savePostedOrders(orders) {
  try {
    if (!orders || !orders.length) return { success: false, error: "No orders provided" };

    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const name = "Posted_Orders";
    let dest   = ss.getSheetByName(name);
    if (!dest) {
      dest = ss.insertSheet(name);
      dest.appendRow([
        "Doc Type", "Invoice No.", "Date & Time", "Party / Firm Name",
        "Address / Location", "Merged Items & Details", "Sales Person",
        "Assigned To", "Party Sign Date", "Who to Handover Bill", "Doc Status", "Saved On"
      ]);
      dest.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#fce8d3");
      dest.setFrozenRows(1);
      _applyHandoverValidation(dest, 2, 1000);
    }

    const now = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy HH:mm");
    orders.forEach(o => dest.appendRow([
      o.docType, o.docNo, o.date, o.party, o.address,
      o.items, o.salesPerson, o.assignedTo,
      o.partySignDate, o.handoverBill, o.status, now
    ]));

    return { success: true, sheetName: name };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


// ══════════════════════════════════════════════════════════════════
//  CLEANUP UTILITIES  ← RUN THESE FROM EDITOR TO FIX SHEET
// ══════════════════════════════════════════════════════════════════

// STEP 1 – Run this to delete all extra columns (col 14+) from Master_FMS
function cleanExtraColumns() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(MASTER_TAB);
  if (!master) { Logger.log("Master_FMS not found"); return; }
  const removed = _cleanExtraColumnsOnSheet(master, MASTER_COL_COUNT);
  Logger.log(removed > 0
    ? `✅ Cleaned ${removed} extra column(s). Sheet now has ${MASTER_COL_COUNT} columns.`
    : "✅ No extra columns found – sheet is clean.");
}

function _cleanExtraColumnsOnSheet(sheet, keepCols) {
  const lastCol = sheet.getLastColumn();
  if (lastCol <= keepCols) return 0;
  const extra = lastCol - keepCols;
  sheet.deleteColumns(keepCols + 1, extra);
  return extra;
}

// STEP 2 – Run this to add the Attachment Date header (col 13) safely
// (does NOT clear existing data)
function addAttachmentDateHeader() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(MASTER_TAB);
  if (!master) { Logger.log("Master_FMS not found"); return; }
  _cleanExtraColumnsOnSheet(master, MASTER_COL_COUNT);
  const cell = master.getRange(1, MASTER_COL_COUNT);
  cell.setValue("Attachment Date")
      .setFontWeight("bold").setBackground("#cfe2f3").setFontColor("#000000");
  Logger.log("'Attachment Date' header set at column 13.");
}

// STEP 3 – Full cleanup: strip extra cols, fix bad rows, resync from both source sheets
function cleanupAndResync() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(MASTER_TAB);
  if (!master) { Logger.log("Master_FMS not found"); return; }

  // Strip extra columns first
  _cleanExtraColumnsOnSheet(master, MASTER_COL_COUNT);

  // Delete rows with missing or Date-typed invoice numbers
  const data     = master.getRange(1, 1, master.getLastRow(), MASTER_COL_COUNT).getValues();
  const toDelete = [];
  for (let i = 1; i < data.length; i++) {
    const inv = data[i][COL.INVOICE_NO - 1];
    if (inv instanceof Date || !String(inv).trim()) toDelete.push(i + 1);
  }
  for (let r = toDelete.length - 1; r >= 0; r--) master.deleteRow(toDelete[r]);
  Logger.log("Removed " + toDelete.length + " bad rows from Master_FMS.");

  // Re-sync live data from both source sheets
  syncBothSheetsSeparately();
  Logger.log("✅ cleanupAndResync complete.");
}

// Migrate samples sheet to 11-column layout (safe – only updates headers)
function migrateSamplesSheet() {
  setupSamplesSheet();
  Logger.log("✅ Samples sheet migrated to 11-column layout.");
}


// ══════════════════════════════════════════════════════════════════
//  LEGACY MIGRATION UTILITIES (run once if needed)
// ══════════════════════════════════════════════════════════════════

function rearrangeMasterSheet() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  let master   = ss.getSheetByName(MASTER_TAB);
  if (!master) { setupMasterHeaders(); return; }

  const allData = master.getDataRange().getValues();
  if (allData.length <= 1) { setupMasterHeaders(); return; }
  if (String(allData[0][6]).trim() === "Sales Person") {
    Logger.log("Sheet already in new layout."); return;
  }

  const newHeaders = [
    "Doc Type", "Invoice No.", "Date & Time", "Party / Firm Name",
    "Address / Location", "Merged Items & Details", "Sales Person",
    "Assigned To", "Party Sign Date", "Who to Handover Bill",
    "Doc Status", "Uploaded Doc Link", "Attachment Date"
  ];
  const newRows = [newHeaders];
  for (let i = 1; i < allData.length; i++) {
    const r = allData[i];
    newRows.push([r[0],r[1],r[2],r[3],r[4],r[5],"",r[9],"","",r[10],r[11],""]);
  }

  master.clear();
  master.getRange(1, 1, newRows.length, MASTER_COL_COUNT).setValues(newRows);
  master.getRange(1, 1, 1, MASTER_COL_COUNT).setFontWeight("bold").setBackground("#cfe2f3");
  master.setFrozenRows(1);
  if (newRows.length > 1) _applyHandoverValidation(master, 2, 1000);
  Logger.log("✅ rearrangeMasterSheet done. " + (newRows.length - 1) + " rows migrated.");
}

function migratePostedOrdersSheet() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("Posted_Orders");
  if (!master) { Logger.log("Posted_Orders not found"); return; }

  const allData = master.getDataRange().getValues();
  if (allData.length <= 1 || String(allData[0][6]).trim() === "Sales Person") {
    Logger.log("Posted_Orders already migrated."); return;
  }

  const newHeaders = [
    "Doc Type", "Invoice No.", "Date & Time", "Party / Firm Name",
    "Address / Location", "Merged Items & Details", "Sales Person",
    "Assigned To", "Party Sign Date", "Who to Handover Bill", "Doc Status", "Saved On"
  ];
  const newRows = [newHeaders];
  for (let i = 1; i < allData.length; i++) {
    const r = allData[i];
    newRows.push([r[0],r[1],r[2],r[3],r[4],r[5],"",r[9],"","",r[10],r[11]]);
  }

  master.clear();
  master.getRange(1, 1, newRows.length, 12).setValues(newRows);
  master.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#fce8d3");
  master.setFrozenRows(1);
  if (newRows.length > 1) _applyHandoverValidation(master, 2, 1000);
  Logger.log("✅ Posted_Orders migrated. " + (newRows.length - 1) + " rows done.");
}


// ══════════════════════════════════════════════════════════════════
//  DRIVE PERMISSION (run once)
// ══════════════════════════════════════════════════════════════════

function grantDrivePermission() {
  DriveApp.getFolderById(DRIVE_FOLDER_ID);
  Logger.log("Drive permission granted.");
}

function forceFullDrivePermission() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const test   = folder.createFile("_permission_check.txt", "OK");
  test.setTrashed(true);
  Logger.log("Full Drive access confirmed.");
}
