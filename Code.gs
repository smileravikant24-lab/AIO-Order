const SHEET1_ID = "1yq1zQy71b2X0DYU3FWeQKquWj4mBmxqF4067O-FW8RE";
const SHEET2_ID = "1Cmxy8vlV-Q6nqezlOnazhW2wFMglarpk4D53xDzKyi4";
const MASTER_TAB = "Master_FMS";
const DRIVE_FOLDER_ID = "1uxvFTunw21NblGfgLUB8eSzr_ybviQvC";

const HANDOVER_OPTIONS = ["Mr. Mukesh", "Mr. Pawan", "Mr. Rishabh", "Mr. Sonu"];

// Master column positions (1-based)
const COL = {
  DOC_TYPE:      1,
  INVOICE_NO:    2,
  DATE:          3,
  PARTY:         4,
  ADDRESS:       5,
  ITEMS:         6,
  SALES_PERSON:  7,
  ASSIGNED_TO:   8,
  SIGN_DATE:     9,
  HANDOVER_BILL: 10,
  DOC_STATUS:    11,
  DOC_LINK:      12
};

function setupMasterHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let master = ss.getSheetByName(MASTER_TAB);
  if (!master) master = ss.insertSheet(MASTER_TAB);

  const headers = [
    "Doc Type", "Invoice No.", "Date & Time", "Party / Firm Name",
    "Address / Location", "Merged Items & Details", "Sales Person",
    "Assigned To", "Party Sign Date", "Who to Handover Bill", "Doc Status", "Uploaded Doc Link"
  ];

  master.clear();
  master.appendRow(headers);
  master.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#cfe2f3");
  master.setFrozenRows(1);

  // Dropdown validation for "Who to Handover Bill" column (col 10), rows 2 onwards
  _applyHandoverValidation(master, 2, 1000);
}

// Apply dropdown validation to a range of rows in the handover column
function _applyHandoverValidation(sheet, startRow, numRows) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(HANDOVER_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, COL.HANDOVER_BILL, numRows, 1).setDataValidation(rule);
}


function syncBothSheetsSeparately() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let master = ss.getSheetByName(MASTER_TAB);
  if (!master) { setupMasterHeaders(); master = ss.getSheetByName(MASTER_TAB); }

  const masterData = master.getDataRange().getValues();
  const existingMap = {};
  for (let r = 1; r < masterData.length; r++) {
    const docType = String(masterData[r][0]).trim();
    const docNo   = String(masterData[r][1]).trim();
    if (docNo) existingMap[`${docType}_${docNo}`] = r + 1;
  }

  const s1Group = {};
  processSheet1Data(s1Group);
  const s2Group = {};
  processSheet2Data(s2Group);

  writeGroupToMaster(master, s1Group, existingMap, "Sales Order");
  writeGroupToMaster(master, s2Group, existingMap, "Direct Dispatch");
}

// Sheet1 FMS: Invoice No. at column BF = index 57 (r[57]), Sales Person at column H = index 7 (r[7]).
function processSheet1Data(group) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET1_ID).getSheetByName("FMS");
    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return;
    const data = sheet.getRange(7, 1, lastRow - 6, 58).getValues();

    data.forEach(r => {
      const invoiceNo = String(r[57]).trim(); // column BF = index 57
      if (!invoiceNo) return;

      if (!group[invoiceNo]) {
        group[invoiceNo] = {
          docNo: invoiceNo,
          date: r[0],
          party: r[2],
          address: `${r[4]} (GST: ${r[3]})`,
          salesPerson: String(r[7] || "").trim(), // column H = index 7
          items: []
        };
      }
      const prod = r[18];
      const qReam = Number(r[19]) || 0;
      const qBox  = Number(r[20]) || 0;

      if (prod) group[invoiceNo].items.push(`• ${prod} | Ream: ${qReam}, Box: ${qBox}`);
    });
  } catch (e) { Logger.log("S1 Err: " + e.message); }
}

// Sheet2 FMS: Invoice No. at column BO = index 66 (r[66]), Sales Person at column Z = index 25 (r[25]).
function processSheet2Data(group) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET2_ID).getSheetByName("FMS");
    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return;
    const data = sheet.getRange(7, 1, lastRow - 6, 67).getValues();

    data.forEach(r => {
      const invoiceNo = String(r[66]).trim(); // column BO = index 66
      if (!invoiceNo) return;

      if (!group[invoiceNo]) {
        group[invoiceNo] = {
          docNo: invoiceNo,
          date: r[0],
          party: String(r[20] || "").trim(),       // Firm Name at column U (index 20)
          address: String(r[3] || "").trim(),      // Dispatch Address at column D (index 3)
          salesPerson: String(r[25] || "").trim(), // Sales Person at column Z (index 25)
          items: []
        };
      }
      const qReam = Number(r[4]) || 0;
      const qBox  = Number(r[5]) || 0;
      const qKg   = Number(r[6]) || 0;

      group[invoiceNo].items.push(`• Ream: ${qReam}, Box: ${qBox}, Kg: ${qKg}`);
    });
  } catch (e) { Logger.log("S2 Err: " + e.message); }
}

function writeGroupToMaster(masterSheet, groupObj, existingMap, docType) {
  Object.keys(groupObj).forEach(docNo => {
    const item = groupObj[docNo];
    const key  = `${docType}_${docNo}`;
    const itemsText = item.items.join("\n");

    if (existingMap[key]) {
      // Update only source-synced columns; preserve Assigned To, Sign Date, Handover, Status, Link
      const r = existingMap[key];
      masterSheet.getRange(r, COL.DOC_TYPE, 1, 7).setValues([[
        docType, item.docNo, item.date, item.party, item.address, itemsText, item.salesPerson
      ]]);
    } else {
      masterSheet.appendRow([
        docType, item.docNo, item.date, item.party, item.address, itemsText,
        item.salesPerson, "Unassigned", "", "", "Pending", ""
      ]);
      const newRow = masterSheet.getLastRow();
      existingMap[key] = newRow;
      // Ensure dropdown applies to the new row
      _applyHandoverValidation(masterSheet, newRow, 1);
    }
  });
}


function doGet(e) {
  const params = e.parameter || {};
  const action = params.action;
  let responseData = { success: false, error: "Invalid action" };

  if (action === "fetchOrders") {
    responseData = { success: true, data: fetchOrdersFromMaster() };
  }

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let params = {};
  if (e.postData && e.postData.contents) {
    try { params = JSON.parse(e.postData.contents); }
    catch(err) { params = e.parameter || {}; }
  } else {
    params = e.parameter || {};
  }

  const action = params.action;
  let responseData = { success: false, error: "Invalid action" };

  if (action === "fetchOrders") {
    responseData = { success: true, data: fetchOrdersFromMaster() };
  } else if (action === "assignOrder") {
    responseData = updateAssignment(params.rowIndex, params.assignTo, params.byPost);
  } else if (action === "uploadDoc") {
    responseData = processFileUpload(params.rowIndex, params.fileData, params.fileName, params.mimeType);
  } else if (action === "savePostedOrders") {
    responseData = savePostedOrders(params.orders);
  } else if (action === "updateHandoverInfo") {
    responseData = updateHandoverInfo(params.rowIndex, params.partySignDate, params.handoverBill);
  }

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}


function fetchOrdersFromMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_TAB);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    list.push({
      rowIndex:     i + 1,
      docType:      String(row[COL.DOC_TYPE      - 1]),
      docNo:        String(row[COL.INVOICE_NO    - 1]),
      date:         row[COL.DATE - 1] instanceof Date
                      ? Utilities.formatDate(row[COL.DATE - 1], "GMT+5:30", "dd/MM/yyyy HH:mm")
                      : String(row[COL.DATE - 1]),
      party:        String(row[COL.PARTY         - 1]),
      address:      String(row[COL.ADDRESS       - 1]),
      items:        String(row[COL.ITEMS         - 1]),
      salesPerson:  String(row[COL.SALES_PERSON  - 1] || "").trim(),
      assignedTo:   String(row[COL.ASSIGNED_TO   - 1] || "Unassigned").trim(),
      partySignDate:String(row[COL.SIGN_DATE      - 1] || "").trim(),
      handoverBill: String(row[COL.HANDOVER_BILL - 1] || "").trim(),
      status:       String(row[COL.DOC_STATUS    - 1] || "Pending").trim(),
      docLink:      String(row[COL.DOC_LINK      - 1] || "").trim()
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

// New: update Party Sign Date and/or Who to Handover Bill for a row
function updateHandoverInfo(rowIndex, partySignDate, handoverBill) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_TAB);
    if (partySignDate !== undefined && partySignDate !== null) {
      sheet.getRange(rowIndex, COL.SIGN_DATE).setValue(partySignDate);
    }
    if (handoverBill !== undefined && handoverBill !== null) {
      if (handoverBill && !HANDOVER_OPTIONS.includes(handoverBill)) {
        return { success: false, error: "Invalid handover option: " + handoverBill };
      }
      sheet.getRange(rowIndex, COL.HANDOVER_BILL).setValue(handoverBill);
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function processFileUpload(rowIndex, base64Data, fileName, mimeType) {
  try {
    const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const decoded = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(decoded, mimeType, fileName);
    const file    = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_TAB);
    sheet.getRange(rowIndex, COL.DOC_STATUS).setValue("Uploaded");
    sheet.getRange(rowIndex, COL.DOC_LINK).setValue(file.getUrl());

    return { success: true, url: file.getUrl() };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function savePostedOrders(orders) {
  try {
    if (!orders || !orders.length) {
      return { success: false, error: "No orders provided" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Posted_Orders";
    let dest = ss.getSheetByName(sheetName);

    if (!dest) {
      dest = ss.insertSheet(sheetName);
      const headers = [
        "Doc Type", "Invoice No.", "Date & Time", "Party / Firm Name",
        "Address / Location", "Merged Items & Details", "Sales Person",
        "Assigned To", "Party Sign Date", "Who to Handover Bill", "Doc Status", "Saved On"
      ];
      dest.appendRow(headers);
      dest.getRange(1, 1, 1, headers.length)
          .setFontWeight("bold")
          .setBackground("#fce8d3");
      dest.setFrozenRows(1);
      _applyHandoverValidation(dest, 2, 1000);
    }

    const now = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy HH:mm");
    orders.forEach(o => {
      dest.appendRow([
        o.docType, o.docNo, o.date, o.party, o.address,
        o.items, o.salesPerson, o.assignedTo,
        o.partySignDate, o.handoverBill, o.status, now
      ]);
    });

    return { success: true, sheetName: sheetName };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// Run this ONCE to migrate existing Master_FMS data to the new column layout.
// Old: DocType | Order/PO No | Date | Party | Address | Items | TotalRem | TotalBox | TotalKg | AssignedTo | DocStatus | DocLink
// New: DocType | Invoice No. | Date | Party | Address | Items | SalesPerson | AssignedTo | SignDate | HandoverBill | DocStatus | DocLink
function rearrangeMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let master = ss.getSheetByName(MASTER_TAB);
  if (!master) { setupMasterHeaders(); return; }

  const allData = master.getDataRange().getValues();

  // If sheet is empty or only has header, just reset headers
  if (allData.length <= 1) { setupMasterHeaders(); return; }

  // Check if already in new layout (7th header = "Sales Person")
  if (String(allData[0][6]).trim() === "Sales Person") {
    Logger.log("Sheet already in new layout. Nothing to do.");
    return;
  }

  const newHeaders = [
    "Doc Type", "Invoice No.", "Date & Time", "Party / Firm Name",
    "Address / Location", "Merged Items & Details", "Sales Person",
    "Assigned To", "Party Sign Date", "Who to Handover Bill", "Doc Status", "Uploaded Doc Link"
  ];

  const newRows = [newHeaders];
  for (let i = 1; i < allData.length; i++) {
    const r = allData[i];
    newRows.push([
      r[0],   // Doc Type
      r[1],   // Order/PO No → Invoice No.
      r[2],   // Date & Time
      r[3],   // Party / Firm Name
      r[4],   // Address / Location
      r[5],   // Merged Items & Details
      "",     // Sales Person (new — empty; will fill on next sync)
      r[9],   // Assigned To (was index 9 / col 10)
      "",     // Party Sign Date (new — empty)
      "",     // Who to Handover Bill (new — empty)
      r[10],  // Doc Status (was index 10 / col 11)
      r[11]   // Uploaded Doc Link (was index 11 / col 12)
    ]);
  }

  master.clear();
  master.getRange(1, 1, newRows.length, 12).setValues(newRows);
  master.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#cfe2f3");
  master.setFrozenRows(1);

  if (newRows.length > 1) _applyHandoverValidation(master, 2, 1000);

  Logger.log("Done. " + (newRows.length - 1) + " rows migrated to new layout.");
}

// Migrate Posted_Orders sheet to new column layout (run once).
function migratePostedOrdersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Posted_Orders";
  const master = ss.getSheetByName(sheetName);
  if (!master) { Logger.log("Posted_Orders sheet not found."); return; }

  const allData = master.getDataRange().getValues();
  if (allData.length <= 1) { Logger.log("Sheet empty, nothing to migrate."); return; }

  // Already in new layout?
  if (String(allData[0][6]).trim() === "Sales Person") {
    Logger.log("Posted_Orders already in new layout."); return;
  }

  const newHeaders = [
    "Doc Type", "Invoice No.", "Date & Time", "Party / Firm Name",
    "Address / Location", "Merged Items & Details", "Sales Person",
    "Assigned To", "Party Sign Date", "Who to Handover Bill", "Doc Status", "Saved On"
  ];

  // Old: DocType|Order/PO No|Date|Party|Address|Items|TotalRem|TotalBox|TotalKg|AssignedTo|DocStatus|SavedOn
  // New: DocType|Invoice No.|Date|Party|Address|Items|SalesPerson|AssignedTo|SignDate|HandoverBill|DocStatus|SavedOn
  const newRows = [newHeaders];
  for (let i = 1; i < allData.length; i++) {
    const r = allData[i];
    newRows.push([
      r[0], r[1], r[2], r[3], r[4], r[5],
      "",    // Sales Person (new — empty)
      r[9],  // Assigned To (was index 9)
      "",    // Party Sign Date (new — empty)
      "",    // Who to Handover Bill (new — empty)
      r[10], // Doc Status (was index 10)
      r[11]  // Saved On (was index 11)
    ]);
  }

  master.clear();
  master.getRange(1, 1, newRows.length, 12).setValues(newRows);
  master.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#fce8d3");
  master.setFrozenRows(1);
  if (newRows.length > 1) _applyHandoverValidation(master, 2, 1000);

  Logger.log("Posted_Orders migrated. " + (newRows.length - 1) + " rows done.");
}

function grantDrivePermission() {
  DriveApp.getFolderById("1uxvFTunw21NblGfgLUB8eSzr_ybviQvC");
  Logger.log("Drive Permission Granted!");
}

function forceFullDrivePermission() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const testFile = folder.createFile("temp_permission_check.txt", "OK");
  testFile.setTrashed(true);
  Logger.log("Full Drive Access Granted Successfully!");
}
