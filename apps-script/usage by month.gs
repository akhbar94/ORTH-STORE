/*************************************************
 * FAST & SAFE GOOGLE SHEETS SCRIPT
 * Fixes "Working..." performance issue
 *************************************************/

/* ========= ON EDIT (LIGHTWEIGHT ROUTER) ========= */
function onEdit(e) {
  if (!e) return;

  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const ss = e.source;

  // 1️⃣ Auto-date column H when column C is edited
  if (col === 3 && isSheetListedInSummary(ss, name)) {
    updateDateColumnH(sheet, row);
    return;
  }

  // 2️⃣ Copy ALL → Record based on SUMMARY!C2
  if (name === "Record" && row === 1) {
    copyRecordBySummaryDate();
    return;
  }

  // 3️⃣ SUMMARY checkbox logic
  if (name === "SUMMARY" && col === 16 && row >= 18 && row <= 30) {
    handleSummaryCheckbox(sheet, row);
  }
}

/* ========= CHECK SHEET NAME AGAINST SUMMARY ========= */
function isSheetListedInSummary(ss, sheetName) {
  const summary = ss.getSheetByName("SUMMARY");
  if (!summary) return false;

  const names = summary
    .getRange("B50:B60")
    .getValues()
    .flat()
    .filter(n => n !== "");

  return names.includes(sheetName);
}

/* ========= 1️⃣ AUTO DATE FUNCTION ========= */
function updateDateColumnH(sheet, row) {
  sheet.getRange(row, 8).setValue(
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    )
  );
}



/* ========= 2️⃣ COPY DATA INTO RECORD ========= */
function copyRecordBySummaryDate() {
  const ss = SpreadsheetApp.getActive();
  const summary = ss.getSheetByName("SUMMARY");
  const record = ss.getSheetByName("Record");
  const all = ss.getSheetByName("ALL");

  if (!summary || !record || !all) return;

  const targetDate = summary.getRange("C2").getValue();
  if (!(targetDate instanceof Date)) return;

  targetDate.setHours(0, 0, 0, 0);

  const headers = record
    .getRange(1, 1, 1, record.getLastColumn())
    .getValues()[0];

  const colIndex = headers.findIndex(h => {
    if (!(h instanceof Date)) return false;
    h.setHours(0, 0, 0, 0);
    return h.getTime() === targetDate.getTime();
  });

  if (colIndex === -1) return;

  const values = all.getRange("C2:C38").getValues();
  record
    .getRange(2, colIndex + 1, values.length, 1)
    .setValues(values);
}

/* ========= 3️⃣ SUMMARY CHECKBOX HANDLER ========= */
function handleSummaryCheckbox(sheet, row) {
  const jCell = sheet.getRange(row, 10);
  const value = jCell.getValue();

  if (!value) return;

  // Reset checkbox
  sheet.getRange(row, 16).setValue(false);
  jCell.clearContent();

  const destRow = 33;
  const c2 = sheet.getRange("C2").getValue();

  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(destRow, 10, lastRow - destRow + 1, 7);
  const data = range.getValues();

  // Shift data down
  data.unshift([value, "", "", "", "", "", c2]);
  data.pop();

  range.setValues(data);
}

/* ========= ON OPEN (LIGHT ONLY) ========= */
function onOpen() {
  SpreadsheetApp
    .getUi()
    .toast("✅ Sheet ready (optimized)");
}

/* ========= MANUAL AUTO-COPY (OPTIONAL BUTTON) ========= */
function runAutoCopy() {
  const ss = SpreadsheetApp.getActive();
  const record = ss.getSheetByName("Record");
  const all = ss.getSheetByName("ALL");

  if (!record || !all) return;

  const tz = ss.getSpreadsheetTimeZone();
  const today = new Date(
    Utilities.formatDate(new Date(), tz, "yyyy-MM-dd")
  );

  const headers = record
    .getRange(1, 1, 1, record.getLastColumn())
    .getValues()[0];

  const data = all.getRange("C2:C40").getValues();
  let count = 0;

  headers.forEach((h, i) => {
    if (!(h instanceof Date)) return;

    h.setHours(0, 0, 0, 0);
    const diff = Math.floor((h - today) / 86400000);

    if (diff >= -2 && diff <= 2) {
      const cell = record.getRange(2, i + 1);
      if (!cell.getValue()) {
        record
          .getRange(2, i + 1, data.length, 1)
          .setValues(data);
        count++;
      }
    }
  });

  ss.toast(`Auto-copy completed: ${count} column(s)`);
}
