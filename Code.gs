/**
 * GJIMT Registered Admissions - Google Sheets Backend
 *
 * This web app returns centrally stored programme counts.
 * Update the numbers in the Google Sheet when registrations are verified.
 *
 * Sheet layout:
 * A1: Programme | B1: Registered Admissions | C1: Last Updated
 * A2: BBA
 * A3: BCA
 * A4: B.Com (Hons.)
 * A5: MBA
 * A6: MCA
 */

// Paste the Google Sheet ID here.
// Example URL:
// https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit
const SPREADSHEET_ID = "PASTE_GOOGLE_SHEET_ID_HERE";
const SHEET_NAME = "Admissions";

const PROGRAMMES = [
  { id: "bba",  label: "BBA",           startingValue: 20 },
  { id: "bca",  label: "BCA",           startingValue: 20 },
  { id: "bcom", label: "B.Com (Hons.)", startingValue: 10 },
  { id: "mba",  label: "MBA",           startingValue: 20 },
  { id: "mca",  label: "MCA",           startingValue: 10 }
];

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || "get");

    if (action !== "get") {
      return jsonResponse({
        success: false,
        message: "Unsupported action."
      });
    }

    const result = getAdmissionCounts();

    return jsonResponse({
      success: true,
      counts: result.counts,
      updatedAt: result.updatedAt
    });
  } catch (error) {
    return jsonResponse({
      success: false,
      message: error.message
    });
  }
}

/**
 * Run this function once from the Apps Script editor.
 * It creates the Admissions sheet and initial rows if required.
 */
function setupAdmissionsSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  sheet.clear();
  sheet.getRange(1, 1, 1, 3).setValues([
    ["Programme", "Registered Admissions", "Last Updated"]
  ]);

  const now = new Date();
  const rows = PROGRAMMES.map(programme => [
    programme.label,
    programme.startingValue,
    now
  ]);

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  sheet.autoResizeColumns(1, 3);
}

/**
 * Reads the current verified values from the central Google Sheet.
 */
function getAdmissionCounts() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Sheet "' + SHEET_NAME + '" was not found. Run setupAdmissionsSheet() first.'
    );
  }

  const values = sheet
    .getRange(2, 1, PROGRAMMES.length, 3)
    .getValues();

  const counts = {};
  let latestUpdated = null;

  PROGRAMMES.forEach((programme, index) => {
    const row = values[index];
    const count = Number(row[1]);

    counts[programme.id] =
      Number.isFinite(count) && count >= 0 ? Math.floor(count) : 0;

    const rowDate = row[2];
    if (rowDate instanceof Date && !isNaN(rowDate.getTime())) {
      if (!latestUpdated || rowDate > latestUpdated) {
        latestUpdated = rowDate;
      }
    }
  });

  return {
    counts: counts,
    updatedAt: Utilities.formatDate(
      latestUpdated || new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy, hh:mm:ss a"
    )
  };
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
