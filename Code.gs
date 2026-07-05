/**
 * BD INITIATIVE REPORT — Backend
 * ------------------------------------------------
 * What this does:
 * 1. New BD Initiative → adds a new row in the chosen year's tab (e.g. "Towhid26")
 * 2. Follow-up Update → finds the exact existing row (you pick Customer + Buyer
 *    from a dropdown, so there's no ambiguity) and fills the next empty
 *    Week#/Recap slot (up to 4 follow-ups per initiative)
 * 3. Emails you a copy with the full workbook attached as .xlsx
 *
 * SETUP STEPS (one time only):
 * 1. Upload your BD workbook (e.g. Towhid_-BD_2026.xlsx) to Google Drive
 * 2. Right-click it > Open with > Google Sheets
 * 3. In that Sheet: Extensions > Apps Script
 * 4. Delete any starter code, paste this whole file in
 * 5. Update MY_EMAIL below
 * 6. Deploy > New deployment > Web app (Execute as: Me, Access: Anyone)
 * 7. Copy the Web app URL into bd-report.html's SCRIPT_URL
 */

// ============ CONFIG — EDIT THIS ============
const MY_EMAIL = "towhid@idealfastener.com"; // <-- your real email address
// ================================================

// Column layout (1-indexed) — matches your original BD workbook exactly
const COL = {
  segment: 1,
  initiativeDate: 2,
  customerName: 3,
  buyerName: 4,
  currentSupplier: 5,
  businessPotential: 6,
  zipperConsumption: 7,
  firstRecap: 8,
  nextSteps: 9,
  gamDate: 10,
  gamNameCountry: 11,
  buyingOfficeName: 12
};

// The 4 follow-up slots as [weekCol, recapCol] pairs
const FOLLOWUP_PAIRS = [[13, 14], [15, 16], [17, 18], [19, 20]];

function normalizeName(str) {
  if (!str) return "";
  return str.toString()
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

// ---------- GET: serves tab list + existing initiatives for dropdowns ----------
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = (e.parameter.action || "tabs");

  if (action === "tabs") {
    const names = ss.getSheets().map(s => s.getName());
    return jsonOut(names);
  }

  if (action === "initiatives") {
    const tabName = e.parameter.tab;
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return jsonOut([]);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonOut([]);
    const data = sheet.getRange(2, COL.customerName, lastRow - 1, COL.buyerName - COL.customerName + 1).getValues();
    const list = data
      .map((r, i) => ({ row: i + 2, customerName: (r[0] || "").toString().trim(), buyerName: (r[COL.buyerName - COL.customerName] || "").toString().trim() }))
      .filter(x => x.customerName !== "");
    return jsonOut(list);
  }

  return jsonOut({ error: "unknown action" });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- POST: new initiative or follow-up update ----------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(data.tab);
    if (!sheet) throw new Error('Tab "' + data.tab + '" not found.');

    if (data.action === "new") {
      addNewInitiative(sheet, data);
      sendEmail("New BD Initiative — " + data.customerName, buildNewInitiativeEmailBody(data));
    } else if (data.action === "followup") {
      addFollowup(sheet, data);
      sendEmail("BD Follow-up — " + data.customerName, buildFollowupEmailBody(data));
    } else {
      throw new Error("Unknown action: " + data.action);
    }

    return jsonOut({ result: "success" });
  } catch (err) {
    return jsonOut({ result: "error", message: err.toString() });
  }
}

function addNewInitiative(sheet, data) {
  const row = new Array(20).fill("");
  row[COL.segment - 1] = data.segment || "";
  row[COL.initiativeDate - 1] = data.initiativeDate || "";
  row[COL.customerName - 1] = data.customerName || "";
  row[COL.buyerName - 1] = data.buyerName || "";
  row[COL.currentSupplier - 1] = data.currentSupplier || "";
  row[COL.businessPotential - 1] = data.businessPotential || "";
  row[COL.zipperConsumption - 1] = data.zipperConsumption || "";
  row[COL.firstRecap - 1] = data.firstRecap || "";
  row[COL.nextSteps - 1] = data.nextSteps || "";
  row[COL.gamDate - 1] = data.gamDate || "";
  row[COL.gamNameCountry - 1] = data.gamNameCountry || "";
  row[COL.buyingOfficeName - 1] = data.buyingOfficeName || "";
  sheet.appendRow(row);
}

function addFollowup(sheet, data) {
  const targetRow = parseInt(data.row, 10);
  if (!targetRow || targetRow < 2) throw new Error("Invalid initiative selected.");

  // Verify the row still matches (safety check in case the sheet changed)
  const existingCustomer = normalizeName(sheet.getRange(targetRow, COL.customerName).getValue());
  if (existingCustomer !== normalizeName(data.customerName)) {
    throw new Error("Selected initiative no longer matches — please refresh and try again.");
  }

  let placed = false;
  for (let i = 0; i < FOLLOWUP_PAIRS.length; i++) {
    const [weekCol, recapCol] = FOLLOWUP_PAIRS[i];
    const existingRecap = sheet.getRange(targetRow, recapCol).getValue();
    if (!existingRecap || existingRecap.toString().trim() === "") {
      sheet.getRange(targetRow, weekCol).setValue(data.week || "");
      sheet.getRange(targetRow, recapCol).setValue(data.recap || "");
      placed = true;
      break;
    }
  }
  if (!placed) {
    throw new Error("All 4 follow-up slots are already filled for this initiative. Please update the sheet manually or extend the columns.");
  }
}

function buildNewInitiativeEmailBody(data) {
  return `
    <table style="width:100%; border-collapse: collapse;">
      <tr><td style="padding:6px 8px; font-weight:bold; width:170px; border-bottom:1px solid #eee;">Segment</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.segment)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Initiative Date</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.initiativeDate)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Customer Name</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.customerName)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Buyer Name</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.buyerName)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Current Supplier</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.currentSupplier)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Business Potential</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.businessPotential)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Zipper Consumption</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.zipperConsumption)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; vertical-align:top; border-bottom:1px solid #eee;">First Meeting Recap</td><td style="padding:6px 8px; border-bottom:1px solid #eee; white-space:pre-wrap;">${escapeHtml(data.firstRecap)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; vertical-align:top;">Next Steps</td><td style="padding:6px 8px; white-space:pre-wrap;">${escapeHtml(data.nextSteps)}</td></tr>
    </table>
  `;
}

function buildFollowupEmailBody(data) {
  return `
    <table style="width:100%; border-collapse: collapse;">
      <tr><td style="padding:6px 8px; font-weight:bold; width:170px; border-bottom:1px solid #eee;">Customer Name</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.customerName)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Buyer Name</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.buyerName)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Week#</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(data.week)}</td></tr>
      <tr><td style="padding:6px 8px; font-weight:bold; vertical-align:top;">Follow-up Recap</td><td style="padding:6px 8px; white-space:pre-wrap;">${escapeHtml(data.recap)}</td></tr>
    </table>
  `;
}

function sendEmail(subject, tableHtml) {
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: #4338CA; color: #fff; padding: 16px 20px; border-radius: 6px 6px 0 0;">
        <h2 style="margin:0; font-size: 18px;">${escapeHtml(subject)}</h2>
      </div>
      <div style="padding-top:14px;">${tableHtml}</div>
      <p style="color:#888; font-size:13px; margin-top:14px;">Full BD workbook (all tabs) attached as Excel.</p>
      <p style="color:#888; font-size:12px; margin-top:6px;">Submitted via mobile BD report tool · System designed & engineered by Towhid</p>
    </div>
  `;
  const excelBlob = getWorkbookAsExcelBlob();
  MailApp.sendEmail(MY_EMAIL, subject, "Please see attached report (also viewable in the email body).", {
    htmlBody: htmlBody,
    attachments: [excelBlob]
  });
}

function getWorkbookAsExcelBlob() {
  SpreadsheetApp.flush();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { "Authorization": "Bearer " + token } });
  const dateStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  return response.getBlob().setName(ss.getName() + " - " + dateStamp + ".xlsx");
}

function escapeHtml(str) {
  if (!str) return "";
  return str.toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
