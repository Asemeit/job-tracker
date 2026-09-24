// ===== CONFIG =====
const SHEET_NAME = "Sheet1"; // change only if you renamed the sheet tab
const EMAIL = "pasemeit@gmail.com";

// Phrases that usually appear in application confirmation emails
const KEYWORDS = [
  "thank you for applying",
  "application received",
  "we have received your application",
  "your application has been submitted",
  "application confirmation",
  "thanks for your application"
];

// ===== WEB APP =====
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Job Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getApplications() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).filter(r => r[0] !== "");

  return rows.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse(); // newest first
}

function addApplication(app) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const newId = lastRow <= 1 ? 1 : Number(sheet.getRange(lastRow, 1).getValue()) + 1;

  sheet.appendRow([
    newId,
    app.date_applied || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    app.company,
    app.position,
    app.type || "Job",
    app.status || "Applied",
    app.link || "",
    app.notes || ""
  ]);

  return { success: true, id: newId };
}

function updateStatus(id, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue(newStatus);
      return { success: true };
    }
  }
  return { success: false };
}

// ===== AUTO EMAIL LOGGER =====
function checkEmailsAndLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const existing = sheet.getDataRange().getValues().slice(1).map(r => (r[2] + " " + r[3]).toLowerCase());

  const query = "newer_than:7d (" + KEYWORDS.map(k => `subject:"${k}"`).join(" OR ") + ")";
  const threads = GmailApp.search(query, 0, 25);

  threads.forEach(thread => {
    const msg = thread.getMessages()[0];
    const subject = msg.getSubject();
    const date = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    let company = "Unknown Company";
    let position = subject
      .replace(/thank you for applying.*/i, "")
      .replace(/application received.*/i, "")
      .replace(/your application.*/i, "")
      .replace(/we have received.*/i, "")
      .trim();

    const key = (company + " " + position).toLowerCase();
    if (existing.includes(key) || existing.includes(position.toLowerCase())) return;

    const lastRow = sheet.getLastRow();
    const newId = lastRow <= 1 ? 1 : Number(sheet.getRange(lastRow, 1).getValue()) + 1;

    sheet.appendRow([
      newId,
      date,
      company,
      position || "Application",
      "Job",
      "Applied",
      "",
      "Auto-logged from email: " + subject
    ]);
  });
}