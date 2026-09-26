// ===== CONFIG =====
const SHEET_NAME = "Sheet1";
const EMAIL = "pasemeit@gmail.com";

// ===== KEYWORDS =====
const KEYWORDS = [
  "thank you for applying",
  "thank you for taking the time to submit your application",
  "thank you for submitting your application",
  "thank you for your application",
  "application received",
  "we have received your application",
  "your application has been submitted",
  "application confirmation",
  "thanks for your application",
  "thanks for applying",
  "submit your application",
  "we received your application",
  "application for"
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

// ===== EMAIL LOGGER =====
function checkEmailsAndLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const existing = data.slice(1).map(r => (String(r[2]) + " " + String(r[3])).toLowerCase());

  Logger.log("Existing entries: " + existing.length);

  const query = "newer_than:10d (" + KEYWORDS.map(k => `"${k}"`).join(" OR ") + ")";
  const threads = GmailApp.search(query, 0, 40);
  Logger.log("Threads found: " + threads.length);

  threads.forEach((thread, index) => {
    const msg = thread.getMessages().pop();
    const subject = msg.getSubject() || "";
    const body = msg.getPlainBody().substring(0, 1200) || "";
    const fullText = (subject + " " + body).toLowerCase();
    const date = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    Logger.log("----- Email " + (index + 1) + " -----");
    Logger.log("Subject: " + subject);

    const matchedKeyword = KEYWORDS.find(k => fullText.includes(k.toLowerCase()));
    if (!matchedKeyword) {
      Logger.log("SKIPPED: No keyword matched");
      return;
    }

    // Company detection
    let company = "Unknown Company";
    if (fullText.includes("microsoft")) company = "Microsoft";
    else if (fullText.includes("google")) company = "Google";
    else if (fullText.includes("amazon") || fullText.includes("aws")) company = "Amazon";
    else if (fullText.includes("meta") || fullText.includes("facebook")) company = "Meta";
    else if (fullText.includes("apple")) company = "Apple";
    else if (fullText.includes("netflix")) company = "Netflix";
    else if (fullText.includes("stripe")) company = "Stripe";
    else if (fullText.includes("aicines")) company = "AICines";
    else {
      const from = msg.getFrom() || "";
      const match = from.match(/@([a-z0-9.-]+)/i);
      if (match) {
        company = match[1].split('.')[0];
        company = company.charAt(0).toUpperCase() + company.slice(1);
      }
    }
    Logger.log("Company: " + company);

    // Position extraction (improved for Microsoft and others)
    let position = "";

    // Microsoft style
    let match = body.match(/application for\s+(.+?)(?:\s*\(Job number|\.\s+We’re glad|\.\s+You may not|Thank you,)/i);
    if (match && match[1]) {
      position = match[1].trim();
    }

    // Fallback cleaning of subject
    if (!position || position.length < 8) {
      position = subject
        .replace(/thank you for (your )?application!?/gi, "")
        .replace(/thanks for applying( to)?/gi, "")
        .replace(/we received your .* application/gi, "")
        .replace(/application for/gi, "")
        .replace(/re:|fw:|fwd:/gi, "")
        .trim();
    }

    if (!position || position.length < 5) position = "Application";
    Logger.log("Position: " + position);

    // Duplicate check
    const key = (company + " " + position).toLowerCase();
    if (existing.some(e => e.includes(company.toLowerCase()) && e.includes(position.toLowerCase().substring(0, 18)))) {
      Logger.log("SKIPPED: Already exists");
      return;
    }

    // Add to sheet
    const lastRow = sheet.getLastRow();
    const newId = lastRow <= 1 ? 1 : Number(sheet.getRange(lastRow, 1).getValue()) + 1;

    const type = (position.toLowerCase().includes("intern") || fullText.includes("intern")) ? "Internship" : "Job";

    sheet.appendRow([
      newId,
      date,
      company,
      position,
      type,
      "Applied",
      "",
      "Auto-logged: " + subject.substring(0, 70)
    ]);

    existing.push(key);
    Logger.log("ADDED: " + company + " → " + position);
  });

  Logger.log("Finished.");
}