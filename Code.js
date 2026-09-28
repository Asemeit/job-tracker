// ===== CONFIG =====
const EMAIL = "pasemeit@gmail.com";

// Strictly target application SUBMISSION confirmations
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
  "we received your application"
];

// Keywords that indicate it's NOT a completed application
const EXCLUDE_KEYWORDS = [
  "one-time passcode",
  "verification code",
  "passcode",
  "draft",
  "incomplete application",
  "complete your application",
  "action required",
  "finish your application"
];

// Portal domains to ignore when guessing company name from sender email
const PORTAL_DOMAINS = [
  "greenhouse", "workday", "successfactors", "lever", 
  "smartrecruiters", "myworkday", "ashbyhq", "gmail", "outlook", "us"
];

// ===== WEB APP =====
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Job Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getApplications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).toLowerCase().trim());
  
  // Dynamic header column matching
  const idIdx = headers.indexOf("id");
  const dateIdx = headers.indexOf("date_applied") !== -1 ? headers.indexOf("date_applied") : headers.indexOf("date");
  const companyIdx = headers.indexOf("company");
  const positionIdx = headers.indexOf("position");
  const typeIdx = headers.indexOf("type");
  const statusIdx = headers.indexOf("status");
  const linkIdx = headers.indexOf("link");
  const notesIdx = headers.indexOf("notes");

  const rows = data.slice(1).filter(r => r[0] !== "" && r[0] !== null && r[0] !== undefined);

  const result = rows.map(row => {
    let rawDate = row[dateIdx];
    let formattedDate = rawDate;
    if (rawDate instanceof Date) {
      formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    return {
      id: row[idIdx] !== undefined ? row[idIdx] : "",
      date_applied: formattedDate || "",
      company: row[companyIdx] || "Unknown",
      position: row[positionIdx] || "N/A",
      type: row[typeIdx] || "Job",
      status: row[statusIdx] || "Applied",
      link: row[linkIdx] || "",
      notes: row[notesIdx] || ""
    };
  }).reverse();

  return result;
}

function getNextId(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 1;
  const ids = data.slice(1).map(r => Number(r[0])).filter(id => !isNaN(id) && id > 0);
  return ids.length > 0 ? Math.max(...ids) + 1 : 1;
}

function addApplication(app) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const newId = getNextId(sheet);

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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const existing = data.slice(1).map(r => (String(r[2]) + " " + String(r[3])).toLowerCase());

  const query = "newer_than:10d (" + KEYWORDS.map(k => `"${k}"`).join(" OR ") + ")";
  const threads = GmailApp.search(query, 0, 40);

  threads.forEach(thread => {
    const msg = thread.getMessages().pop();
    const subject = msg.getSubject() || "";
    const body = msg.getPlainBody().substring(0, 1500) || "";
    const fullText = (subject + " " + body).toLowerCase();
    const date = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // 1. Exclude one-time passcodes and incomplete draft reminders
    if (EXCLUDE_KEYWORDS.some(ex => fullText.includes(ex))) return;

    // 2. Validate keyword match
    const matchedKeyword = KEYWORDS.find(k => fullText.includes(k.toLowerCase()));
    if (!matchedKeyword) return;

    // 3. Extract Company Name (Subject line first, then domain)
    let company = "";
    const subMatch = subject.match(/(?:at|to|with)\s+([A-Z0-9\s&]+)(?:$|\s+for|\!|\.)/i);
    if (subMatch && subMatch[1] && subMatch[1].trim().length > 1) {
      company = subMatch[1].trim();
    }

    if (!company) {
      const from = msg.getFrom() || "";
      const domainMatch = from.match(/@([a-z0-9.-]+)/i);
      if (domainMatch) {
        let domainParts = domainMatch[1].toLowerCase().split('.');
        let mainDomain = domainParts.length > 2 ? domainParts[domainParts.length - 2] : domainParts[0];
        if (!PORTAL_DOMAINS.includes(mainDomain)) {
          company = mainDomain.charAt(0).toUpperCase() + mainDomain.slice(1);
        }
      }
    }
    if (!company) company = "Unknown Company";

    // 4. Extract Position Title
    let position = "";
    let posMatch = body.match(/(?:application for|position of|role:?)\s+([A-Za-z0-9\s\-\/]{4,40})/i);
    if (posMatch && posMatch[1]) {
      position = posMatch[1].split("\n")[0].trim();
    }

    if (!position || position.length < 4) {
      position = subject
        .replace(/thank you for (your )?application/gi, "")
        .replace(/thanks for applying( to)?/gi, "")
        .replace(/we received your .* application/gi, "")
        .replace(/application for/gi, "")
        .replace(/re:|fw:|fwd:/gi, "")
        .replace(/\[.*?\]/g, "")
        .trim();
    }

    if (!position || position.length < 3 || position.toLowerCase().startsWith("thank you")) {
      position = "Application Submitted";
    }

    // 5. Prevent Duplicates
    if (existing.some(e => e.includes(company.toLowerCase()) && e.includes(position.toLowerCase().substring(0, 15)))) {
      return;
    }

    const newId = getNextId(sheet);
    const type = (position.toLowerCase().includes("intern") || fullText.includes("intern")) ? "Internship" : "Job";

    sheet.appendRow([
      newId,
      date,
      company,
      position,
      type,
      "Applied",
      "",
      "Auto-logged from email: " + subject.substring(0, 70)
    ]);

    existing.push((company + " " + position).toLowerCase());
  });
}