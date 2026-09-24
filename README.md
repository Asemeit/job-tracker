# Job Application Tracker

A simple Google Apps Script web app that helps you keep track of job and internship applications.

### Features
- Clean web interface to view and add applications
- Automatic logging of application confirmation emails from Gmail
- Google Sheet as the database (easy to view and export)
- Quick add via web form or bookmarklet

### How it works
1. Applications are stored in a Google Sheet
2. A web app provides a clean interface to manage them
3. A daily trigger checks your Gmail for confirmation emails and logs them automatically

### Setup
1. Create a Google Sheet with the following headers:  
   `id | date_applied | company | position | type | status | link | notes`
2. Open **Extensions → Apps Script** and paste the code from this repository
3. Deploy the project as a Web App
4. Set up a time-driven trigger for the `checkEmailsAndLog` function

### Tech
- Google Apps Script
- Google Sheets
- Gmail

---

Made for personal use.
