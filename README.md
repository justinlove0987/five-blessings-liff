# Apps Script Backend

This folder stores the Google Apps Script backend source code for the Five Blessings LIFF MVP.

Important:
- This folder is for source control and AI agent editing.
- It is not automatically deployed to Google Apps Script.
- After editing Code.js, manually copy the updated code into Google Apps Script.
- Then redeploy the Apps Script Web App with a new version.

Current architecture:
- Frontend: GitHub Pages / index.html
- Login: LINE LIFF
- Backend: Google Apps Script Web App
- Database: Google Sheet

MVP rule:
Keep changes simple. Do not add complex auth, build tools, frameworks, or CI/CD yet.
