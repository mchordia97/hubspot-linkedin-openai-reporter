/**
 * Build and email an HTML summary of contacts and AI summaries.
 *
 * Expects a sheet (e.g., "Sheet2") with columns including:
 * - Activity type
 * - Activity assigned to
 * - Associated Contacts
 * - Task Title
 * - summary
 * - Contact LinkedIn
 */

function createAndEmailContactSummaryDocSimple() {
  const sheetName = "Sheet2"; // change to your actual sheet name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found.`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (!data.length) {
    Logger.log("No data in summary sheet.");
    return;
  }

  const headers = data[0];

  const idx = {
    type: headers.indexOf("Activity type"),
    assignedTo: headers.indexOf("Activity assigned to"),
    contact: headers.indexOf("Associated Contacts"),
    task_title: headers.indexOf("Task Title"),
    summary: headers.indexOf("summary"),
    linkedin_url: headers.indexOf("Contact LinkedIn")
  };

  const requiredCols = ["type", "assignedTo", "contact", "task_title", "summary", "linkedin_url"];
  for (const key of requiredCols) {
    if (idx[key] === -1) {
      Logger.log(`Missing required column: ${key}`);
      return;
    }
  }

  const seenContacts = new Set();
  const groupedData = {};

  // Start from row 2 (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const type = row[idx.type];
    const assignedTo = row[idx.assignedTo];
    const contact = row[idx.contact];
    const summary = row[idx.summary];
    const taskTitle = row[idx.task_title];
    const linkedinUrl = row[idx.linkedin_url];

    if (type !== "CALL" || !contact) continue;
    if (seenContacts.has(contact)) continue;
    seenContacts.add(contact);

    const key = `${assignedTo}|||${type}`;
    if (!groupedData[key]) groupedData[key] = [];
    groupedData[key].push({
      contact,
      summary,
      taskTitle,
      linkedinUrl
    });
  }

  let htmlContent = "<h1>🗂️ HubSpot Contact Summary</h1>";

  Object.keys(groupedData).forEach(key => {
    const [assignedTo, type] = key.split("|||");
    htmlContent += `<h2>Assigned to: ${assignedTo}</h2>`;
    htmlContent += `<h3>Activity type: ${type}</h3>`;

    groupedData[key].forEach(entry => {
      const contact = entry.contact || "(No name)";
      const title = entry.taskTitle || "(No task title)";
      const url = entry.linkedinUrl || "(No LinkedIn URL)";
      const summary = entry.summary || "(No summary available)";

      htmlContent += `
        <p>
          <strong>• ${contact}</strong><br>
          <span style="margin-left: 20px;"><strong>Task Title:</strong> ${title}</span><br>
          <span style="margin-left: 20px;"><strong>LinkedIn URL:</strong> ${url}</span><br>
          <span style="margin-left: 20px;">${summary}</span>
        </p>
      `;
    });

    htmlContent += `<br>`;
  });

  // Adjust recipients as needed
  const recipients = [
    "chordiamanali@gmail.com"
  ].join(",");

  MailApp.sendEmail({
    to: recipients,
    subject: "HubSpot Contact Summary Report",
    htmlBody: htmlContent
  });

  Logger.log("✅ Email with contact summary sent successfully.");
}
