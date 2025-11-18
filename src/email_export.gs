/**
 * Export current spreadsheet as Excel and email it.
 */

function sendSpreadsheetAsExcelDaily() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Change this to your distribution list
  const emails = "chordiamanali@gmail.com";
  const subject = "HubSpot Report - " + spreadsheet.getName();
  const body = "Please find the attached HubSpot task report.";

  // Export the entire spreadsheet as Excel
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=xlsx`;

  const options = {
    method: "get",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const blob = response.getBlob().setName(spreadsheet.getName() + ".xlsx");

  MailApp.sendEmail({
    to: emails,
    subject: subject,
    body: body,
    attachments: [blob]
  });

  Logger.log("✅ Spreadsheet exported and emailed as Excel.");
}
