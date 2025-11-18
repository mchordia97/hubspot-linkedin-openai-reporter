/**
 * OpenAI GPT-4o → Business intelligence summary into "summary" column.
 */

function summarizeRowOpenAI(rowIndex) {
  const OPENAI_API_KEY = getScriptProp('OPENAI_API_KEY');
  const OPENAI_ENDPOINT = 'https://api.openai.com/v1/chat/completions';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (rowIndex < 2) {
    Logger.log("Invalid row index (must be >= 2).");
    return;
  }

  // Columns G (7) to U (21) - adjust if your layout changes
  const rowRange = sheet.getRange(rowIndex, 7, 1, 15);
  const rowValues = rowRange.getValues()[0];
  const rowText = rowValues.join(", ");

  const prompt =
    `Write a concise and professional business intelligence summary in paragraph form based on the following contact data extracted from columns G to U of a Google Sheet. Mention the Task Title. Include the person's full name, current job title and company, current work location (city, state), and whether they are based at a headquarters or field location. If headquarters, specify the geography or business unit they likely oversee. If field, use the company name to identify the types of assets present in that location (e.g., pipelines, wells, terminals, gas processing plants). Briefly describe the company’s core business, especially if related to oil, gas, or midstream operations. Include previous roles or work locations if available. Estimate years of experience if inferable. Based on the role and experience, describe which facilities or assets the person is likely responsible for. Return the response in fluent paragraph form only, without bullet points or headers.\n\n` +
    rowText;

  const payload = {
    model: "gpt-4o",
    messages: [{ role: "user", content: prompt }],
    temperature: 0.5,
    max_tokens: 300
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(OPENAI_ENDPOINT, options);
    const json = JSON.parse(response.getContentText());
    const summary = json.choices?.[0]?.message?.content || "No summary returned";

    // Column U (21) = final AI summary
    sheet.getRange(rowIndex, 21).setValue(summary);
    Logger.log(`Summary written to row ${rowIndex}, column 21`);
  } catch (e) {
    Logger.log("OpenAI request failed: " + e);
    sheet.getRange(rowIndex, 21).setValue("API error");
  }
}

/**
 * Run summaries for all rows in the active sheet.
 */
function runRowSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    summarizeRowOpenAI(rowIndex);
    Utilities.sleep(2000); // simple rate limit spacing
  }
}
