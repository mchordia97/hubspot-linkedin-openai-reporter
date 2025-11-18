/**
 * Bright Data (LinkedIn) → "linkedin summary" column.
 */

function summarizeLinkedInProfilesToColumnT() {
  const API_TOKEN = getScriptProp('BRIGHTDATA_API_TOKEN');
  const DATASET_ID = getScriptProp('BRIGHTDATA_DATASET_ID');
  const API_BASE = 'https://api.brightdata.com/datasets/v3';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const urlsToProcess = [];
  const rowMap = {};

  // Start from row 2 (skip header)
  for (let i = 1; i < data.length; i++) {
    const url = data[i][10]; // Column K: Contact LinkedIn
    if (url && typeof url === 'string' && url.includes("linkedin.com/in/")) {
      urlsToProcess.push({ url });
      rowMap[url] = i + 1; // store sheet row (1-based)
    }
  }

  if (!urlsToProcess.length) {
    Logger.log("No valid LinkedIn URLs found.");
    return;
  }

  const snapshotId = triggerCollection_(API_BASE, DATASET_ID, API_TOKEN, urlsToProcess);
  if (!snapshotId) return;

  Logger.log("Waiting for Bright Data snapshot to be ready...");
  while (true) {
    Utilities.sleep(5000);
    const status = checkStatus_(API_BASE, API_TOKEN, snapshotId);
    Logger.log(`Status: ${status}`);
    if (status === "ready") break;
    if (status === "failed" || status === "error") {
      Logger.log("❌ Bright Data collection failed.");
      return;
    }
  }

  const profiles = getData_(API_BASE, API_TOKEN, snapshotId);
  if (!profiles || !profiles.length) {
    Logger.log("No profile data returned.");
    return;
  }

  profiles.forEach(profile => {
    const row = rowMap[profile.input_url];
    if (!row) return;

    const summary = buildLinkedInSummary_(profile);
    // Column T = 20th column (linkedin summary)
    sheet.getRange(row, 20).setValue(summary);
  });
}

function triggerCollection_(API_BASE, DATASET_ID, API_TOKEN, profiles) {
  try {
    const response = UrlFetchApp.fetch(
      `${API_BASE}/trigger?dataset_id=${DATASET_ID}`,
      {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: `Bearer ${API_TOKEN}` },
        payload: JSON.stringify(profiles)
      }
    );
    const result = JSON.parse(response.getContentText());
    return result.snapshot_id;
  } catch (e) {
    Logger.log("Error triggering Bright Data collection: " + e);
    return null;
  }
}

function checkStatus_(API_BASE, API_TOKEN, snapshotId) {
  try {
    const response = UrlFetchApp.fetch(
      `${API_BASE}/progress/${snapshotId}`,
      { headers: { Authorization: `Bearer ${API_TOKEN}` } }
    );
    const result = JSON.parse(response.getContentText());
    return result.status || "error";
  } catch (e) {
    Logger.log("Error checking Bright Data status: " + e);
    return "error";
  }
}

function getData_(API_BASE, API_TOKEN, snapshotId) {
  try {
    const response = UrlFetchApp.fetch(
      `${API_BASE}/snapshot/${snapshotId}?format=json`,
      { headers: { Authorization: `Bearer ${API_TOKEN}` } }
    );
    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log("Error fetching Bright Data snapshot: " + e);
    return null;
  }
}

function buildLinkedInSummary_(profile) {
  const name = profile.name || "N/A";
  const location = profile.city || profile.location || "N/A";
  const title = profile.position || (profile.current_company?.title || "N/A");
  const company = profile.current_company?.name || "N/A";
  const industry = profile.current_company?.industry || "N/A";
  const headline = profile.headline || "N/A";

  const educationList = (profile.education || [])
    .map(e => e.title)
    .filter(Boolean)
    .join("; ") || "N/A";

  const experienceEntries = [];
  const expData = profile.experience || [];

  expData.forEach(exp => {
    if (exp.positions && exp.positions.length) {
      exp.positions.forEach(pos => {
        const role = pos.title || "N/A";
        const comp = exp.company || "N/A";
        const start = pos.start_date || "N/A";
        const end = pos.end_date || "Present";
        const companyLocation = exp.location || "N/A";
        experienceEntries.push(
          `${role} at ${comp} — ${companyLocation} (${start} – ${end})`
        );
      });
    } else {
      const role = exp.title || "N/A";
      const comp = exp.company || "N/A";
      const start = exp.start_date || "N/A";
      const end = exp.end_date || "Present";
      const companyLocation = exp.location || "N/A";
      experienceEntries.push(
        `${role} at ${comp} — ${companyLocation} (${start} – ${end})`
      );
    }
  });

  const experience = experienceEntries.length
    ? experienceEntries.join(" | ")
    : "Experience not listed.";

  return (
    `Name: ${name}\n` +
    `Title: ${title}\n` +
    `Company: ${company}\n` +
    `Industry: ${industry}\n` +
    `Location: ${location}\n` +
    `Headline: ${headline}\n` +
    `Education: ${educationList}\n` +
    `Experience: ${experience}`
  );
}
