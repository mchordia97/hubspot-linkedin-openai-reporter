/**
 * HubSpot → Google Sheets task ingestion.
 */

function fetchAllHubSpotTasks() {
  const HUBSPOT_TOKEN = getScriptProp('HUBSPOT_TOKEN');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Clear existing data and write header
  sheet.clear();
  sheet.appendRow([
    "Record ID", "Activity type", "Activity date", "Task Title", "Details",
    "Activity assigned to", "Associated Contacts", "Associated Companies",
    "Task Status", "Contact Email", "Contact LinkedIn", "Contact Phone", "Mobile Phone",
    "Job Title", "Company Name", "Industry", "Company Description", "Company LinkedIn",
    "Employees", "linkedin summary", "summary"
  ]);

  let after = null;
  let hasMore = true;

  while (hasMore) {
    const endpoint = 'https://api.hubapi.com/crm/v3/objects/tasks/search';

    const payload = {
      filterGroups: [{
        filters: [
          {
            propertyName: "hs_task_status",
            operator: "IN",
            values: ["NOT_STARTED", "IN_PROGRESS", "WAITING", "DEFERRED"]
          },
          {
            propertyName: "hubspot_owner_id",
            operator: "IN",
            // Replace with the HubSpot owner IDs you care about
            values: ["80252376", "77457823"]
          }
        ]
      }],
      properties: [
        "hs_task_type",
        "hs_task_status",
        "hubspot_owner_id",
        "hs_timestamp",
        "hs_task_body",
        "hs_task_subject"
      ],
      limit: 100,
      after: after
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${HUBSPOT_TOKEN}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(endpoint, options);
    const data = JSON.parse(response.getContentText());
    const tasks = data.results || [];

    tasks.forEach(task => {
      const taskId = task.id;
      const props = task.properties || {};
      const ownerId = props.hubspot_owner_id || "";

      const ownerName = getOwnerName_(HUBSPOT_TOKEN, ownerId);
      const contactIds = getAssociatedEntities_(HUBSPOT_TOKEN, taskId, "contacts");
      const companyIds = getAssociatedEntities_(HUBSPOT_TOKEN, taskId, "companies");

      const contactInfo = contactIds.length
        ? getContactDetails_(HUBSPOT_TOKEN, contactIds[0])
        : {};

      const companyInfo = contactInfo.companyId
        ? getCompanyDetails_(HUBSPOT_TOKEN, contactInfo.companyId)
        : (companyIds.length
          ? getCompanyDetails_(HUBSPOT_TOKEN, companyIds[0])
          : {});

      sheet.appendRow([
        taskId,
        props.hs_task_type || "",
        props.hs_timestamp || "",
        props.hs_task_subject || "",
        props.hs_task_body || "",
        ownerName,
        `${contactInfo.firstname || ''} ${contactInfo.lastname || ''}`.trim(),
        companyInfo.name || "",
        props.hs_task_status || "",
        contactInfo.email || "",
        contactInfo.linkedinurl || contactInfo.linkedin || "",
        contactInfo.phone || "",
        contactInfo.mobile || "",
        contactInfo.jobtitle || "",
        companyInfo.name || "",
        companyInfo.industry || "",
        companyInfo.description || "",
        companyInfo.linkedin || "",
        companyInfo.employees || "",
        "",   // linkedin summary (to be filled later)
        ""    // OpenAI summary (to be filled later)
      ]);
    });

    after = data.paging && data.paging.next ? data.paging.next.after : null;
    hasMore = !!after;
  }
}

function getOwnerName_(HUBSPOT_TOKEN, ownerId) {
  if (!ownerId) return "";
  const url = `https://api.hubapi.com/crm/v3/owners/${ownerId}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${HUBSPOT_TOKEN}` },
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  const first = data.firstName || "";
  const last = data.lastName || "";
  const full = `${first} ${last}`.trim();
  return full || ownerId;
}

function getAssociatedEntities_(HUBSPOT_TOKEN, objectId, toObjectType) {
  const url = `https://api.hubapi.com/crm/v3/objects/tasks/${objectId}/associations/${toObjectType}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${HUBSPOT_TOKEN}` },
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  return (data.results || []).map(r => r.id);
}

function getContactDetails_(HUBSPOT_TOKEN, contactId) {
  if (!contactId) return {};

  const url =
    `https://api.hubapi.com/crm/v3/objects/contacts/${contactId}` +
    `?properties=firstname,lastname,email,phone,mobilephone,jobtitle,linkedinbio,linkedin_account,linkedinurl`;

  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${HUBSPOT_TOKEN}` },
    muteHttpExceptions: true
  });

  const response = JSON.parse(res.getContentText());
  const props = response.properties || {};

  let companyId = null;
  try {
    const assocRes = UrlFetchApp.fetch(
      `https://api.hubapi.com/crm/v3/objects/contacts/${contactId}/associations/companies`,
      { headers: { Authorization: `Bearer ${HUBSPOT_TOKEN}` }, muteHttpExceptions: true }
    );
    const assocData = JSON.parse(assocRes.getContentText());
    if (assocData.results && assocData.results.length > 0) {
      companyId = assocData.results[0].id;
    }
  } catch (e) {
    Logger.log("Error fetching contact-company association: " + e);
  }

  return {
    firstname: props.firstname || "",
    lastname: props.lastname || "",
    email: props.email || "",
    phone: props.phone || "",
    mobile: props.mobilephone || "",
    jobtitle: props.jobtitle || "",
    linkedin: props.linkedinbio || "",
    linkedinurl: props["linkedin_account"] || props.linkedinurl || "",
    companyId: companyId
  };
}

function getCompanyDetails_(HUBSPOT_TOKEN, companyId) {
  if (!companyId) return {};
  const url =
    `https://api.hubapi.com/crm/v3/objects/companies/${companyId}` +
    `?properties=name,industry,description,linkedin_company_page,numberofemployees`;

  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${HUBSPOT_TOKEN}` },
    muteHttpExceptions: true
  });

  const props = JSON.parse(res.getContentText()).properties || {};

  return {
    name: props.name || "",
    industry: props.industry || "",
    description: props.description || "",
    linkedin: props.linkedin_company_page || "",
    employees: props.numberofemployees || ""
  };
}
