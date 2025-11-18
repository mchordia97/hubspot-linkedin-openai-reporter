HubSpot → LinkedIn → OpenAI Automated Sales Intelligence Pipeline

This project automates the full sales intelligence workflow by integrating HubSpot CRM, LinkedIn (via Bright Data), OpenAI GPT-4o, Google Sheets, and Gmail. It retrieves HubSpot tasks, enriches contacts with LinkedIn profile data, generates AI-powered business intelligence summaries, and emails daily Excel and HTML reports to sales leadership.

Overview
The system performs five automated steps:
1. Extract HubSpot tasks using the CRM API.
2. Enrich contacts with LinkedIn profile data via Bright Data.
3. Generate AI business summaries using OpenAI GPT-4o.
4. Update Google Sheets with enriched CRM & LinkedIn data.
5. Email Excel & HTML reports to the sales team.

Features

1. HubSpot Task Ingestion
- Fetches tasks where status is NOT_STARTED, IN_PROGRESS, WAITING, or DEFERRED.
- Retrieves associated contacts, companies, and owners.
- Pulls job title, phone, email, LinkedIn, industry, employee size.
- Writes structured rows into Sheets.

2. LinkedIn Profile Enrichment
- Detects LinkedIn URLs.
- Sends them to Bright Data.
- Fetches profile insights.
- Writes a LinkedIn summary per contact.

3. AI-Powered Business Intelligence (GPT-4o)
- Produces professional BI paragraphs on responsibilities, assets, experience, and company background.

4. Daily Excel Report
- Exports the Google Sheet to Excel.
- Emails to recipients.

5. HTML Sales Intelligence Digest
- Groups contacts by owner/activity.
- Generates HTML summaries with LinkedIn + AI insights.

Architecture:
HubSpot → Google Sheets → LinkedIn (Bright Data) → OpenAI GPT-4o
      ↓                                        ↓
 Excel Export                          HTML Summary Email

Tech Stack:
Google Apps Script, HubSpot API, Bright Data, OpenAI, Gmail, Sheets.

Security:
API keys stored in Script Properties:
HUBSPOT_TOKEN, BRIGHTDATA_API_TOKEN, BRIGHTDATA_DATASET_ID, OPENAI_API_KEY.

Business Impact:
- Reduces SDR research time by ~90%
- Improves outreach personalization
- Provides leadership-ready intelligence automatically

