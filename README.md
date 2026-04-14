Bluebeam Studio — Dynamic Project Review PoC
Disclaimer
This repository contains a proof-of-concept integration showing how an external system can drive a Bluebeam Studio review workflow across Studio Projects and Studio Sessions using the Bluebeam API.

This code is a reference implementation for evaluation and development use only, and it is not intended for production use.

What This PoC Does
This server creates a new Studio Project for each run, uploads review documents into that project, checks those files out into a Studio Session for markup, then checks them back in and extracts markup data for downstream processing.

It also handles project permissions, folder permissions, reviewer permissions, webhook subscription, snapshot generation, and PDF state-model injection before upload.

Workflow
Step	Endpoint	What it does
0a	POST /poc/setup-project	Creates a new restricted Studio Project, creates standard folders, and uploads custom-columns.xml.
0b	POST /poc/upload-to-project	Uploads PDF files into the project, optionally injecting the built-in 5-step QC state model first.
0c	POST /poc/apply-custom-columns	Applies custom columns to uploaded project files.
1	POST /poc/trigger	Simulates the source-system workflow event and logs the active files/reviewers.
2	POST /poc/create-session	Creates a restricted Studio Session and applies default session permissions.
3	POST /poc/register-webhook	Registers a session webhook subscription unless the callback URL is localhost.
4	POST /poc/checkout-to-session	Checks uploaded project files out into the Studio Session.
5	POST /poc/invite-reviewers	Invites reviewers to the session and project, retrieves project user IDs, and applies folder/user permissions.
6	(manual)	Reviewers open the Session in Revu and add markups.
7	POST /poc/checkin	Checks session files back into the project using the documented check-in pattern.
8	POST /poc/export-markups	Runs exportmarkups to generate XML markup exports in the project.
9	POST /poc/run-markuplist-job	Downloads exported XML and extracts normalized markup data into memory.
9b	POST /poc/downstream-process	Combined downstream step: check-in, export markups, and extract markup data.
10	POST /poc/finalize	Finalizes the Studio Session.
11	POST /poc/snapshot	Creates snapshots of reviewed session files and saves PDFs into public/.
12	POST /poc/cleanup	Deletes the webhook subscription and deletes the session.
Architecture
text
Source System (stubbed)
        │
        │ workflow event
        ▼
Express Backend (server.js)
├── OAuth token management (tokenManager.js)
├── Dynamic project creation
├── Folder + user permission setup
├── PDF state model injection (pdf-lib)
├── Project upload / session checkout / check-in flow
├── Markup export + XML parsing
├── Webhook receiver
└── Static UI served from /public
        │
        ▼
Bluebeam Studio API
├── /publicapi/v1
└── /publicapi/v2
The application is a single Express server that keeps all workflow state in memory, including project IDs, session IDs, uploaded file IDs, webhook events, extracted markups, and step logs.

Core Features
Creates a fresh restricted project on each run.

Creates standard folders named resources, review-documents, and markup-exports.

Uploads custom-columns.xml into the project during setup.

Injects a 5-step QC Review state model into PDFs before upload using pdf-lib.

Invites reviewers to both the Session and the Project, then applies folder and per-user permissions.

Exports markups to XML and normalizes them into structured JSON in memory.

Generates reviewed PDF snapshots and saves them to the public/ directory.

5-Step QC State Model
Before upload, the server can inject a custom Bluebeam state model called 5_step_QC_Review into each PDF.

The injected model includes these states: Step 3 - Agree, Step 3 - Disagree, Step 3 - Address in Future Submittal, Step 4 - Revisions Made, Step 5 - Revisions Verified/Concur, and Step 5 - Revisions Incomplete/Disagree.

The server also removes known custom demo models before injection so repeated test runs stay idempotent.

Requirements
Node.js 18+.

Bluebeam API credentials in a .env file.

A valid webhook callback URL if you want real webhook delivery; localhost callback URLs are detected and skipped.

A resources/custom-columns.xml file for project setup, because /poc/setup-project requires it to exist.

One or more PDF files to upload for review.

Quick Start
bash
# 1. Clone the repo
git clone <repo-url>
cd <repo-directory>

# 2. Install dependencies
npm install

# 3. Configure environment
cp .env.example .env

# 4. Start the server
npm start
When the server starts, it runs on http://localhost:3000 by default. The app also serves static content from the public/ folder and exposes the PoC endpoints from the same server.

Environment Variables
Required
Variable	Description
BB_CLIENT_ID	Bluebeam API client ID.
BB_CLIENT_SECRET	Bluebeam API client secret.
Optional
Variable	Description
PORT	Server port, defaults to 3000.
WEBHOOK_CALLBACK_URL	Public callback URL for session event webhooks; defaults to http://localhost:{PORT}/webhook/studio-events.
DEMO_PROJECT_NAME	Default project name used during setup.
DEMO_DOCUMENT_ID	Default document ID used by the in-memory demo stub.
DEMO_DESCRIPTION	Default review description.
The in-memory stub also initializes reviewer data and session permissions, and those values can be updated at runtime through the PoC endpoints.

Example
text
BB_CLIENT_ID=your_client_id
BB_CLIENT_SECRET=your_client_secret

PORT=3000
WEBHOOK_CALLBACK_URL=https://your-public-url.ngrok.io/webhook/studio-events

DEMO_PROJECT_NAME=Demo Project
DEMO_DOCUMENT_ID=DOC-001
DEMO_DESCRIPTION=Design review — coordination update
Default Project Structure
Each run creates a project with these folders:

text
resources/
review-documents/
markup-exports/
The resources folder is used for custom-columns.xml, the review-documents folder is used for uploaded PDFs, and the markup-exports folder is used for exported markup XML files.

API Endpoints
Health and State
Method	Path	Description
GET	/health	Returns health, client ID presence, webhook status, and custom-columns.xml availability.
GET	/poc/state	Returns current in-memory state plus the current stub config.
GET	/poc/stub	Returns the active demo stub.
POST	/poc/configure	Updates project name, document ID, description, reviewer email, or session permissions.
POST	/poc/remove-reviewer	Removes a reviewer from the stub, except the primary reviewer.
POST	/poc/reset	Resets in-memory workflow state.
Workflow Endpoints
Method	Path
POST	/poc/setup-project
POST	/poc/upload-to-project
POST	/poc/apply-custom-columns
POST	/poc/trigger
POST	/poc/create-session
POST	/poc/register-webhook
POST	/poc/checkout-to-session
POST	/poc/invite-reviewers
POST	/poc/checkin
POST	/poc/export-markups
POST	/poc/run-markuplist-job
POST	/poc/downstream-process
POST	/poc/finalize
POST	/poc/snapshot
POST	/poc/cleanup
Standalone Utility Endpoints
Method	Path	Description
POST	/poc/inject-state-model	Uploads a single PDF and returns a state-injected version for testing.
POST	/webhook/studio-events	Receives Bluebeam webhook events and stores them in memory.
GET	/api/project-markups	Returns normalized markup data from the most recent extraction run.
Typical Run Order
For a full happy-path run, use the endpoints in this order:

POST /poc/setup-project

POST /poc/upload-to-project

POST /poc/apply-custom-columns (optional)

POST /poc/trigger

POST /poc/create-session

POST /poc/register-webhook

POST /poc/checkout-to-session

POST /poc/invite-reviewers

Review in Revu

POST /poc/checkin

POST /poc/export-markups

POST /poc/run-markuplist-job

POST /poc/finalize

POST /poc/snapshot

POST /poc/cleanup

If you want one combined downstream step after review, you can replace separate check-in/export/extract calls with POST /poc/downstream-process.

Reviewer Permissions
The server invites reviewers to the Session and also to the Project, then fetches project users so it can assign folder-level and user-level permissions.

For non-owner users, it grants ReadWrite on review-documents, Read on markup-exports, Read on resources, and denies selected project capabilities like CreateSessions and ManagePermissions at the user level.

Markup Extraction
Markup extraction is driven from exported XML, not directly from an active session response.

The server downloads the XML export, parses it with xml2js, normalizes each markup record, and stores fields like author, type, subject, comment, status, layer, page, replies, status history, custom values, and extended properties in memory.

Snapshot Output
The snapshot step polls until each snapshot reaches Complete, downloads the resulting PDF, and writes it to the public/ directory using a reviewed filename pattern based on the project and source file name.

That makes the reviewed PDFs immediately accessible from the same Express static server.

Project Structure
text
├── server.js
├── tokenManager.js
├── public/
├── resources/
│   └── custom-columns.xml
├── .env
├── .env.example
└── package.json
server.js contains the full PoC workflow, tokenManager.js handles OAuth access token refresh, public/ serves the UI and snapshot outputs, and resources/custom-columns.xml is required during project setup.
