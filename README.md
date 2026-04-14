Bluebeam Studio — Dynamic Project Review PoC
Disclaimer
This repository contains a proof-of-concept integration showing how an external system can drive a Bluebeam Studio review workflow across Studio Projects and Studio Sessions using the Bluebeam API.

This project is a reference implementation for evaluation and development use only and is not intended for production use.

Overview
This PoC:

Creates a new restricted Studio Project for each run.

Sets project, folder, and per-user permissions programmatically.

Uploads review PDFs into a dedicated project folder.

Optionally injects a 5‑step QC Review state model into PDFs before upload (via pdf-lib).

Checks project files out into a Studio Session, invites reviewers, and later checks them back in.

Exports markups to XML and normalizes them into structured JSON in memory.

Creates snapshots of reviewed PDFs and saves them to public/ for download.

Workflow
The typical end‑to‑end workflow is:

Step	Endpoint	Description
0a	POST /poc/setup-project	Creates a new restricted Studio Project, creates standard folders, and uploads custom-columns.xml.
0b	POST /poc/upload-to-project	Uploads PDF files into the project, optionally injecting the 5‑step QC state model before upload.
0c	POST /poc/apply-custom-columns	Applies custom columns to uploaded project files.
1	POST /poc/trigger	Simulates the source-system workflow event and logs the active files and reviewers.
2	POST /poc/create-session	Creates a restricted Studio Session and applies default session permissions.
3	POST /poc/register-webhook	Registers a session webhook subscription (skipped automatically if the callback URL is localhost).
4	POST /poc/checkout-to-session	Checks uploaded project files out into the Studio Session.
5	POST /poc/invite-reviewers	Invites reviewers to the Session and Project, fetches project user IDs, and sets folder and per‑user permissions.
6	(manual)	Reviewers open the Session in Revu and add markups.
7	POST /poc/checkin	Checks session files back into the project using the documented check‑in pattern.
8	POST /poc/export-markups	Runs exportmarkups to generate XML markup exports in the project.
9	POST /poc/run-markuplist-job	Downloads exported XML and extracts normalized markup data into memory.
9b	POST /poc/downstream-process	Combined downstream step: check‑in, export markups, and extract markup data.
10	POST /poc/finalize	Finalizes the Studio Session.
11	POST /poc/snapshot	Creates snapshots of reviewed session files and saves PDFs into public/.
12	POST /poc/cleanup	Deletes the webhook subscription (if any) and deletes the Session.
You can either call steps 7–9 individually or use the combined POST /poc/downstream-process endpoint after review.

Architecture
text
Source System (stubbed)
        │
        │  workflow event
        ▼
Express Backend (server.js)
├── OAuth token management (TokenManager)
├── Dynamic project creation
├── Folder + per-user permission setup
├── PDF state model injection (pdf-lib)
├── Project upload / session checkout / check-in flow
├── Markup export + XML parsing
├── Webhook receiver (/webhook/studio-events)
└── Static UI served from /public
        │
        ▼
Bluebeam Studio API
├── /publicapi/v1
└── /publicapi/v2
The application is a single Express server that keeps all workflow state (project IDs, session IDs, file IDs, webhook events, extracted markups, and step logs) in memory.

5‑Step QC State Model
Before uploading PDFs, the server can inject a custom Bluebeam state model named 5_step_QC_Review:

Step 3 – Agree

Step 3 – Disagree

Step 3 – Address in Future Submittal

Step 4 – Revisions Made

Step 5 – Revisions Verified/Concur

Step 5 – Revisions Incomplete/Disagree

Injection is done using pdf-lib by adding a Collab.addStateModel document‑level script and setting an OpenAction. Known demo models are removed first so repeated runs remain idempotent.

There is also a standalone endpoint:

POST /poc/inject-state-model — multipart upload (file field), returns the same PDF with the QC state model injected.

Requirements
Node.js 18+.

Bluebeam API credentials (BB_CLIENT_ID, BB_CLIENT_SECRET).

Publicly reachable WEBHOOK_CALLBACK_URL if you want to receive real Studio webhooks (localhost URLs are detected and skipped).

resources/custom-columns.xml present on disk for /poc/setup-project.

One or more review PDFs to upload.

Environment Variables
Create a .env file (you can base it on .env.example).

Required
Variable	Description
BB_CLIENT_ID	Bluebeam API client ID.
BB_CLIENT_SECRET	Bluebeam API client secret.
Optional
Variable	Default	Description
PORT	3000	Port the Express server listens on.
WEBHOOK_CALLBACK_URL	http://localhost:{PORT}/webhook/studio-events	Webhook callback URL; if this points to localhost, webhook registration is skipped.
DEMO_PROJECT_NAME	Demo Project	Default project name used for setup-project.
DEMO_DOCUMENT_ID	DOC-001	Default document ID used in stub state.
DEMO_DESCRIPTION	Design review — coordination update	Default review description.
Example:

text
BB_CLIENT_ID=your_client_id
BB_CLIENT_SECRET=your_client_secret

PORT=3000
WEBHOOK_CALLBACK_URL=https://your-public-url.ngrok.io/webhook/studio-events

DEMO_PROJECT_NAME=Demo Project
DEMO_DOCUMENT_ID=DOC-001
DEMO_DESCRIPTION=Design review — coordination update
Project Folders
Each run creates a project with these folders:

text
resources/
review-documents/
markup-exports/
resources — holds custom-columns.xml.

review-documents — holds uploaded PDFs.

markup-exports — holds exported XML from exportmarkups jobs.

API Surface
Health & Configuration
Method	Path	Description
GET	/health	Returns basic health plus flags for client ID presence, webhook URL, and custom-columns.xml existence.
GET	/poc/state	Returns current in‑memory PoC state (IDs, markups, webhook events, log, stub).
GET	/poc/stub	Returns the current demo stub configuration.
POST	/poc/configure	Updates stub fields (project name, document ID, description, reviewer email, permissions) at runtime.
POST	/poc/remove-reviewer	Removes a reviewer from the stub (except the primary reviewer).
POST	/poc/reset	Resets in‑memory state for a fresh run.
Workflow Endpoints
These implement the workflow table above:

POST /poc/setup-project

POST /poc/upload-to-project

POST /poc/apply-custom-columns

POST /poc/trigger

POST /poc/create-session

POST /poc/register-webhook

POST /poc/checkout-to-session

POST /poc/invite-reviewers

POST /poc/checkin

POST /poc/export-markups

POST /poc/run-markuplist-job

POST /poc/downstream-process

POST /poc/finalize

POST /poc/snapshot

POST /poc/cleanup

Utility & Webhook
Method	Path	Description
POST	/poc/inject-state-model	Returns a single uploaded PDF with the QC state model injected (no Bluebeam API calls).
POST	/webhook/studio-events	Receives Studio webhook events, logs them, and stores them under pocState.webhookEvents.
GET	/api/project-markups	Returns normalized markup records from the last extraction run.
Typical Run (Step‑By‑Step)
POST /poc/setup-project

POST /poc/upload-to-project (with one or more PDFs)

POST /poc/apply-custom-columns (optional but uses custom-columns.xml)

POST /poc/trigger

POST /poc/create-session

POST /poc/register-webhook

POST /poc/checkout-to-session

POST /poc/invite-reviewers

Review in Revu

POST /poc/checkin or POST /poc/downstream-process if you want the combined downstream step.

If not using downstream-process: POST /poc/export-markups, then POST /poc/run-markuplist-job.

POST /poc/finalize

POST /poc/snapshot

POST /poc/cleanup.

Reviewer & Permission Model
The demo stub includes:

Project name, document ID, description.

A list of reviewer emails, including a primary reviewer.

Default session permissions (e.g., Markup, SaveCopy, PrintCopy allowed; AddDocuments and FullControl denied).

/poc/invite-reviewers:

Invites each reviewer to the Session.

Invites each reviewer to the Project.

Fetches project users to map emails to user IDs.

Sets folder permissions per user (e.g., review-documents as ReadWrite, markup-exports as Read).

Sets per‑user project permissions such as denying CreateSessions and ManagePermissions.

Markup Export & Extraction
The markup extraction pipeline is:

POST /poc/export-markups — submits exportmarkups jobs to XML for each session file, stored under markup-exports/.

POST /poc/run-markuplist-job — downloads XML by path, parses it with xml2js, and normalizes records into a flat JSON array.

Normalization includes fields such as ID, author, subject, comment, status, layer, page, created/modified dates, status history, replies, and extended/custom properties. The results are stored on pocState.markups and returned from /api/project-markups.

Snapshots & Output
POST /poc/snapshot:

Submits snapshot jobs for each session file.

Polls until the snapshot is available.

Downloads each resulting PDF and writes it into public/ with a name like <ProjectName>_<FileName>_Reviewed.pdf.

These PDFs are then served directly by the Express static middleware.

Project Structure
text
├── server.js            # Express backend and all PoC routes
├── tokenManager.js      # OAuth 2.0 token management with auto-refresh
├── public/              # Static UI + snapshot PDFs
├── resources/
│   └── custom-columns.xml
├── .env                 # Your credentials and config (gitignored)
├── .env.example         # Template env file
└── package.json
