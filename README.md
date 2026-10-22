# DigiClarity Compliance Dashboard — Cloud migration, consolidated migration workspace (incomplete; staging only)

This is a real React + Node.js project built from the uploaded Apps Script V65 / Index.html V112 source set. It starts the migration with **Google sign-in plus the existing dashboard password, read-only compliance register, department filtering, record details, audit-history reads, and executive summaries**. The demo runs on synthetic data.

This is a working first milestone, not a claim that the full application has already been converted. The 130 protected Apps Script operations are mapped in [`docs/source-inventory.json`](docs/source-inventory.json); most still need a Node implementation and workflow tests. Maker–checker, evidence upload/download, password setup/reset and OTP, IAM administration, approval centers, reminders, monitoring schedules, retention requests, archives, backups and recovery are not implemented in this milestone. The original Apps Script snapshot is retained for private developer reference; the React app never runs it in a compatibility VM and the browser cannot invoke it.

The imported Apps Script rules used by the first milestone are extracted **byte-for-byte** from the uploaded source into `server/src/source-rules.js`. A test compares the ported functions with the original source, so changes can be reviewed before later functions are added. Register and overview read operations now carry department scope, deletion mask, filtering, pagination, history, and the original summary's open/high-risk/overdue/pending-evidence/closed counts.

## Run the synthetic preview

Requires Node.js 22.12 or later.

```sh
cp .env.example .env
npm ci
npm test
npm run build
npm run dev
```

Open <http://localhost:5173>. Choose the Administrator or HR demo role. Both use synthetic records and have no connection to the company spreadsheet. In the preview, Admin sees six made-up records; the HR role sees two. Nothing in demo mode can reach Google APIs.

For a single local production-style process, build and run with `npm run build` then `npm start`; the app serves the built React assets at <http://localhost:8080>.

## Google-backed staging setup

Live staging is opt-in (`APP_MODE=google`). The initial milestone preserves **Google Sheets as the database**. It reads the tabs that already exist in the bound spreadsheet; it does not initialize sheets or alter compliance records. Authentication writes are limited to the existing `Dashboard_Credentials` values in columns H:J for failed-attempt lockout and last-login bookkeeping. No evidence or business data is changed.

1. Create a private GitHub repository and add the project. The reference copy contains internal business logic and organization-specific identifiers; do not make this repository public.
2. Create a Google OAuth web client for the company sign-in origin. Set the client ID in `GOOGLE_CLIENT_ID`, set exact authorized site origin in `APP_ORIGIN`, and list only approved Workspace domains in `ALLOWED_DOMAINS`.
3. Create a dedicated Google Cloud Run service identity. Enable the Google Sheets API and grant that identity access only to the existing dashboard spreadsheet. Configure `GOOGLE_SHEETS_ID`. For department-scoped accounts, set `CLOUD_RUN_SERVICE_ACCOUNT_EMAIL` and `IAM_STORAGE_OWNER` to the exact reviewed service identity; department data access fails closed unless they match. Use Cloud Run Application Default Credentials; do not upload or commit a service-account key.
4. Set `DGS_DASHBOARD_PASSWORD_PEPPER_V1` to the **existing** pepper value from Apps Script Script Properties. Copying the same value is required to validate existing dashboard password hashes. Do not rotate or regenerate it during this migration. This milestone cannot create first-time passwords or complete forgot-password/OTP flows.
5. Build and test in staging with synthetic records or an approved test spreadsheet. Confirm user/role tabs and `Compliance_Master_Register`, `Record_Deletion_Register`, and `Record_History` headers exactly match the current source. A malformed or missing authorization/data schema fails closed.

Deploy with a single Cloud Run instance during this milestone because lockout serialization and session revocation are process-local. An instance restart signs users out; multiple instances do not share session state yet. Do not enable production traffic until the migration covers all protected operations and approved UAT completes.

```sh
gcloud run deploy digiclarity-compliance \
  --source . \
  --region REGION \
  --service-account CLOUD_RUN_SERVICE_ACCOUNT \
  --max 1 \
  --set-env-vars APP_MODE=google,APP_ORIGIN=https://YOUR_HOST,GOOGLE_CLIENT_ID=YOUR_CLIENT_ID,GOOGLE_SHEETS_ID=YOUR_SHEET_ID,ALLOWED_DOMAINS=digiclarity.in\,clarityrcm.com
```

Set the `IAM_STORAGE_OWNER` and `DGS_DASHBOARD_PASSWORD_PEPPER_V1` using Cloud Run secret references after the authorized administrator verifies the project and service identity.

Pass the pepper with a Cloud Run secret/environment secret reference, not in the checked-in command, source, or repo. Set `APP_ORIGIN` to the final HTTPS origin. The Cloud Run service identity and OAuth consent/authentication setup must be reviewed by the company's Workspace/Cloud administrators.

## Review map

- `client/src/` is the native React application. It has no `google.script.run` dependency.
- `server/src/` is the Node API, access checks, identity verification, spreadsheet API access and read-only register rules.
- `reference/apps-script/` is a private, unmodified copy of the supplied Apps Script project for engineers to compare against.
- `docs/migration-map.md` describes current coverage and the order for later migration phases.
- `docs/source-inventory.json` records source hashes, Apps Script services and the complete protected-operation inventory.

Every record/detail/history API resolves the authenticated Workspace identity, verifies its active Sheet role, checks temporary-access recovery state, enforces department scope and checks permissions on the server. It never accepts an email or role sent by the browser. The Google ID token is verified server-side against the OAuth client ID, verified company email and Workspace `hd` domain. Password comparisons retain the uploaded HMAC-SHA256 chain (180 rounds and existing Script Property pepper); the plaintext password is never stored. Business-record routes are read-only in this milestone.

## Limitations before production

- This project is a verified migration starting point, not an accepted replacement for the existing prototype.
- Only 4 of 130 protected server operations have a corresponding native operation (`getRecords`, `getRecordById`, `getRecordHistory`, and `getDashboardData` summary semantics). See the generated inventory for the rest.
- Existing signup, password changes, policy upgrade, session history, OTP, mail, file handling, triggers, exports, approvals and admin functions remain in Apps Script until separately ported and re-tested.
- Process-local sessions and rate limits require a shared, durable session/lock store before scaling beyond one Cloud Run instance. Session revocation from the legacy app does not reach this new session store.
- Tests validate code and synthetic fixture behavior. They cannot validate this company's OAuth client, service-account access or production sheet state from the local workspace.
- Keep real compliance records and evidence out of fixtures and Git. Import real data only after privacy, role, session, file-access, logging, backup, ownership and operational acceptance checks pass.


## Milestone 2 increment

A native, read-only `/api/analytics` endpoint and React analytics page were added. These report *record distributions within the authenticated user’s existing scope*, not audit readiness or certified control effectiveness. Details, remaining work, and test caveats are in [`docs/milestone-2.md`](docs/milestone-2.md). All previously identified unfinished workflows remain unfinished.

## Latest consolidated-source addition

Read-only CSV export (`/api/records-export.csv`) is scoped to visible records on the server. It uses a fixed allowlisted set of columns and escapes CSV formulas. Export integration, all unported workflows, live Google access and full production acceptance remain unverified. **Not a complete or production-ready application.**
