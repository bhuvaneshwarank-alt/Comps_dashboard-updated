# Conversion map and acceptance gates

| Area | Milestone 1 state | Next implementation gate |
|---|---|---|
| Apps Script `Index.html` | Native React sign-in, executive summary, filterable register, detail and history views. | Port remaining navigation, modals, forms, settings and every view; validate visual and keyboard parity. |
| Password and login | Google Workspace ID token verified on Node; existing dashboard password hash, expiry, policy and lockout are reused. | First-time setup, password change, admin reset, OTP issue/verify/replay, complete audit history, recovery, password change revocation and WebAuthn/security review. |
| Server sessions | HTTP-only SameSite cookie, five-minute inactivity window, 24-hour absolute limit, server-side role lookup. | Durable distributed sessions/locks, legacy revocation bridge, concurrent tests, production cookie/domain configuration, identity/session audit log. |
| `getRecords`, `getRecordById`, `getRecordHistory`, `getDashboardData` | Four native read equivalents in the protected API. Summary carries the original open, high/critical, overdue, pending-evidence and closed counts. Register supports department scope, deletion mask, filters, pagination and history. | Compare every field/filter/aggregate and chart against original backend with approved test data and authorization matrix. |
| Create/edit/delete and maker–checker | Not migrated. No record writes exist in the new API. | Native transaction/service layer; immutable history; assignments; separate Approve/Send Back paths; race/duplicate/self-approval tests. |
| IAM/user/roles/temporary access | Sheet-based identity mapping and server-side read scopes exist for the register. Department access fails closed unless the configured Cloud Run service identity equals the approved storage owner. Admin UI and writes not migrated. | Convert IAM read/write workflows, access reviews/requests, protected privilege grants, expiry worker, immutable logs; test TC-02/TC-06 and TC-09. |
| Evidence and Drive | Not migrated. Register data only. | Google Drive API ownership/share controls, file upload/download, scanning, metadata/expiry/approval workflows and department scope; regression tests before importing evidence. |
| Control testing, monitoring, audit readiness, exports | Only a basic register summary is present. | Port controls, reviews, exception states, exports and audit-trail behavior; compare each generated artifact with current prototype. |
| Email and scheduled jobs | Not migrated. | Replace MailApp and Apps Script triggers with a managed email provider and authenticated Cloud Run Jobs/Tasks; idempotency, alerting, owner continuity and recovery tests. |
| Retention, archive, backup and disaster recovery | Not migrated. | Port TC-24 draft/hold/disposal-review rules with no automatic destruction, then TC-21/25 archive integrity, backup location, restore rehearsal, job and log ownership. |
| Storage | Google Sheets API is connected through Node using Cloud Run ADC; demo uses memory fixtures. | Staging verification of approved spreadsheet and service identity; concurrency/quotas and audit protection review. |

## Conversion order

1. Agree and verify the exact source baseline (the upload includes V65 code comments and Index V112 markers, plus TC-24 and TC-25 companion modules).
2. Complete read parity and role-scope testing for the register, dashboards, deletion mask and history.
3. Port security-critical workflows: durable session/revocation, maker–checker, IAM, and authorization tests.
4. Port evidence and controlled file operations; establish service identity, Drive boundaries and immutable audit records.
5. Port scheduled jobs, email, retention, backups, restore, exports and remaining dashboard pages.
6. Run the complete UAT matrix with test accounts across Admin, Compliance Manager, Auditor and department roles. Keep the current Apps Script deployment authoritative until acceptance and the management walkthrough.

Before sharing externally, keep this repository private. The reference snapshot is included only so the technology team can trace the new service back to the exact input files; the new React and Node application does not execute Apps Script code.
