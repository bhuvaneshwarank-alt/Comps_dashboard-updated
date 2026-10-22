# Milestone 2 — Native read-only analytics (partial Phase 2)

This increment extends M1 without switching on any write workflow.

## Implemented
- `GET /api/analytics` uses existing authenticated session and checks both dashboard and record-view privileges.
- Calls `visibleRecords()` before aggregating, retaining department scoping and deletion masking.
- Returns record counts by department, primary framework, type, status and risk.
- React analytics navigation displays grouped totals and horizontal bars, with no Apps Script browser dependency.
- Tests cover source fixture totals, subset scoping, and empty input.

## Not represented
This is **not** the complete original Analytics tab or the full Phase 2 migration. Cross-framework mappings, control-level reporting, risk heat maps, readiness scoring, export functionality and full UI parity remain open. The source `Code.gs` and `Index.html` snapshots are included for the migration engineer to compare. Existing original UI is not replaced.

## Validation
`node --test server/test/analytics.test.js` passes on local Node without npm dependencies.
A full `npm ci && npm test && npm run build` must be executed on a Node 22.12+ machine with network package access before accepting the artifact; the execution environment lacked a complete npm package cache.

## Security note
The API never trusts browser-supplied department, user or role input to determine scope. Any role or permission updates and record modification endpoints must preserve the same server-side constraints and receive independent tests before enabling write operations.
