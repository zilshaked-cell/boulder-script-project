# Copilot – Boulder Shifts Apps Script

You are working inside the **Google Apps Script backend** for Boulder Shifts. It exposes a Web App that the Next.js app calls through `/api/appscript`.

## Architecture

- Main entrypoint: `index.gs` + `entrypoints.js`.
- Each action (e.g., `jobTypes.list`, `getCurrentEmployee`, `employeeExistsByEmail`, `shiftReport.submit`) is part of the public contract with the web app.
- Data source: Google Sheets (employees, job types, shift reports, requests). Sheet names and headers are part of the schema contract.

## Safety & schema

- **Never** change script IDs, deployment IDs or triggers. Do not add/modify time-based or installable triggers without explicit instruction.
- When changing sheet structure: do not rename sheets/headers; add new columns only to the **right**; never reuse ID columns.
- Preserve existing ID fields (`employeeId`, shift/report IDs, request IDs); do not regenerate IDs for existing rows.

## jsonResponse & errors

- All HTTP responses from the Web App must use `jsonResponse`.
- `jsonResponse` returns `ContentService.MimeType.JSON` and must **not** call `setResponseCode` on the `TextOutput`.
- If a numeric `status` is provided, store it in `body.meta.statusCode` (do not change HTTP status). Responses follow `{ ok, data?, errorCode?, error?, meta: { traceId?, operation?, statusCode? } }`.

## Logging

- Use `Logger.log` only for technical diagnostics (operation, errorCode, counts). Keep data sanitized (IDs/flags/booleans only).
- Never log full request bodies, employee records, or raw personal data. Prefer “what happened” (counts, flags) over full payloads.

## Employee lookup

- `getCurrentEmployee` / `employeeExistsByEmail_` should match by email in the employees sheet, respect “active” flags, and return null/errorCode when missing/inactive rather than throwing.
