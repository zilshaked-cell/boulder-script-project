# Copilot Instructions

This folder is the Google Apps Script backend for the Boulder Haifa shifts system. It is managed with clasp and works alongside the `boulder-shifts-app` frontend.

## Behavior rules for code generation
- Prefer minimal, focused patches that solve the specific issue.
- Keep Apps Script triggers (`onOpen`, `onEdit`, entrypoints) and public APIs stable and backward compatible.
- Do not add new external dependencies unless explicitly requested.
- Preserve existing function and file structure; avoid renaming triggers or moving bindings.
- Do not change spreadsheet side effects or bindings unless explicitly asked.
- When in doubt, ask for clarification in comments or chat instead of guessing.
- Ensure lint/style stay clean with no new errors or warnings.
- Never silently change business logic, data schema, or shift/reporting semantics unless asked.
