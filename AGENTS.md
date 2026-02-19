# Agents for Boulder Script Project

## BoulderSafety

- Purpose: Default agent for all coding/architecture tasks in this repo.
- Mandatory protocols:
  - `../safe-code-boulder/docs/spec/SAFETY_PROTOCOL_V2.md`
  - `../safe-code-boulder/docs/spec/CODE_CHANGE_PROTOCOL_V2.md`
  - `../safe-code-boulder/docs/spec/BOULDER_L3_GUARDRAILS_V2.md`
- Never edit code/config without an active CHG in `../safe-code-boulder/out/changes/current-change.json`.
- If no active CHG exists, start one with:
  - `npm --prefix ../safe-code-boulder run change:start -- --title "<title>" --surfaces code,api,schema --actor "ai" --reason "<reason>"`
- Before patch: present a short plan.
- After patch: update evidence and run:
  - `npm --prefix ../safe-code-boulder run safety:local`
  - `npm --prefix ../safe-code-boulder run safety:status`
  - `npm --prefix ../safe-code-boulder run change:ready`
- L3 scope (IDs/pay/hours/triggers/automation):
  - DRY_RUN before EXECUTE
  - backup before first EXECUTE
  - max 100 rows per EXECUTE unless user explicitly raises
  - no trigger changes without explicit user approval
