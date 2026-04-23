# Runbook: Host smoke failures

## Purpose

Triage failures from `npm run test:host-smoke` in local runs or GitHub Actions.

## Fast checks

1. Confirm `MARKOUT_HOST_SMOKE_STORAGE_STATE_JSON` or
   `MARKOUT_HOST_SMOKE_STORAGE_STATE` contains a valid authenticated Outlook
   session.
2. Confirm `MARKOUT_HOST_SMOKE_RECIPIENT` points to the dedicated test mailbox.
3. Confirm the compose URL still opens the expected Outlook compose surface.
4. Confirm the expected taskpane URL prefix matches the channel under test.

## Common failure classes

- Authentication expired: refresh the Playwright storage state.
- Selector drift: update the `MARKOUT_HOST_SMOKE_*` selectors and keep README,
  CONTRIBUTING, and the workflow aligned.
- Channel drift: the taskpane URL, add-in ID, or manifest no longer match the
  intended beta or production channel.
- Outlook body/notification regressions: inspect the failing screenshot in
  `output/playwright/` and rerun against the local or beta channel.

## Escalation

- Block production promotion until the beta host smoke is green again.
- If only production fails, roll back first and investigate second.
- Update the affected runbook and repository contract checks when the root cause
  comes from documented process drift.
