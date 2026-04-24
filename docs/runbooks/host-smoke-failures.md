# Runbook: Host smoke failures

## Purpose

Triage failures from `npm run test:host-smoke` when a human intentionally runs
the optional diagnostic script. OWA checks are not scheduled GitHub Actions and
are not release CI gates.

## Fast checks

1. Confirm `MARKOUT_HOST_SMOKE_STORAGE_STATE` contains a valid authenticated
   Outlook session.
2. Confirm `MARKOUT_HOST_SMOKE_RECIPIENT` points to the dedicated test mailbox.
3. Confirm the compose URL still opens the expected Outlook compose surface.
4. Confirm the expected taskpane URL prefix matches the channel under test.

## Common failure classes

- Authentication expired: refresh the local Playwright storage state.
- Selector drift: update the `MARKOUT_HOST_SMOKE_*` selector used for the
  human-initiated diagnostic run.
- Channel drift: the taskpane URL, add-in ID, or manifest no longer match the
  intended beta or production channel.
- Outlook body/notification regressions: inspect the failing screenshot in
  `output/playwright/` and rerun against the local or beta channel.

## Escalation

- Block production promotion until manual beta verification is complete for the
  exact target SHA.
- If only production fails, roll back first and investigate second.
- Update the affected runbook and repository contract checks when the root cause
  comes from documented process drift.
