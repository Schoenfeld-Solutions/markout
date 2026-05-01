# Runbook: Beta verification

## Purpose

Verify the hosted beta channel that is published from `main`.

## Preconditions

- The PR was merged into `main`.
- The `Build and Publish GitHub Pages` workflow succeeded for that merge SHA.
- The beta manifest is installed in Outlook with **Add from File**.

## Steps

1. Confirm the GitHub Actions release run for the target `main` SHA is green.
2. Install or refresh `manifest.beta.xml`.
3. Open Outlook compose and confirm the taskpane loads from
   `https://schoenfeld-solutions.github.io/markout/outlook-beta/taskpane.html`.
4. Verify the main compose flows:
   - render selection
   - render Markdown-looking draft blocks without changing a non-Markdown
     signature
   - insert rendered Markdown
   - toggle auto-render in settings
5. Confirm no browser settings, notification state, or restore-state behavior
   leaked from production or local add-ins.
6. Record the verified SHA, verifier, date, and checked flows for the
   `beta_verification_notes` promotion input.

## Abort criteria

- The beta taskpane loads from the wrong URL or wrong add-in ID.
- Auto-render notifications or restore-state behavior leak from another channel.
- The release workflow is not green for the target SHA.
- The verifier cannot confirm the required OWA flows manually.

## Verification log

### 2026-05-01 beta quick check

- Result: passed.
- Verified SHA: `066a73f` (`test(taskpane): extract app action coverage`).
- Channel and manifest: beta channel through `manifest.beta.xml`.
- Client: Outlook Web compose with the hosted beta taskpane.
- Verifier: human verification by Gabriel.
- Evidence: screenshots were supplied in the project thread and intentionally
  were not committed to the repository.
- Checked behavior:
  - The taskpane opened as the beta add-in.
  - The settings panel loaded with the expected localized UI, theme controls,
    language selector, and settings toggles.
  - The insert panel loaded in dark mode.
  - Markdown input rendered a heading, paragraph, and list preview.
  - The bottom toolbar stayed visible while the content area scrolled above it.
  - No visible beta/production channel-state leak was observed in the quick
    check.

This log is human OWA verification evidence only. It is not automated OWA test
evidence and must not be used to justify adding unattended OWA automation.
