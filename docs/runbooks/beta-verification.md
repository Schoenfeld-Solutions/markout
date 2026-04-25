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
