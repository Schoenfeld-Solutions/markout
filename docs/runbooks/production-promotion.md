# Runbook: Promote production

## Purpose

Move one validated `main` commit into the stable `release/production` branch.

## Preconditions

- The target SHA is reachable from `origin/main`.
- The beta verification runbook is complete for that SHA.
- The `Build and Publish GitHub Pages` workflow succeeded for that SHA.
- The person promoting production has a concise manual OWA verification note
  for the exact SHA.
- The `production-promotion` environment requires explicit reviewer approval.
- The `markout-release-bot` GitHub App credentials are configured as
  `MARKOUT_RELEASE_BOT_APP_ID` and `MARKOUT_RELEASE_BOT_PRIVATE_KEY`.
- The `release/production` ruleset blocks direct human pushes and allows only
  the release bot integration bypass.

## Steps

1. Open the `Promote Production Channel` workflow.
2. Enter the exact validated `main` SHA.
3. Set `beta_verification_confirmed` to `true`.
4. Enter `beta_verification_notes` with verifier, date, and checked OWA flows.
5. Approve the `production-promotion` environment when prompted.
6. Wait for the workflow to verify beta deployment evidence and push the fast
   forward update to `release/production` with the release bot token.
7. Watch the follow-up release workflow for the `release/production` push.
8. Confirm the production channel manually in OWA with `manifest.xml`.

## Rollback

- If production fails after promotion, follow
  [rollback.md](./rollback.md) with the previously known-good
  `release/production` SHA.

## Notes

- This workflow is the only supported way to update production.
- If repository rules, release bot credentials, or beta evidence are missing,
  stop and configure them before promotion.
- Bootstrap and validation details live in
  [release-bot-bootstrap.md](./release-bot-bootstrap.md).
