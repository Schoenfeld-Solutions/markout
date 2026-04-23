# Runbook: Roll back production

## Purpose

Restore the stable production channel to the previously known-good
`release/production` commit.

## Trigger conditions

- Production host smoke fails after promotion.
- Outlook host compatibility regresses on the production channel.
- A manifest/runtime mismatch breaks the installed production add-in.

## Steps

1. Identify the last known-good `release/production` SHA from the previous green
   promotion or release run.
2. Re-run the `Promote Production Channel` workflow with that known-good SHA.
3. Wait for the production release workflow to rebuild GitHub Pages.
4. Confirm the post-deploy host smoke is green again.
5. Document the incident and the failing promoted SHA in the pull request or
   incident notes.

## Do not do this

- Do not force-push arbitrary history onto `release/production`.
- Do not roll back only the manifest or only the hosted runtime files.
- Do not skip the production host smoke after rollback.
