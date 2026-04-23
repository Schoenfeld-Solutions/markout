# Runbook: Promote production

## Purpose

Move one validated `main` commit into the stable `release/production` branch.

## Preconditions

- The target SHA is reachable from `origin/main`.
- The beta verification runbook is complete for that SHA.
- The `Build and Publish GitHub Pages` workflow succeeded for that SHA.
- The `production-promotion` environment requires explicit reviewer approval.
- GitHub branch rules or rulesets block direct human pushes to
  `release/production`.

## Steps

1. Open the `Promote Production Channel` workflow.
2. Enter the exact validated `main` SHA.
3. Approve the `production-promotion` environment when prompted.
4. Wait for the workflow to verify beta deployment evidence and push the fast
   forward update to `release/production`.
5. Watch the follow-up release workflow for the `release/production` push.
6. Confirm the post-deploy production host smoke passes.

## Rollback

- If production fails after promotion, follow
  [rollback.md](./rollback.md) with the previously known-good
  `release/production` SHA.

## Notes

- This workflow is the only supported way to update production.
- If repository rules are missing, stop and configure them before the next
  promotion.
