# Runbook: 10/10 recovery and continuation

## Purpose

Keep the MarkOut hardening program resumable without chat context. Start here
when a previous agent stopped mid-workstream or when the operational state of
`main`, `release/production`, GitHub Actions, manual OWA verification, or release
governance is uncertain.

## Resume entry point

Run these checks before changing code or rerunning GitHub Actions:

```bash
pwd
git status --short --branch
git log --oneline --decorate -n 8 --all
gh run list --limit 12
gh secret list
gh variable list
gh api repos/Schoenfeld-Solutions/markout/environments
gh api repos/Schoenfeld-Solutions/markout/environments/github-pages/deployment-branch-policies
gh api repos/Schoenfeld-Solutions/markout/rulesets
npm run test:coverage
npm audit --json
```

Current operational baseline on 2026-05-01:

- `main` and `origin/main` point at `066a73f`
  (`test(taskpane): extract app action coverage`).
- The latest `Build and Publish GitHub Pages` run for `066a73f` is green.
- `github-pages` allows deployment from `main` and `release/production`.
- `production-promotion` exists and requires review by `gabrielschoenfeld`.
- GitHub Settings Audit is green.
- `release/production` is protected by the active
  `release-production-automation-only` repository ruleset.
- The only ruleset bypass actor is the `markout-release-bot` GitHub App
  integration with app ID `3567817`; human and repository-role bypass actors
  are not allowed.
- Required release-governance secrets are present:
  `MARKOUT_GITHUB_SETTINGS_AUDIT_TOKEN`, `MARKOUT_RELEASE_BOT_APP_ID`, and
  `MARKOUT_RELEASE_BOT_PRIVATE_KEY`.
- Manual OWA beta verification passed on 2026-05-01 for `066a73f`; see
  [beta-verification.md](./beta-verification.md).
- `npm audit --json` reports zero vulnerabilities.
- Coverage is above the 9/10 baseline but still below the final 10/10 target.
  The known post-ratchet baseline is approximately
  `87.78 / 78.49 / 87.50 / 87.68`.
- `src/taskpane/taskpane-app.tsx` is materially improved but still below the
  final critical-file target for functions and branches, at approximately
  `89.83 / 82.85 / 81.81 / 89.74`.

If these facts differ, update this runbook or the active PR description before
rerunning workflow IDs, changing release settings, or opening the next coverage
ratchet PR.

## Workstream order

1. Production promotion: if the human verifier approves `066a73f` for
   production, run `Promote Production Channel` with that SHA and the
   2026-05-01 beta verification note, approve the `production-promotion`
   environment, wait for the follow-up `release/production` Pages run, and
   perform a human production OWA check.
2. Coverage ratchet: raise global coverage from the current high-80s baseline
   to at least 90% and close critical-file gaps, especially
   `taskpane-app.tsx` functions and branches, `runtime.tsx`, `panels.tsx`, and
   other taskpane edge paths.
3. Targeted edge-case testing: add behavior tests for taskpane lifecycle races,
   hidden-panel fallback, failed optimistic preference saves, stale async
   completions, notification fallback without service support, and developer
   diagnostics visibility.
4. CI maintenance: remove the known GitHub Actions Node 20 deprecation warning,
   especially around dependency-review tooling, without adding paid services or
   long-running jobs.
5. Security and fault-injection durability: keep expanding sanitizer corpus
   tests, Office API failure tests, restore-state edge tests, and notification
   race tests whenever new critical behavior is touched.
6. Documentation and contracts: keep this runbook, beta verification evidence,
   release-governance documentation, and repo-contract checks synchronized with
   every operational policy change.

## Acceptance for the whole program

The project is not `10/10` until all of these are true at the same time:

- The current `main` release is green and beta was manually verified in OWA for
  the exact SHA being considered for promotion.
- `release/production` can only move through the approved promotion path.
- Production promotion and the follow-up production release are green.
- OWA verification remains human-confirmed and is recorded through the
  promotion workflow inputs and `production-promotion` approval.
- `npm audit --json` reports no vulnerabilities.
- GitHub Actions logs do not contain Node 20 deprecation warnings.
- Restore-state cannot lose unrecoverable draft content through TTL expiry.
- Sanitizer, fault-injection, restore, and notification race suites cover the
  critical behavior.
- Coverage is at least 90% globally and at least 85% for critical modules.
- Repository contracts enforce channel, manifest, release, governance, and
  documentation invariants.
