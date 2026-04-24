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

Expected recovery baseline on 2026-04-24:

- `main` and `origin/main` point at
  `726bb1b41cec31fc435a5ebd147110b6cd3b331f`.
- `origin/release/production` points at
  `c4795a4d5ad56ef685416eb2d506a979197169a2`.
- The `Build and Publish GitHub Pages` run `24838117797` for
  `726bb1b41cec31fc435a5ebd147110b6cd3b331f` is the red `main` push run that
  must be rerun after external prerequisites are set.
- `github-pages` allows deployment from `main` and `release/production`.
- `production-promotion` exists and requires review by `gabrielschoenfeld`.
- Before the release-bot workstream is complete, `release/production` may still
  lack an automation-only ruleset.
- Coverage is expected to be below the final target. The known recovery baseline
  is approximately `71.21 / 59.63 / 71.57 / 71.13`.
- `npm audit --json` is expected to report critical and moderate findings until
  the supply-chain workstream removes the legacy Office add-in tooling.

If these facts differ, update this runbook or the active PR description before
rerunning old workflow IDs.

## Workstream order

1. Merge operational recovery: remove automated OWA credential gates, rerun the
   failed `main` push run after the release workflow no longer requires OWA
   secrets, verify beta manually in OWA, promote the verified SHA, and verify
   production manually.
2. Release governance: use GitHub-native branch rules only if they can block
   human pushes while preserving promotion automation. Otherwise use the
   `markout-release-bot` GitHub App documented in
   [release-bot-bootstrap.md](./release-bot-bootstrap.md).
3. Manual OWA verification durability: keep the beta verification runbook
   explicit enough that a human can verify the exact SHA without chat context.
4. Supply-chain cleanup: remove the vulnerable legacy Office tooling and keep
   `npm audit --json` free of critical and moderate findings.
5. Restore-state hardening: guarantee that MarkOut never loses current draft
   content because of a timer. Only reconstructable artifacts may expire.
6. Security, fault injection, and observability: expand sanitizer corpus tests,
   Office API failure tests, notification race tests, and in-memory diagnostic
   events.
7. Taskpane decomposition and coverage: split the remaining large taskpane
   shell and raise global coverage to at least 90% with critical files at least
   85%.

## Acceptance for the whole program

The project is not `10/10` until all of these are true at the same time:

- The `main` release for `726bb1b41cec31fc435a5ebd147110b6cd3b331f` is green
  and beta was manually verified in OWA.
- `release/production` can only move through the approved promotion path.
- Production promotion and the follow-up production release are green.
- OWA verification remains human-confirmed and is recorded through the
  promotion workflow inputs and `production-promotion` approval.
- `npm audit --json` reports no critical or moderate findings.
- GitHub Actions logs do not contain Node 20 deprecation warnings.
- Restore-state cannot lose unrecoverable draft content through TTL expiry.
- Sanitizer, fault-injection, restore, and notification race suites cover the
  critical behavior.
- Coverage is at least 90% globally and at least 85% for critical modules.
- Repository contracts enforce channel, manifest, release, governance, and
  documentation invariants.
