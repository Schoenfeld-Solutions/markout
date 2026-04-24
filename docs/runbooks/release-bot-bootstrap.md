# Runbook: Release bot bootstrap

## Purpose

Make `release/production` automation-only without adding a server, scheduler, or
paid external release system. The release bot exists only to let the approved
promotion workflow fast-forward production while blocking normal human pushes.

## Required GitHub App

Create a dedicated GitHub App named `markout-release-bot`.

- Repository access: only `Schoenfeld-Solutions/markout`
- Repository permissions: `Contents: Read and write`
- Organization permissions: none
- Webhook: disabled
- External runtime: none

Install the app on this repository only.

## Repository secrets

Store the app credentials as repository secrets:

```bash
gh secret set MARKOUT_RELEASE_BOT_APP_ID
gh secret set MARKOUT_RELEASE_BOT_PRIVATE_KEY < path/to/private-key.pem
```

Do not store the private key in the repository, logs, issue comments, or pull
request descriptions.

## Ruleset

After the initial `release/production` bootstrap branch exists, create an active
repository ruleset for `refs/heads/release/production`.

The ruleset must:

- block branch creation, updates, and deletion for normal actors
- include no user, team, repository-role, or admin bypass actors
- allow only the `markout-release-bot` GitHub App as an integration bypass

This keeps production stable while allowing the approved
`Promote Production Channel` workflow to mint a short-lived installation token
and fast-forward `release/production`.

## Validation

Run the audit after secrets and the ruleset are configured:

```bash
gh workflow run "GitHub Settings Audit"
gh run list --workflow "GitHub Settings Audit" --limit 3
```

The local script can also be run with a token that can read repository settings:

```bash
MARKOUT_GITHUB_TOKEN=<settings-read-token> npm run check:github-release-governance
```

The audit must fail if:

- `release/production` is missing
- `github-pages` does not allow `main` and `release/production`
- `production-promotion` lacks required reviewers
- release-bot secrets are missing
- the production ruleset is absent or includes human bypasses

## Manual push test

After the ruleset is active, a direct human push must fail:

```bash
git push origin HEAD:refs/heads/release/production
```

Do not force-push. A rejected non-forced update is sufficient evidence that the
ruleset is active.

## Rollback

If the app or ruleset blocks a required emergency promotion, pause and fix the
ruleset or app installation. Do not bypass the model with a direct human push
unless production is already unavailable and the incident owner explicitly
approves the exception in writing.
