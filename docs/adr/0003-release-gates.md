# ADR 0003: Enforce release and promotion through explicit GitHub gates

## Status

Accepted

## Context

MarkOut uses GitHub Pages for both the hosted beta channel and the stable
production channel. A previous workflow allowed fallback behavior when
`release/production` was missing, and production promotion only checked
fast-forward viability.

## Decision

MarkOut now treats release and promotion as fail-closed workflows.

- Pull requests into `main` must pass repo contracts, build/test gates,
  dependency review, `test:taskpane-ui`, and the coverage gate.
- The release workflow refuses to publish if `origin/release/production` does
  not exist.
- Promotion requires a specific `main` SHA with successful beta deployment
  evidence, explicit manual OWA beta verification input, and the protected
  `production-promotion` environment approval.
- Production branch updates use a dedicated `markout-release-bot` GitHub App
  token when repository rules cannot safely distinguish approved automation from
  human pushes through GitHub-native branch protection alone. The app has no
  server, no scheduler, no external runtime, and only repository-scoped
  `Contents: write` access.
- A scheduled GitHub settings audit verifies release branches, Pages branch
  policies, promotion environment reviewers, release bot credentials, and the
  `release/production` ruleset.

## Consequences

- Production cannot advance accidentally from a normal `main` push.
- OWA testing remains human-confirmed. GitHub Actions does not store permanent
  Outlook Web test credentials or run scheduled OWA smoke tests.
- Repository settings must keep the `release/production` ruleset aligned with
  the release bot. Drift is an operational failure, not a documentation issue.
- The release bot private key is a security-sensitive credential and must be
  rotated if it is exposed.
