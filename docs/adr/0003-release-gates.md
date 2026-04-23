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
- Push releases require host smoke prerequisites up front and execute the host
  smoke after deployment.
- Promotion requires a specific `main` SHA with successful beta deployment
  evidence and runs behind the protected `production-promotion` environment.

## Consequences

- Production cannot advance accidentally from a normal `main` push.
- Missing host-smoke credentials or channel drift are release blockers by
  design, not advisory warnings.
- Repository settings still need a GitHub branch/ruleset configuration so only
  the promotion workflow can update `release/production`.
