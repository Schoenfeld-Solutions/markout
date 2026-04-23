# ADR 0001: Isolate production, beta, and local add-ins

## Status

Accepted

## Context

MarkOut publishes three Outlook add-in manifests:

- `manifest.xml` for stable production
- `manifest.beta.xml` for the hosted post-merge preview channel
- `manifest-localhost.xml` for local development

The repository previously reused one add-in ID and one shared browser storage
namespace across those channels. That made Outlook state, dismissals, and
restore-state persistence ambiguous and allowed beta or local experiments to
pollute production behavior.

## Decision

MarkOut now treats production, beta, and local as three isolated add-ins.

- Every manifest has its own add-in ID.
- Every hosted runtime URL carries an explicit `channel` query parameter.
- Runtime config resolves the channel explicitly and scopes roaming settings,
  notifications, and restore-state keys through the channel storage namespace.
- Raw draft HTML is no longer persisted in browser `localStorage`.

## Consequences

- Users can install production and beta side by side without state leakage.
- Channel-specific regressions are easier to reason about because storage and
  host wiring are deterministic.
- Future channel additions must update `src/lib/runtime.ts`, all manifests, and
  the repository contract checks in one coherent change.
