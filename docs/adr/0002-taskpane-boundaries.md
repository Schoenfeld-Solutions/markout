# ADR 0002: Split the taskpane into a shell plus explicit controllers

## Status

Accepted

## Context

The original `src/taskpane/app.tsx` accumulated rendering, host integration,
selection polling, notification lifecycle management, editor bootstrapping, and
all panel UI in one file. That made the taskpane difficult to test and forced
behavioral changes to cross unrelated concerns.

## Decision

The taskpane is split into smaller modules with explicit boundaries.

- `taskpane-app.tsx` remains the composition shell for state and user actions.
- `controllers.ts` owns preview refresh, selection observation, and auto-render
  notification lifecycle.
- `editor.tsx` owns CodeMirror bootstrapping and diagnostics.
- `preferences.ts`, `theme.ts`, `toolbar.tsx`, and `file-drop.ts` contain pure
  helpers or focused hooks.
- `panels.tsx` owns panel rendering and keeps panel components free from Office
  side effects.
- Runtime diagnostics are an in-memory ring buffer exposed to the developer
  panel through props. Diagnostic events store operational codes, levels, and
  bounded metadata only; draft body, selection text, session data, tokens, and
  recipient fields are redacted before display.

## Consequences

- Taskpane logic can be tested at module level instead of only through one large
  integration surface.
- Office/Outlook IO is concentrated in runtime wiring and service boundaries.
- Future panel additions should land as new props and pure components instead of
  new side effects in the shell.
