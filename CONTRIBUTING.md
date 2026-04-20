# Contributing to MarkOut

MarkOut is an Outlook compose add-in that is actively maintained as a fork of
`SierraSoftworks/markout`. Contributions should keep the repo releasable,
security-conscious, and consistent with the Outlook add-in runtime and manifest
constraints documented in [README.md](README.md).

## Working baseline

- Node.js support policy is `>=25 <27`.
- The validated CI baseline is Node 25.
- Before changing dependencies, treat `package.json` as the source of truth and
  keep `package-lock.json` synchronized.
- `jsdom`, `@types/jsdom`, and `markdown-it-emoji` are intentionally pinned below
  `latest` because newer majors are not green in this repo today.

## Branches and pull requests

- Do not commit feature work directly on `main`.
- Use a short-lived feature branch such as `codex/<topic>` or another explicitly
  requested branch name.
- Keep pull requests focused and behaviorally coherent.
- Update README and manifests in the same workstream when setup, hosting,
  support URLs, or Outlook behavior changes.

## Commit format

Conventional Commits are mandatory for tracked changes.

- Required format: `<type>(<scope>): <description>`
- Allowed types: `feat`, `fix`, `refactor`, `docs`, `test`, `chore`, `ci`,
  `build`, `perf`, `revert`
- Use a meaningful scope such as `taskpane`, `renderer`, `manifest`, `tooling`,
  `ci`, `docs`, `commands`, `launchevent`, or `security`
- Keep the subject imperative, specific, and lowercase after the colon
- Do not use vague subjects such as `misc`, `updates`, `cleanup`, `stuff`, `tmp`,
  or `wip`

Examples:

- `feat(taskpane): lazy-load preview renderer`
- `fix(manifest): switch production support url to fork`
- `chore(tooling): widen node policy to 25 and 26`
- `docs(readme): document github pages sideload urls`

## Required quality gates

Run these before every commit, push, or pull request update:

```bash
npm run check
```

This includes formatting checks, linting, type checking, unit tests, the
production build, and deployable manifest validation.

Additional checks when relevant:

- `npm run validate:manifest:localhost` for local manifest work
- `npm run dev` plus manual **Add from File** sideload with `manifest-localhost.xml`
- `npm run start:desktop` if desktop auto-sideload behavior is the thing being changed
- `npm run test:host-smoke` for changes that affect Outlook compose flows, task
  pane behavior, Smart Alerts, selectors, or send-time rendering

If the host smoke cannot be run because credentials or Outlook test
infrastructure are unavailable, call that out explicitly in the PR.

## Dependency and automation policy

- Dependabot version updates are grouped into one weekly tooling PR across
  `npm` and `github-actions`.
- Security updates are intentionally separate from normal version updates.
- Because this repository is a fork, Dependabot security updates and grouped
  security updates must be enabled in GitHub settings for the fork.
- Do not add new runtime dependencies without explicit approval.

## Documentation and licensing

- Keep source code, UI copy, and Markdown docs in English.
- Preserve visible credit to the upstream source in README.
- Do not change the license text unless there is a verified legal reason to do
  so.
