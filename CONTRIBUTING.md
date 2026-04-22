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
- Use a short-lived feature branch in the form `dev/<topic>`.
- Treat `main` as the integration branch for the hosted beta/testing channel.
- Treat `release/production` as the stable production source branch.
- Keep pull requests focused and behaviorally coherent.
- Update README and manifests in the same workstream when setup, hosting,
  support URLs, or Outlook behavior changes.

## Taskpane UX guardrails

- Treat the task pane as the only primary compose workspace for MarkOut.
- Keep `Open MarkOut` as the single compose command entry point unless a new
  command is explicitly approved and documented.
- Do not build native Outlook context-menu flows for Markdown rendering in this
  repo. Outlook web add-ins do not provide a stable path for that scenario, so
  MarkOut uses taskpane-first selection and insert flows instead.
- Follow Microsoft Office add-in design guidance for task panes and layout:
  - [Office Add-in design language](https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-design-language)
  - [Task panes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/task-pane-add-ins)
  - [Layout guidelines for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-layout)
  - [Use Fluent UI React in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/using-office-ui-fabric-react)
- For repo history and product constraints, also review:
  - [OfficeDev issue #2364](https://github.com/OfficeDev/office-js/issues/2364)
  - [OfficeDev issue #5943](https://github.com/OfficeDev/office-js/issues/5943)
- Use Fluent UI React for dynamic taskpane UI. Only fall back to lighter-weight
  static styling when a surface is intentionally static and non-interactive.

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
production build, bundle budget checks, and deployable manifest validation.

Additional checks when relevant:

- `npm run validate:manifest:localhost` for local manifest work
- `npm run dev` plus manual **Add from File** sideload with `manifest-localhost.xml`
- `npm run start:desktop` if desktop auto-sideload behavior is the thing being changed
- `npm run test:host-smoke` for changes that affect Outlook compose flows, task
  pane behavior, Smart Alerts, selectors, or send-time rendering

If the host smoke cannot be run because credentials or Outlook test
infrastructure are unavailable, call that out explicitly in the PR.

## Release channel workflow

- MarkOut intentionally uses a **post-merge preview model**.
- There is no separate PR-preview host in the default delivery stack.
- GitHub Actions + GitHub Pages are the only supported hosted delivery path.
- `manifest.beta.xml` and `/outlook-beta/` are the post-merge preview/testing
  channel sourced from `main`.
- `manifest.xml` and `/outlook/` are the stable production channel sourced from
  `release/production`.
- Bootstrap `release/production` once from the current stable production commit
  before relying on independent promotions.
- Normal pushes to `main` must not move production.
- Production is updated only through the manual
  **Promote Production Channel** workflow by selecting a validated `main` SHA.
- The expected rollout path is:
  1. merge to `main`
  2. verify with `manifest.beta.xml`
  3. promote the validated commit
  4. verify with `manifest.xml`

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
