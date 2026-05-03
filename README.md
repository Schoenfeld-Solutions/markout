# MarkOut

**Safe Markdown authoring for Outlook compose**

MarkOut is an Outlook compose add-in that turns Markdown into inline-styled email HTML.
The add-in keeps the authoring experience focused on Markdown while giving users a task pane
for previewing the default theme, customizing the stylesheet, and enabling automatic rendering
through Outlook Smart Alerts.

## Compose UX

MarkOut is intentionally **taskpane-first** in Outlook compose.

- `Open MarkOut` is the single compose command and opens the task pane.
- Manual Markdown work happens inside the task pane: render the current body
  selection, render Markdown-looking blocks in the draft while preserving
  non-Markdown HTML such as signatures, or insert rendered Markdown fragments
  at the current body selection or cursor.
- The task pane includes light, dark, and system theme modes, first-run intro
  content, help, credits, developer tooling, and Smart Alerts settings.

MarkOut does **not** implement a native Outlook context-menu integration for
Markdown conversion. This is a deliberate product decision based on the current
Outlook web add-in platform constraints discussed in
[OfficeDev/office-js#2364](https://github.com/OfficeDev/office-js/issues/2364)
and
[OfficeDev/office-js#5943](https://github.com/OfficeDev/office-js/issues/5943).
The UI therefore follows the Microsoft Office add-in taskpane model instead of
trying to fake a host context menu.

## Localization and notifications

MarkOut currently supports `en-US` and `de-DE` in the taskpane runtime.

- The taskpane has a runtime language switcher with `Browser default`,
  `English`, and `Deutsch`.
- `Browser default` resolves from `Office.context.displayLanguage`, with
  `navigator.language` as a fallback and English as the final default.
- Add-in-only manifest strings are localized with `Override Locale="de-de"`
  entries for the visible Outlook labels and tooltips.
- When Smart Alerts auto-render is enabled, MarkOut uses Outlook informational
  notifications with `persistent: true`.
- Normal render, insert, and selection feedback is delivered through Outlook
  infobar notifications in the compose surface and is actively removed after a
  short lifetime.
- Pane-local message bars are reserved for taskpane-internal states such as
  stylesheet linting, editor failures, and notification API fallback.
- Developer tools include a pane-local diagnostic ring buffer for recent
  preview, selection, render, restore, and notification events. The buffer is
  in-memory only and stores operational codes and sizes, not draft content,
  selection text, tokens, recipient data, or session state.

Repository documentation, ADRs, runbooks, PR descriptions, code comments, and
English source copy are authored in English. Product locale literals such as
`de-DE`, the visible language label `Deutsch`, localized runtime strings, and
proper nouns such as `Gabriel-Johannes Schönfeld` may appear when documenting or
implementing current localization behavior.

References:

- [Localization for Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/develop/localization)
- [Create notifications for your Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/notifications)

## Fork and credits

MarkOut originated in the upstream project
[`SierraSoftworks/markout`](https://github.com/SierraSoftworks/markout) and continues to credit
that repository for the original product direction and implementation work.

This fork is now independently maintained by `Schoenfeld Solutions`, with ongoing maintenance by
`Gabriel-Johannes Schönfeld`. The repository also retains Microsoft Office Add-in scaffold code
that remains covered by the MIT license in [LICENSE](LICENSE).

The hosted app icon assets now derive from the Microsoft Fluent UI System Icons
Markdown glyph and follow the upstream MIT license terms as documented in the
Fluent icon project.

## Why this repo still uses the XML add-in manifest

MarkOut stays on the Outlook add-in-only XML manifest because Outlook on Mac still doesn't support
the unified Microsoft 365 manifest for this scenario. The repo therefore uses:

- `manifest.xml` for stable production deployments
- `manifest.beta.xml` for post-merge preview and testing deployments
- `manifest-localhost.xml` for local sideloading and development

Each manifest is now a distinct add-in with its own add-in ID and path-based
runtime channel.
No browser settings or restore-state keys are shared across production, beta, and local add-ins.

## Installation

Microsoft no longer supports installing Outlook add-ins with **Add from URL**. To sideload MarkOut
for testing, use one of the supported flows documented by Microsoft:

1. Open Outlook and go to **Get Add-ins** > **My add-ins** > **Add a custom add-in**.
2. Choose **Add from File**.
3. Select the manifest you want to test:
   - `manifest-localhost.xml` for local development
   - `manifest.xml` for the stable production deployment
   - `manifest.beta.xml` for the post-merge preview and testing channel

The command buttons appear in message compose and appointment compose surfaces. The task pane is used
for all manual rendering and insertion work, while Smart Alerts can auto-render content before send
when the user enables that option.

## Development

### Prerequisites

- Node.js 25 Current or Node.js 26 once available in your environment
- An Outlook client that supports Mailbox 1.12 Smart Alerts
- Microsoft 365 sideloading enabled for your tenant or account

MarkOut currently validates against Node 25 in CI and allows the next Node major without reopening
the repo policy immediately.
DOM-dependent tests use `jest-environment-jsdom` instead of a direct `jsdom`
dependency. The renderer accepts both the old function-style
`markdown-it-emoji` export shape and the v3 `{ full }` plugin export shape.

### Setup

```bash
npm ci
npm run dev
```

`npm run dev` starts the repo-native local development server for
`manifest-localhost.xml`. It runs webpack in watch mode and serves the generated
`dist/` assets over HTTPS without `webpack-dev-server`.
`npm run start` is kept as a convenience alias to the same local dev server.
Then sideload the add-in manually in Outlook with **Add from File** and select
`manifest-localhost.xml`.

On macOS you may be prompted to trust the local developer CA certificate before Outlook or your browser
will trust `https://localhost:3000`.

`npm run dev` requires local TLS material because Outlook sideloading requires
HTTPS. Set `MARKOUT_DEV_TLS_CERT_PATH` and `MARKOUT_DEV_TLS_KEY_PATH`, or keep
trusted files at the legacy default paths
`~/.office-addin-dev-certs/localhost.crt` and
`~/.office-addin-dev-certs/localhost.key`. The repository no longer generates or
trusts certificates for you; use your local OS tooling or an external tool such
as `mkcert` if you need to create a new trusted localhost certificate.

For Outlook desktop testing, keep `npm run dev` running and sideload
`manifest-localhost.xml` manually with **Add from File**. The repository does
not carry the deprecated Office auto-debugging CLI because it adds a large
transitive maintenance and audit surface for a workflow that is easy to perform
manually.

For hosted builds, download one of the published manifests first and then use
**Add from File**. Browsers may display GitHub Pages XML inline instead of
downloading it. Use the download buttons on
`https://schoenfeld-solutions.github.io/markout/`, use the browser's **Save
Page As** command, or download explicitly with `curl`:

```bash
curl -L -o manifest.xml https://schoenfeld-solutions.github.io/markout/manifest.xml
curl -L -o manifest.beta.xml https://schoenfeld-solutions.github.io/markout/manifest.beta.xml
```

Hosted channel semantics:

- `manifest.xml` installs as **MarkOut (Production)**, is the stable production
  channel, and is sourced from the `release/production` branch.
- `manifest.beta.xml` is the post-merge preview/testing channel and is sourced
  from `main`; it installs as **MarkOut (Beta)**.

The Pages root at `https://schoenfeld-solutions.github.io/markout/` now serves a static MarkOut
landing page instead of a generic GitHub 404.

The taskpane implementation is built with React and Fluent UI and follows the
Microsoft Office add-in design guidance for task panes and layout:

- [Office Add-in design language](https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-design-language)
- [Task panes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/task-pane-add-ins)
- [Layout guidelines for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-layout)
- [Use Fluent UI React in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/using-office-ui-fabric-react)

### Local commands

```bash
npm run dev
npm run format
npm run format:check
npm run lint
npm run typecheck
npm test
npm run test:taskpane-ui
npm run build
npm run check:bundle
npm run validate:manifest
npm run check:github-release-governance
npm run check
npm run check:ci
```

`npm run validate:manifest` performs offline repo-native contract checks for the
deployable production and beta manifests. It enforces the expected add-in IDs,
versions, channel URLs, HTTPS-only hosted URLs, Smart Alerts entry points, and
icon/support URL invariants without calling the deprecated Office validation
toolchain.
`npm run validate:manifest:localhost` is available for local manifest checks, but `manifest-localhost.xml`
is intentionally not a Marketplace-valid manifest because it targets `https://localhost:3000`.

`npm run test:taskpane-ui` starts the same local server in HTTP mode with a
test-only taskpane mock plus a minimal OWA-like drawer host. Playwright verifies
the isolated taskpane and the host iframe layout so short panels keep the bottom
toolbar pinned while only the content viewport scrolls. Use
`npm run test:taskpane-ui:headed` when a visible local browser check is needed;
it still does not automate Outlook Web or replace mandatory human OWA
verification.

`npm run check` is the local pre-merge gate. It runs formatting checks, linting, type checking, unit tests,
the production build, bundle budget checks, and deployable manifest contract validation.
`npm run check:ci` is the pull request gate variant. It keeps the same
formatting, linting, type, build, bundle, manifest, and repo-contract checks,
but runs the coverage test pass instead of a separate unit-test pass so GitHub
Actions does not execute Jest twice.

`npm run check:repo-contracts` enforces manifest/version/channel invariants and
release-policy documentation drift.
`npm run check:github-release-governance` audits the GitHub settings required
for production promotion when a settings-read token is available.
Pull requests also require a Conventional Commit PR title, dependency review,
`npm run test:taskpane-ui`, and the coverage-backed `npm run check:ci` gate in
GitHub Actions.

## Release channels

MarkOut follows a deliberate **post-merge preview model** on GitHub Pages.

- `main` is the integration branch for the hosted beta channel.
- `release/production` is the stable source branch for the hosted production
  channel.
- GitHub Actions packages both channels onto the same GitHub Pages site:
  - `/markout/outlook/` from `release/production`
  - `/markout/outlook-beta/` from `main`
- Normal pushes to `main` must not silently move production.
- Promotion from beta to production happens manually through the
  **Promote Production Channel** workflow by choosing a validated `main` commit.
- Release packaging hard-fails if `release/production` is missing.
- Promotion requires a successful beta release run for the exact `main` SHA and
  explicit manual OWA beta verification for that SHA.
- `release/production` is intended to be automation-only. If GitHub-native
  rules cannot block human pushes while preserving promotion automation, the
  `markout-release-bot` GitHub App is used with repository-only
  `Contents: write` access and no external runtime.
- The scheduled **GitHub Settings Audit** workflow checks branch, environment,
  Pages policy, release-bot, and production ruleset drift.

The expected testing flow is therefore:

1. Review and merge the PR into `main`.
2. Install `manifest.beta.xml` in OWA with **Add from File**.
3. Verify the new hosted beta channel in compose.
4. Promote the validated `main` commit to `release/production`.
5. Verify the stable production channel with `manifest.xml`.

Architecture decisions and operating procedures for this flow live in:

- [`docs/adr/0001-channel-isolation.md`](docs/adr/0001-channel-isolation.md)
- [`docs/adr/0002-taskpane-boundaries.md`](docs/adr/0002-taskpane-boundaries.md)
- [`docs/adr/0003-release-gates.md`](docs/adr/0003-release-gates.md)
- [`docs/runbooks/10-10-continuation.md`](docs/runbooks/10-10-continuation.md)
- [`docs/runbooks/beta-verification.md`](docs/runbooks/beta-verification.md)
- [`docs/runbooks/release-bot-bootstrap.md`](docs/runbooks/release-bot-bootstrap.md)
- [`docs/runbooks/production-promotion.md`](docs/runbooks/production-promotion.md)
- [`docs/runbooks/rollback.md`](docs/runbooks/rollback.md)
- [`docs/runbooks/host-smoke-failures.md`](docs/runbooks/host-smoke-failures.md)

## Dependency maintenance

- Dependabot version updates are grouped into one weekly tooling PR for `npm` and `github-actions`.
- Security updates are intentionally kept separate from normal version updates.
- Pull request supply-chain blocking is repo-native: PR CI runs
  `npm run audit:ci`, which fails on moderate or higher npm advisories without
  the Node 20 runtime warning from the legacy dependency-review action.
- Because this repository is a fork, GitHub may disable Dependabot security updates by default for the fork.
  Enable both **Dependabot security updates** and **Grouped security updates** in the repository settings
  if you want grouped security PRs to be created here.
- Contributor workflow, commit rules, and required local gates are documented in [CONTRIBUTING.md](CONTRIBUTING.md).

## Manual OWA verification

OWA verification is deliberately human-confirmed. GitHub Actions must not store
Outlook Web storage-state JSON, scheduled OWA test credentials, or a permanent
test mailbox secret.

MarkOut keeps an env-guarded Outlook on the web smoke script at
`npm run test:host-smoke` for human-initiated diagnostics only. It is not a
release CI gate. The script requires:

- a browser executable such as Chrome or Chromium
- a Playwright storage state JSON file for an already authenticated Outlook test account
- an Outlook compose URL where the add-in is available
- a dedicated recipient mailbox for the send-flow check

Required environment variables:

- `MARKOUT_HOST_SMOKE_STORAGE_STATE`
- `MARKOUT_HOST_SMOKE_RECIPIENT`

Production promotion requires a human to confirm beta verification in OWA for
the exact `main` SHA being promoted. The `production-promotion` environment
approval and the `beta_verification_confirmed` workflow input are the durable
promotion evidence.

Common optional overrides:

- `MARKOUT_HOST_SMOKE_COMPOSE_URL`
- `MARKOUT_HOST_SMOKE_AUTORENDER_SWITCH_SELECTOR`
- `MARKOUT_HOST_SMOKE_BROWSER_EXECUTABLE`
- `MARKOUT_HOST_SMOKE_EXPECTED_TASKPANE_URL_PREFIX`
- `MARKOUT_HOST_SMOKE_INSERT_PANEL_BUTTON_SELECTOR`
- `MARKOUT_HOST_SMOKE_INTRO_CONFIRM_BUTTON_SELECTOR`
- `MARKOUT_HOST_SMOKE_INTRO_PANEL_BUTTON_SELECTOR`
- `MARKOUT_HOST_SMOKE_OPEN_BUTTON_SELECTOR`
- `MARKOUT_HOST_SMOKE_MESSAGE_BODY_SELECTOR`
- `MARKOUT_HOST_SMOKE_RENDER_BUTTON_SELECTOR`
- `MARKOUT_HOST_SMOKE_SETTINGS_PANEL_BUTTON_SELECTOR`
- `MARKOUT_HOST_SMOKE_TASKPANE_FRAME_SELECTOR`
- `MARKOUT_HOST_SMOKE_TASKPANE_READY_SELECTOR`
- `MARKOUT_HOST_SMOKE_SEND_BUTTON_SELECTOR`

The smoke verifies that the task pane opens, the preview loads, auto-render
remains enabled after reload, manual rendering updates the draft body, and the
Smart Alerts send flow completes successfully.

`MARKOUT_HOST_SMOKE_EXPECTED_TASKPANE_URL_PREFIX` can be used to assert that the
opened taskpane iframe comes from the intended hosted channel:

- `https://schoenfeld-solutions.github.io/markout/outlook-beta/taskpane.html`
  for post-merge beta verification
- `https://schoenfeld-solutions.github.io/markout/outlook/taskpane.html`
  for stable production verification

## Deployment notes

- Support and issue tracking live at `https://github.com/Schoenfeld-Solutions/markout`.
- The published GitHub Pages site is `https://schoenfeld-solutions.github.io/markout/`.
- The Pages root serves a static MarkOut landing page with download links to hosted manifests, task pane runtimes, and the repository.
- Unknown GitHub Pages paths use a custom MarkOut 404 page instead of the default GitHub Pages error screen.
- Production assets are served from `https://schoenfeld-solutions.github.io/markout/outlook/` and sourced from `release/production`.
- Beta assets are served from `https://schoenfeld-solutions.github.io/markout/outlook-beta/` and sourced from `main`.
- Manifest variants must remain behaviorally aligned.

[gfm]: https://github.github.com/gfm/
