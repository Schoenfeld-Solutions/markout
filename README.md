# MarkOut

**Safe Markdown authoring for Outlook compose**

MarkOut is an Outlook compose add-in that turns Markdown into inline-styled email HTML.
The add-in keeps the authoring experience focused on Markdown while giving users a task pane
for previewing the default theme, customizing the stylesheet, and enabling automatic rendering
through Outlook Smart Alerts.

## Fork and credits

MarkOut originated in the upstream project
[`SierraSoftworks/markout`](https://github.com/SierraSoftworks/markout) and continues to credit
that repository for the original product direction and implementation work.

This fork is now independently maintained by `Schoenfeld Solutions`, with ongoing maintenance by
`Gabriel-Johannes Schönfeld`. The repository also retains Microsoft Office Add-in scaffold code
that remains covered by the MIT license in [LICENSE](LICENSE).

## Why this repo still uses the XML add-in manifest

MarkOut stays on the Outlook add-in-only XML manifest because Outlook on Mac still doesn't support
the unified Microsoft 365 manifest for this scenario. The repo therefore uses:

- `manifest.xml` for production deployments
- `manifest.beta.xml` for staging deployments
- `manifest-localhost.xml` for local sideloading and development

## Installation

Microsoft no longer supports installing Outlook add-ins with **Add from URL**. To sideload MarkOut
for testing, use one of the supported flows documented by Microsoft:

1. Open Outlook and go to **Get Add-ins** > **My add-ins** > **Add a custom add-in**.
2. Choose **Add from File**.
3. Select the manifest you want to test:
   - `manifest-localhost.xml` for local development
   - `manifest.xml` for the production deployment

The command buttons appear in message compose and appointment compose surfaces. The task pane is used
for theme editing and manual rendering, while Smart Alerts can auto-render content before send when
the user enables that option.

## Development

### Prerequisites

- Node.js 25 Current or Node.js 26 once available in your environment
- An Outlook client that supports Mailbox 1.12 Smart Alerts
- Microsoft 365 sideloading enabled for your tenant or account

MarkOut currently validates against Node 25 in CI and allows the next Node major without reopening
the repo policy immediately.
`jsdom` stays on the latest green 21.x line because newer majors break the current Jest stack,
and `markdown-it-emoji` stays on 2.x because v3 changes the plugin shape used by the renderer.

### Setup

```bash
npm ci
npm run dev
```

`npm run dev` starts the local HTTPS webpack dev server for `manifest-localhost.xml`.
`npm run start` is kept as a convenience alias to the same local dev server.
Then sideload the add-in manually in Outlook with **Add from File** and select `manifest-localhost.xml`.

On macOS you may be prompted to trust the local developer CA certificate before Outlook or your browser
will trust `https://localhost:3000`.

Use `npm run start:desktop` if you want to target the Outlook desktop client instead.

For hosted builds, download one of the published manifests first and then use **Add from File**:

- `https://schoenfeld-solutions.github.io/markout/manifest.xml`
- `https://schoenfeld-solutions.github.io/markout/manifest.beta.xml`

### Local commands

```bash
npm run dev
npm run format
npm run format:check
npm run lint
npm run typecheck
npm test
npm run build
npm run validate:manifest
npm run check
```

`npm run validate:manifest:localhost` is available for local manifest checks, but `manifest-localhost.xml`
is intentionally not a Marketplace-valid manifest because it targets `https://localhost:3000`.

`npm run check` is the local pre-merge gate. It runs formatting checks, linting, type checking, unit tests,
the production build, and deployable manifest validation.

## Dependency maintenance

- Dependabot version updates are grouped into one weekly tooling PR for `npm` and `github-actions`.
- Security updates are intentionally kept separate from normal version updates.
- Because this repository is a fork, GitHub may disable Dependabot security updates by default for the fork.
  Enable both **Dependabot security updates** and **Grouped security updates** in the repository settings
  if you want grouped security PRs to be created here.
- Contributor workflow, commit rules, and required local gates are documented in [CONTRIBUTING.md](CONTRIBUTING.md).

## Outlook host smoke

MarkOut includes an env-guarded Outlook on the web smoke script at
`npm run test:host-smoke`. The smoke is intentionally separate from the normal unit suite because it requires:

- a browser executable such as Chrome or Chromium
- a Playwright storage state JSON file for an already authenticated Outlook test account
- an Outlook compose URL where the add-in is available
- a dedicated recipient mailbox for the send-flow check

Required environment variables:

- `MARKOUT_HOST_SMOKE_STORAGE_STATE`
- `MARKOUT_HOST_SMOKE_RECIPIENT`

Common optional overrides:

- `MARKOUT_HOST_SMOKE_COMPOSE_URL`
- `MARKOUT_HOST_SMOKE_BROWSER_EXECUTABLE`
- `MARKOUT_HOST_SMOKE_OPEN_BUTTON_SELECTOR`
- `MARKOUT_HOST_SMOKE_MESSAGE_BODY_SELECTOR`
- `MARKOUT_HOST_SMOKE_TASKPANE_FRAME_SELECTOR`
- `MARKOUT_HOST_SMOKE_SEND_BUTTON_SELECTOR`

The smoke verifies that the task pane opens, the preview loads, auto-render remains enabled after reload,
manual rendering updates the draft body, and the Smart Alerts send flow completes successfully.

## Deployment notes

- Support and issue tracking live at `https://github.com/Schoenfeld-Solutions/markout`.
- The published GitHub Pages site is `https://schoenfeld-solutions.github.io/markout/`.
- Production assets are served from `https://schoenfeld-solutions.github.io/markout/outlook/`.
- Staging assets are served from `https://schoenfeld-solutions.github.io/markout/outlook-beta/`.
- Manifest variants must remain behaviorally aligned.

[gfm]: https://github.github.com/gfm/
