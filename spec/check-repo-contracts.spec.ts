import {
  checkEnglishDocumentationPolicy,
  checkDeployableManifestInvariants,
  checkOfficeMailManifestInvariants,
  checkPullRequestSupplyChainPolicy,
  checkReleaseBotDocumentationPolicy,
  type DocumentationPolicySnapshot,
  type ManifestSnapshot,
  type RuntimeChannelConfigSnapshot,
} from "../scripts/check-repo-contracts";

const runtimeChannelConfig: RuntimeChannelConfigSnapshot = {
  addInId: "05c2e1c9-3e1d-406e-9a91-e9ac64854143",
  appBaseUrl: "https://schoenfeld-solutions.github.io/markout/outlook",
  channelId: "production",
  commandsUrl:
    "https://schoenfeld-solutions.github.io/markout/outlook/commands.html?channel=production",
  launcheventUrl:
    "https://schoenfeld-solutions.github.io/markout/outlook/launchevent.js?channel=production",
  storageNamespace: "markout.production",
  supportUrl: "https://github.com/Schoenfeld-Solutions/markout",
  taskpaneUrl:
    "https://schoenfeld-solutions.github.io/markout/outlook/taskpane.html?channel=production",
};

function createManifestSnapshot(
  overrides: Partial<ManifestSnapshot> = {}
): ManifestSnapshot {
  const text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xsi:type="MailApp">
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.12" />
    </Sets>
  </Requirements>
  <Permissions>ReadWriteItem</Permissions>
  <ExtensionPoint xsi:type="MessageComposeCommandSurface">
  </ExtensionPoint>
  <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  </ExtensionPoint>
  <ExtensionPoint xsi:type="LaunchEvent">
    <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock" />
    <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="SoftBlock" />
  </ExtensionPoint>
</OfficeApp>`;

  return {
    appDomains: ["https://schoenfeld-solutions.github.io"],
    displayName: "MarkOut",
    highResolutionIconUrl:
      "https://raw.githubusercontent.com/Schoenfeld-Solutions/markout/main/assets/icon-128.png",
    iconUrl:
      "https://raw.githubusercontent.com/Schoenfeld-Solutions/markout/main/assets/icon-64.png",
    id: runtimeChannelConfig.addInId,
    path: "manifest.xml",
    sourceLocation: runtimeChannelConfig.taskpaneUrl,
    supportUrl: runtimeChannelConfig.supportUrl,
    text,
    urls: {
      "Commands.Url": runtimeChannelConfig.commandsUrl,
      "JSRuntime.Url": runtimeChannelConfig.launcheventUrl,
      "Taskpane.Url": runtimeChannelConfig.taskpaneUrl,
      "WebViewRuntime.Url": runtimeChannelConfig.commandsUrl,
    },
    version: "1.0.1.0",
    ...overrides,
  };
}

function createDocumentationPolicySnapshot(
  overrides: Partial<DocumentationPolicySnapshot> = {}
): DocumentationPolicySnapshot {
  return {
    agents:
      "Repository documentation, ADRs, runbooks, PR descriptions, code comments, and English source copy stay in English.\n" +
      "Product locale literals such as `de-DE`, the visible language label `Deutsch`, localized runtime strings, and proper nouns such as `Gabriel-Johannes Schönfeld` may appear only when documenting or implementing current localization behavior.",
    contributing:
      "Repository documentation, ADRs, runbooks, PR descriptions, code comments, and English source copy must be authored in English.\n" +
      "Product locale literals such as `de-DE`, the visible language label `Deutsch`, localized runtime strings, and proper nouns such as `Gabriel-Johannes Schönfeld` are allowed only when documenting or implementing current localization behavior.",
    readme:
      "Repository documentation, ADRs, runbooks, PR descriptions, code comments, and English source copy are authored in English.\n" +
      "Product locale literals such as `de-DE`, the visible language label `Deutsch`, localized runtime strings, and proper nouns such as `Gabriel-Johannes Schönfeld` may appear when documenting or implementing current localization behavior.",
    ...overrides,
  };
}

describe("repository contract manifest checks", () => {
  it("accepts the expected Office mail and deployable manifest invariants", () => {
    const snapshot = createManifestSnapshot();
    const contractErrors: string[] = [];

    checkOfficeMailManifestInvariants(snapshot, contractErrors);
    checkDeployableManifestInvariants(
      runtimeChannelConfig,
      snapshot,
      contractErrors
    );

    expect(contractErrors).toEqual([]);
  });

  it("rejects deployable manifests that leak local channel URLs", () => {
    const snapshot = createManifestSnapshot({
      sourceLocation: "https://localhost:3000/taskpane.html?channel=local",
      text: `${createManifestSnapshot().text}\nhttps://localhost:3000?channel=local`,
    });
    const contractErrors: string[] = [];

    checkDeployableManifestInvariants(
      runtimeChannelConfig,
      snapshot,
      contractErrors
    );

    expect(contractErrors).toEqual(
      expect.arrayContaining([
        "manifest.xml is deployable and must not reference localhost.",
        "manifest.xml is deployable and must not reference the local channel.",
      ])
    );
  });

  it("rejects missing Smart Alerts launch event wiring", () => {
    const snapshot = createManifestSnapshot({
      text: createManifestSnapshot().text.replace(
        '<LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock" />',
        ""
      ),
    });
    const contractErrors: string[] = [];

    checkOfficeMailManifestInvariants(snapshot, contractErrors);

    expect(contractErrors).toContain(
      "manifest.xml is missing message send Smart Alerts handler."
    );
  });

  it("rejects deployable manifest URLs outside the runtime path", () => {
    const snapshot = createManifestSnapshot({
      urls: {
        ...createManifestSnapshot().urls,
        "Taskpane.Url":
          "https://schoenfeld-solutions.github.io/markout/other/taskpane.html?channel=production",
      },
    });
    const contractErrors: string[] = [];

    checkDeployableManifestInvariants(
      runtimeChannelConfig,
      snapshot,
      contractErrors
    );

    expect(contractErrors).toContain(
      "manifest.xml Taskpane.Url must stay under `/markout/outlook/`."
    );
  });
});

describe("repository contract documentation policy checks", () => {
  it("accepts the English-only repository documentation policy", () => {
    const contractErrors: string[] = [];

    checkEnglishDocumentationPolicy(
      createDocumentationPolicySnapshot(),
      contractErrors
    );

    expect(contractErrors).toEqual([]);
  });

  it("allows product locale literals and proper nouns in English policy prose", () => {
    const contractErrors: string[] = [];

    checkEnglishDocumentationPolicy(
      createDocumentationPolicySnapshot({
        readme:
          "Repository documentation, ADRs, runbooks, PR descriptions, code comments, and English source copy are authored in English.\n" +
          "Product locale literals such as `de-DE`, the visible language label `Deutsch`, localized runtime strings, and proper nouns such as `Gabriel-Johannes Schönfeld` may appear when documenting or implementing current localization behavior.\n" +
          "The taskpane still documents the `de-DE` runtime locale and the `Deutsch` label in English.",
      }),
      contractErrors
    );

    expect(contractErrors).toEqual([]);
  });

  it("rejects German-default repository documentation rules", () => {
    const contractErrors: string[] = [];

    checkEnglishDocumentationPolicy(
      createDocumentationPolicySnapshot({
        contributing:
          "Repository documentation, ADRs, runbooks, PR descriptions, code comments, and English source copy must be authored in English.\n" +
          "Product locale literals such as `de-DE`, the visible language label `Deutsch`, localized runtime strings, and proper nouns such as `Gabriel-Johannes Schönfeld` are allowed only when documenting or implementing current localization behavior.\n" +
          "Markdown docs must stay in German.",
      }),
      contractErrors
    );

    expect(contractErrors).toContain(
      "CONTRIBUTING.md must not define German as the repository documentation language."
    );
  });

  it("rejects German-language documentation policy phrases", () => {
    const contractErrors: string[] = [];

    checkEnglishDocumentationPolicy(
      createDocumentationPolicySnapshot({
        agents:
          "Repository documentation, ADRs, runbooks, PR descriptions, code comments, and English source copy stay in English.\n" +
          "Product locale literals such as `de-DE`, the visible language label `Deutsch`, localized runtime strings, and proper nouns such as `Gabriel-Johannes Schönfeld` may appear only when documenting or implementing current localization behavior.\n" +
          "Dokumentation ist auf Deutsch zu schreiben.",
      }),
      contractErrors
    );

    expect(contractErrors).toContain(
      "AGENTS.md must not define German as the repository documentation language."
    );
  });
});

describe("repository contract pull request supply-chain policy checks", () => {
  const packageJson = {
    scripts: {
      "audit:ci": "npm audit --audit-level=moderate",
    },
    version: "1.0.1",
  };
  const workflowText = `
name: Pull Request Gates

jobs:
  validate-pr-title:
    runs-on: ubuntu-latest
  quality:
    needs: [validate-pr-title]
    steps:
      - name: Run supply-chain audit gate
        run: npm run audit:ci
`;

  it("accepts the repo-native npm audit PR gate", () => {
    const contractErrors: string[] = [];

    checkPullRequestSupplyChainPolicy(
      workflowText,
      packageJson,
      contractErrors
    );

    expect(contractErrors).toEqual([]);
  });

  it("rejects the node20 dependency-review action", () => {
    const contractErrors: string[] = [];

    checkPullRequestSupplyChainPolicy(
      `${workflowText}\n        uses: actions/dependency-review-action@v4\n`,
      packageJson,
      contractErrors
    );

    expect(contractErrors).toContain(
      ".github/workflows/pull-request.yaml must not use actions/dependency-review-action@v4 because it runs on node20."
    );
  });

  it("rejects missing audit scripts and separate dependency-review jobs", () => {
    const contractErrors: string[] = [];

    checkPullRequestSupplyChainPolicy(
      `
jobs:
  dependency-review:
    runs-on: ubuntu-latest
  quality:
    needs: [validate-pr-title, dependency-review]
`,
      { scripts: {}, version: "1.0.1" },
      contractErrors
    );

    expect(contractErrors).toEqual(
      expect.arrayContaining([
        "package.json must define `audit:ci` as `npm audit --audit-level=moderate`.",
        ".github/workflows/pull-request.yaml must run `npm run audit:ci` in the PR quality gate.",
        ".github/workflows/pull-request.yaml must not keep a separate dependency-review job.",
      ])
    );
  });
});

describe("repository contract release-bot documentation checks", () => {
  const releaseBotRunbook = `
Create a dedicated GitHub App named \`markout-release-bot\`.

- Repository permissions: \`Contents: Read and write\`
- Repository permissions: \`Workflows: Read and write\`

GitHub rejects those updates unless the App installation token has workflow
write permission, even when the App is the only ruleset bypass actor.
`;
  const productionPromotionRunbook = `
The \`markout-release-bot\` GitHub App has \`Contents: Read and write\` and
\`Workflows: Read and write\`, because promoted commits may include workflow
file changes.

If the push fails with \`refusing to allow a GitHub App to create or update
workflow\`, update the release-bot App permission.
`;

  it("accepts documented release-bot workflow write permission", () => {
    const contractErrors: string[] = [];

    checkReleaseBotDocumentationPolicy(
      releaseBotRunbook,
      productionPromotionRunbook,
      contractErrors
    );

    expect(contractErrors).toEqual([]);
  });

  it("rejects release-bot docs that omit workflow write permission", () => {
    const contractErrors: string[] = [];

    checkReleaseBotDocumentationPolicy(
      releaseBotRunbook.replace(
        "- Repository permissions: `Workflows: Read and write`",
        ""
      ),
      productionPromotionRunbook.replace(
        " and\n`Workflows: Read and write`",
        ""
      ),
      contractErrors
    );

    expect(contractErrors).toEqual(
      expect.arrayContaining([
        "docs/runbooks/release-bot-bootstrap.md is missing the required release-bot permission snippet: `Workflows: Read and write`",
        "docs/runbooks/production-promotion.md is missing the required release-bot permission snippet: `Contents: Read and write` and `Workflows: Read and write`",
      ])
    );
  });
});
