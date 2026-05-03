import { readFile } from "fs/promises";
import path from "path";
import { pathToFileURL } from "url";

export type ChannelId = "beta" | "local" | "production";

export interface RuntimeChannelConfigSnapshot {
  addInId: string;
  appBaseUrl: string;
  channelId: ChannelId;
  commandsUrl: string;
  launcheventUrl: string;
  storageNamespace: string;
  supportUrl: string;
  taskpaneUrl: string;
}

export interface ManifestSnapshot {
  appDomains: string[];
  displayName: string;
  highResolutionIconUrl: string;
  iconUrl: string;
  id: string;
  path: string;
  sourceLocation: string;
  supportUrl: string;
  text: string;
  urls: Record<string, string>;
  version: string;
}

interface ExpectedManifestContract {
  displayName: string;
  manifestPath: string;
}

export interface DocumentationPolicySnapshot {
  agents?: string;
  contributing: string;
  readme: string;
}

export interface PackageJsonSnapshot {
  scripts?: Record<string, string>;
  version: string;
}

const REPOSITORY_ROOT = process.cwd();
const DEPLOYABLE_MANIFEST_CHANNELS = new Set<ChannelId>(["beta", "production"]);
const deployableManifestsOnly = process.argv.includes(
  "--deployable-manifests-only"
);
const EXPECTED_RELEASE_POLICY_SNIPPETS = [
  {
    file: "README.md",
    snippet: "`main` is the integration branch for the hosted beta channel.",
  },
  {
    file: "README.md",
    snippet:
      "`release/production` is the stable source branch for the hosted production channel.",
  },
  {
    file: "README.md",
    snippet:
      "No browser settings or restore-state keys are shared across production, beta, and local add-ins.",
  },
  {
    file: "README.md",
    snippet: "OWA verification is deliberately human-confirmed.",
  },
  {
    file: "README.md",
    snippet:
      "The scheduled **GitHub Settings Audit** workflow checks branch, environment, Pages policy, release-bot, and production ruleset drift.",
  },
  {
    file: "CONTRIBUTING.md",
    snippet:
      "`manifest.beta.xml` and `/outlook-beta/` are the post-merge preview/testing channel sourced from `main`.",
  },
  {
    file: "CONTRIBUTING.md",
    snippet:
      "`manifest.xml` and `/outlook/` are the stable production channel sourced from `release/production`.",
  },
  {
    file: "CONTRIBUTING.md",
    snippet:
      "The `Promote Production Channel` workflow must push `release/production` with the release bot token once the automation-only ruleset is active.",
  },
  {
    file: "CONTRIBUTING.md",
    snippet:
      "Do not turn OWA checks into scheduled GitHub Actions or release CI gates.",
  },
];

const EXPECTED_ENGLISH_DOCUMENTATION_POLICY_SNIPPETS = [
  {
    file: "CONTRIBUTING.md",
    snippet:
      "Repository documentation, ADRs, runbooks, PR descriptions, code comments, and English source copy must be authored in English.",
  },
  {
    file: "CONTRIBUTING.md",
    snippet:
      "Product locale literals such as `de-DE`, the visible language label `Deutsch`, localized runtime strings, and proper nouns such as `Gabriel-Johannes Schönfeld` are allowed only when documenting or implementing current localization behavior.",
  },
  {
    file: "README.md",
    snippet:
      "Repository documentation, ADRs, runbooks, PR descriptions, code comments, and English source copy are authored in English.",
  },
  {
    file: "README.md",
    snippet:
      "Product locale literals such as `de-DE`, the visible language label `Deutsch`, localized runtime strings, and proper nouns such as `Gabriel-Johannes Schönfeld` may appear when documenting or implementing current localization behavior.",
  },
];

const GERMAN_DOCUMENTATION_POLICY_PATTERNS = [
  /\bGerman\s+(?:by default|default|required|mandatory|only|first)\b/iu,
  /\b(?:docs|documentation|Markdown docs|Markdown documentation|ADRs|runbooks)\s+(?:stay|must stay|are|must be|should be|default to)\s+(?:in\s+)?German\b/iu,
  /\b(?:German|Deutsch)\s+(?:docs|documentation|ADRs|runbooks)\s+(?:by default|required|mandatory|only|first)\b/iu,
  /\b(?:Dokumentation|Markdown-Doku|ADRs|Runbooks)[^.]*\b(?:Deutsch|deutsch)\b/iu,
  /\b(?:Deutsch|deutsch)[^.]*\b(?:Dokumentation|Markdown-Doku|ADRs|Runbooks)\b/iu,
];

const EXPECTED_WORKFLOW_SNIPPETS = [
  {
    file: ".github/workflows/release.yaml",
    snippet:
      "Missing origin/release/production. Refusing to publish GitHub Pages without an explicit production source branch.",
  },
  {
    file: ".github/workflows/promote-production.yaml",
    snippet: "production-promotion",
  },
  {
    file: ".github/workflows/promote-production.yaml",
    snippet: "Build and Publish GitHub Pages",
  },
  {
    file: ".github/workflows/promote-production.yaml",
    snippet: "MARKOUT_RELEASE_BOT_APP_ID",
  },
  {
    file: ".github/workflows/promote-production.yaml",
    snippet: "beta_verification_confirmed",
  },
  {
    file: ".github/workflows/github-settings-audit.yaml",
    snippet: "npm run check:github-release-governance",
  },
  {
    file: "docs/runbooks/10-10-continuation.md",
    snippet: "Resume entry point",
  },
  {
    file: "docs/runbooks/release-bot-bootstrap.md",
    snippet: "markout-release-bot",
  },
];

const EXPECTED_MANIFESTS: Record<ChannelId, ExpectedManifestContract> = {
  beta: {
    displayName: "MarkOut (Beta)",
    manifestPath: "manifest.beta.xml",
  },
  local: {
    displayName: "MarkOut (Local)",
    manifestPath: "manifest-localhost.xml",
  },
  production: {
    displayName: "MarkOut",
    manifestPath: "manifest.xml",
  },
};

async function main(): Promise<void> {
  const packageJson = await readJson<PackageJsonSnapshot>("package.json");
  const { getAllRuntimeChannelConfigs } = (await import(
    pathToFileURL(path.join(REPOSITORY_ROOT, "src/lib/runtime.ts")).href
  )) as {
    getAllRuntimeChannelConfigs: () => RuntimeChannelConfigSnapshot[];
  };

  const contractErrors: string[] = [];
  const expectedManifestVersion = `${packageJson.version}.0`;
  const runtimeChannelConfigs = getAllRuntimeChannelConfigs().filter(
    (runtimeChannelConfig) =>
      !deployableManifestsOnly ||
      DEPLOYABLE_MANIFEST_CHANNELS.has(runtimeChannelConfig.channelId)
  );
  const manifests = await Promise.all(
    runtimeChannelConfigs.map(async (runtimeChannelConfig) => {
      const expectedManifest =
        EXPECTED_MANIFESTS[runtimeChannelConfig.channelId];
      return {
        expectedManifest,
        runtimeChannelConfig,
        snapshot: await readManifest(expectedManifest.manifestPath),
      };
    })
  );

  const uniqueAddInIds = new Set<string>();

  for (const {
    expectedManifest,
    runtimeChannelConfig,
    snapshot,
  } of manifests) {
    uniqueAddInIds.add(snapshot.id);

    checkOfficeMailManifestInvariants(snapshot, contractErrors);
    if (DEPLOYABLE_MANIFEST_CHANNELS.has(runtimeChannelConfig.channelId)) {
      checkDeployableManifestInvariants(
        runtimeChannelConfig,
        snapshot,
        contractErrors
      );
    }

    if (snapshot.displayName !== expectedManifest.displayName) {
      contractErrors.push(
        `${snapshot.path} should use display name \`${expectedManifest.displayName}\`, found \`${snapshot.displayName}\`.`
      );
    }

    if (snapshot.id !== runtimeChannelConfig.addInId) {
      contractErrors.push(
        `${snapshot.path} add-in ID does not match src/lib/runtime.ts for channel \`${runtimeChannelConfig.channelId}\`.`
      );
    }

    if (snapshot.supportUrl !== runtimeChannelConfig.supportUrl) {
      contractErrors.push(
        `${snapshot.path} SupportUrl must match the runtime support URL for channel \`${runtimeChannelConfig.channelId}\`.`
      );
    }

    if (snapshot.version !== expectedManifestVersion) {
      contractErrors.push(
        `${snapshot.path} version \`${snapshot.version}\` does not match package.json version \`${expectedManifestVersion}\`.`
      );
    }

    if (snapshot.sourceLocation !== runtimeChannelConfig.taskpaneUrl) {
      contractErrors.push(
        `${snapshot.path} SourceLocation must match the runtime taskpane URL for channel \`${runtimeChannelConfig.channelId}\`.`
      );
    }

    const expectedUrls: Record<string, string> = {
      "Commands.Url": runtimeChannelConfig.commandsUrl,
      "JSRuntime.Url": runtimeChannelConfig.launcheventUrl,
      "Taskpane.Url": runtimeChannelConfig.taskpaneUrl,
      "WebViewRuntime.Url": runtimeChannelConfig.commandsUrl,
    };

    for (const [urlId, expectedUrl] of Object.entries(expectedUrls)) {
      if (snapshot.urls[urlId] !== expectedUrl) {
        contractErrors.push(
          `${snapshot.path} ${urlId} must equal \`${expectedUrl}\`.`
        );
      }
    }

    const runtimeOrigin = new URL(runtimeChannelConfig.appBaseUrl).origin;
    if (!snapshot.appDomains.includes(runtimeOrigin)) {
      contractErrors.push(
        `${snapshot.path} AppDomains must include \`${runtimeOrigin}\`.`
      );
    }
  }

  if (uniqueAddInIds.size !== manifests.length) {
    contractErrors.push("Each manifest must use a distinct add-in ID.");
  }

  if (!deployableManifestsOnly) {
    const [readme, contributing, agents] = await Promise.all([
      readText("README.md"),
      readText("CONTRIBUTING.md"),
      readOptionalText("AGENTS.md"),
    ]);
    checkSnippetPresence(readme, contributing, contractErrors);
    checkEnglishDocumentationPolicy(
      agents === undefined
        ? { contributing, readme }
        : { agents, contributing, readme },
      contractErrors
    );
    await checkWorkflowSnippets(contractErrors);
    checkSiteManifestDownloadLinks(
      await readText("site/index.html"),
      await readText("site/404.html"),
      contractErrors
    );
    checkReleaseBotDocumentationPolicy(
      await readText("docs/runbooks/release-bot-bootstrap.md"),
      await readText("docs/runbooks/production-promotion.md"),
      contractErrors
    );
    checkPullRequestSupplyChainPolicy(
      await readText(".github/workflows/pull-request.yaml"),
      packageJson,
      contractErrors
    );
  }

  if (contractErrors.length > 0) {
    console.error("MarkOut repository contract check failed:");
    for (const error of contractErrors) {
      console.error(`- ${error}`);
    }
    process.exit(1);
  }

  console.log(
    deployableManifestsOnly
      ? "MarkOut deployable manifest check passed."
      : "MarkOut repository contract check passed."
  );
}

export function checkOfficeMailManifestInvariants(
  snapshot: ManifestSnapshot,
  contractErrors: string[]
): void {
  const requiredSnippets = [
    ["OfficeApp root type", 'xsi:type="MailApp"'],
    ["Mailbox host", '<Host Name="Mailbox" />'],
    ["Mailbox 1.12 requirement", '<Set Name="Mailbox" MinVersion="1.12" />'],
    ["Read/write item permission", "<Permissions>ReadWriteItem</Permissions>"],
    [
      "message compose command surface",
      '<ExtensionPoint xsi:type="MessageComposeCommandSurface">',
    ],
    [
      "appointment organizer command surface",
      '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">',
    ],
    ["LaunchEvent extension point", '<ExtensionPoint xsi:type="LaunchEvent">'],
    [
      "message send Smart Alerts handler",
      '<LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock" />',
    ],
    [
      "appointment send Smart Alerts handler",
      '<LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="SoftBlock" />',
    ],
  ] as const;

  if (!snapshot.text.trimStart().startsWith('<?xml version="1.0"')) {
    contractErrors.push(`${snapshot.path} must start with an XML declaration.`);
  }

  if (!snapshot.text.trimEnd().endsWith("</OfficeApp>")) {
    contractErrors.push(`${snapshot.path} must end with </OfficeApp>.`);
  }

  for (const [label, snippet] of requiredSnippets) {
    if (!snapshot.text.includes(snippet)) {
      contractErrors.push(`${snapshot.path} is missing ${label}.`);
    }
  }
}

export function checkDeployableManifestInvariants(
  runtimeChannelConfig: RuntimeChannelConfigSnapshot,
  snapshot: ManifestSnapshot,
  contractErrors: string[]
): void {
  if (snapshot.text.includes("https://localhost")) {
    contractErrors.push(
      `${snapshot.path} is deployable and must not reference localhost.`
    );
  }

  if (snapshot.text.includes("channel=local")) {
    contractErrors.push(
      `${snapshot.path} is deployable and must not reference the local channel.`
    );
  }

  const urlsToValidate = [
    ["SourceLocation", snapshot.sourceLocation],
    ["SupportUrl", snapshot.supportUrl],
    ["IconUrl", snapshot.iconUrl],
    ["HighResolutionIconUrl", snapshot.highResolutionIconUrl],
    ...Object.entries(snapshot.urls),
  ] as const;

  for (const [label, url] of urlsToValidate) {
    let parsedUrl: URL;
    try {
      parsedUrl = new URL(url);
    } catch {
      contractErrors.push(`${snapshot.path} ${label} must be a valid URL.`);
      continue;
    }

    if (parsedUrl.protocol !== "https:") {
      contractErrors.push(`${snapshot.path} ${label} must use HTTPS.`);
    }

    if (parsedUrl.search.length > 0) {
      contractErrors.push(
        `${snapshot.path} ${label} must not include query strings in deployable manifests.`
      );
    }
  }

  for (const [label, url] of [
    ["IconUrl", snapshot.iconUrl],
    ["HighResolutionIconUrl", snapshot.highResolutionIconUrl],
  ] as const) {
    if (!url.startsWith("https://raw.githubusercontent.com/")) {
      contractErrors.push(
        `${snapshot.path} ${label} must use the raw GitHub icon host so validation works before Pages deploys.`
      );
    }

    if (!url.endsWith(".png")) {
      contractErrors.push(`${snapshot.path} ${label} must reference a PNG.`);
    }
  }

  const appBaseUrl = new URL(runtimeChannelConfig.appBaseUrl);
  for (const [urlId, url] of Object.entries(snapshot.urls)) {
    const parsedUrl = new URL(url);
    if (parsedUrl.origin !== appBaseUrl.origin) {
      contractErrors.push(
        `${snapshot.path} ${urlId} must use runtime origin \`${appBaseUrl.origin}\`.`
      );
    }

    if (!parsedUrl.pathname.startsWith(`${appBaseUrl.pathname}/`)) {
      contractErrors.push(
        `${snapshot.path} ${urlId} must stay under \`${appBaseUrl.pathname}/\`.`
      );
    }
  }
}

function checkSnippetPresence(
  readme: string,
  contributing: string,
  contractErrors: string[]
): void {
  const fileContents: Record<string, string> = {
    "CONTRIBUTING.md": normalizeWhitespace(contributing),
    "README.md": normalizeWhitespace(readme),
  };

  for (const { file, snippet } of EXPECTED_RELEASE_POLICY_SNIPPETS) {
    if (!fileContents[file]?.includes(normalizeWhitespace(snippet))) {
      contractErrors.push(
        `${file} is missing the required release-policy snippet: ${snippet}`
      );
    }
  }
}

export function checkEnglishDocumentationPolicy(
  snapshot: DocumentationPolicySnapshot,
  contractErrors: string[]
): void {
  const fileContents: Record<string, string> = {
    "CONTRIBUTING.md": normalizeWhitespace(snapshot.contributing),
    "README.md": normalizeWhitespace(snapshot.readme),
  };

  if (snapshot.agents !== undefined) {
    fileContents["AGENTS.md"] = normalizeWhitespace(snapshot.agents);
  }

  for (const {
    file,
    snippet,
  } of EXPECTED_ENGLISH_DOCUMENTATION_POLICY_SNIPPETS) {
    if (!fileContents[file]?.includes(normalizeWhitespace(snippet))) {
      contractErrors.push(
        `${file} is missing the required English documentation policy snippet: ${snippet}`
      );
    }
  }

  for (const [file, content] of Object.entries(fileContents)) {
    for (const pattern of GERMAN_DOCUMENTATION_POLICY_PATTERNS) {
      if (pattern.test(content)) {
        contractErrors.push(
          `${file} must not define German as the repository documentation language.`
        );
      }
    }
  }
}

export function checkPullRequestSupplyChainPolicy(
  workflowText: string,
  packageJson: PackageJsonSnapshot,
  contractErrors: string[]
): void {
  const normalizedWorkflowText = normalizeWhitespace(workflowText);
  const auditScript = packageJson.scripts?.["audit:ci"];

  if (auditScript !== "npm audit --audit-level=moderate") {
    contractErrors.push(
      "package.json must define `audit:ci` as `npm audit --audit-level=moderate`."
    );
  }

  if (!normalizedWorkflowText.includes("npm run audit:ci")) {
    contractErrors.push(
      ".github/workflows/pull-request.yaml must run `npm run audit:ci` in the PR quality gate."
    );
  }

  if (/actions\/dependency-review-action@v4\b/u.test(workflowText)) {
    contractErrors.push(
      ".github/workflows/pull-request.yaml must not use actions/dependency-review-action@v4 because it runs on node20."
    );
  }

  if (/(^|\n)\s{2}dependency-review:\s*(?:\n|$)/u.test(workflowText)) {
    contractErrors.push(
      ".github/workflows/pull-request.yaml must not keep a separate dependency-review job."
    );
  }
}

export function checkSiteManifestDownloadLinks(
  siteIndexText: string,
  siteNotFoundText: string,
  contractErrors: string[]
): void {
  const pages: Record<string, string> = {
    "site/404.html": siteNotFoundText,
    "site/index.html": siteIndexText,
  };

  for (const [file, content] of Object.entries(pages)) {
    for (const manifestPath of ["manifest.xml", "manifest.beta.xml"] as const) {
      if (!hasManifestDownloadLink(content, manifestPath)) {
        contractErrors.push(
          `${file} must include a same-origin download link for ${manifestPath}.`
        );
      }
    }
  }
}

export function checkReleaseBotDocumentationPolicy(
  releaseBotRunbookText: string,
  productionPromotionRunbookText: string,
  contractErrors: string[]
): void {
  const releaseBotRunbook = normalizeWhitespace(releaseBotRunbookText);
  const productionPromotionRunbook = normalizeWhitespace(
    productionPromotionRunbookText
  );

  const requiredReleaseBotSnippets = [
    "`Contents: Read and write`",
    "`Workflows: Read and write`",
    "GitHub rejects those updates unless the App installation token has workflow write permission",
  ];
  const requiredPromotionSnippets = [
    "`Contents: Read and write` and `Workflows: Read and write`",
    "refusing to allow a GitHub App to create or update workflow",
  ];

  for (const snippet of requiredReleaseBotSnippets) {
    if (!releaseBotRunbook.includes(normalizeWhitespace(snippet))) {
      contractErrors.push(
        `docs/runbooks/release-bot-bootstrap.md is missing the required release-bot permission snippet: ${snippet}`
      );
    }
  }

  for (const snippet of requiredPromotionSnippets) {
    if (!productionPromotionRunbook.includes(normalizeWhitespace(snippet))) {
      contractErrors.push(
        `docs/runbooks/production-promotion.md is missing the required release-bot permission snippet: ${snippet}`
      );
    }
  }
}

function hasManifestDownloadLink(
  content: string,
  manifestPath: string
): boolean {
  const escapedManifestPath = escapeRegExp(manifestPath);
  return new RegExp(
    `<a\\b(?=[^>]*\\bhref="${escapedManifestPath}")(?=[^>]*\\bdownload="${escapedManifestPath}")[^>]*>`,
    "u"
  ).test(content);
}

async function checkWorkflowSnippets(contractErrors: string[]): Promise<void> {
  for (const { file, snippet } of EXPECTED_WORKFLOW_SNIPPETS) {
    const workflowText = normalizeWhitespace(await readText(file));
    if (!workflowText.includes(normalizeWhitespace(snippet))) {
      contractErrors.push(
        `${file} is missing the required workflow contract snippet: ${snippet}`
      );
    }
  }
}

async function readJson<T>(relativePath: string): Promise<T> {
  const fileContent = await readText(relativePath);
  return JSON.parse(fileContent) as T;
}

async function readManifest(relativePath: string): Promise<ManifestSnapshot> {
  const manifestText = await readText(relativePath);
  const urls: Record<string, string> = {};

  for (const urlId of [
    "Commands.Url",
    "JSRuntime.Url",
    "Taskpane.Url",
    "WebViewRuntime.Url",
  ]) {
    urls[urlId] = readRequiredMatch(
      manifestText,
      new RegExp(
        `<bt:Url id="${escapeRegExp(urlId)}" DefaultValue="([^"]+)"`,
        "u"
      ),
      `${relativePath} ${urlId}`
    );
  }

  const appDomains = Array.from(
    manifestText.matchAll(/<AppDomain>([^<]+)<\/AppDomain>/gu),
    (match) => match[1]
  ).filter((value): value is string => value !== undefined);

  return {
    appDomains,
    displayName: readRequiredMatch(
      manifestText,
      /<DisplayName DefaultValue="([^"]+)"/u,
      `${relativePath} DisplayName`
    ),
    highResolutionIconUrl: readRequiredMatch(
      manifestText,
      /<HighResolutionIconUrl DefaultValue="([^"]+)"/u,
      `${relativePath} HighResolutionIconUrl`
    ),
    iconUrl: readRequiredMatch(
      manifestText,
      /<IconUrl DefaultValue="([^"]+)"/u,
      `${relativePath} IconUrl`
    ),
    id: readRequiredMatch(
      manifestText,
      /<Id>([^<]+)<\/Id>/u,
      `${relativePath} Id`
    ),
    path: relativePath,
    sourceLocation: readRequiredMatch(
      manifestText,
      /<SourceLocation DefaultValue="([^"]+)"/u,
      `${relativePath} SourceLocation`
    ),
    supportUrl: readRequiredMatch(
      manifestText,
      /<SupportUrl DefaultValue="([^"]+)"/u,
      `${relativePath} SupportUrl`
    ),
    text: manifestText,
    urls,
    version: readRequiredMatch(
      manifestText,
      /<Version>([^<]+)<\/Version>/u,
      `${relativePath} Version`
    ),
  };
}

function readRequiredMatch(
  input: string,
  pattern: RegExp,
  label: string
): string {
  const match = pattern.exec(input);
  if (match?.[1] === undefined) {
    throw new Error(`Could not read ${label}.`);
  }

  return match[1];
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

function normalizeWhitespace(value: string): string {
  return value.replace(/\s+/gu, " ").trim();
}

async function readText(relativePath: string): Promise<string> {
  return readFile(path.join(REPOSITORY_ROOT, relativePath), "utf8");
}

async function readOptionalText(
  relativePath: string
): Promise<string | undefined> {
  try {
    return await readText(relativePath);
  } catch (error: unknown) {
    if (
      typeof error === "object" &&
      error !== null &&
      "code" in error &&
      error.code === "ENOENT"
    ) {
      return undefined;
    }

    throw error;
  }
}

const isDirectExecution =
  process.argv[1]?.endsWith("check-repo-contracts.ts") ?? false;

if (isDirectExecution) {
  void main().catch((error: unknown) => {
    console.error("MarkOut repository contract check crashed.", error);
    process.exitCode = 1;
  });
}
