import { readFile } from "fs/promises";
import path from "path";

type ChannelId = "beta" | "local" | "production";

interface RuntimeChannelConfigSnapshot {
  addInId: string;
  appBaseUrl: string;
  channelId: ChannelId;
  commandsUrl: string;
  launcheventUrl: string;
  storageNamespace: string;
  supportUrl: string;
  taskpaneUrl: string;
}

interface ManifestSnapshot {
  appDomains: string[];
  displayName: string;
  id: string;
  path: string;
  sourceLocation: string;
  urls: Record<string, string>;
  version: string;
}

interface ExpectedManifestContract {
  displayName: string;
  manifestPath: string;
}

const REPOSITORY_ROOT = process.cwd();
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
    file: "CONTRIBUTING.md",
    snippet:
      "`manifest.beta.xml` and `/outlook-beta/` are the post-merge preview/testing channel sourced from `main`.",
  },
  {
    file: "CONTRIBUTING.md",
    snippet:
      "`manifest.xml` and `/outlook/` are the stable production channel sourced from `release/production`.",
  },
];

const EXPECTED_WORKFLOW_SNIPPETS = [
  {
    file: ".github/workflows/release.yaml",
    snippet:
      "Missing origin/release/production. Refusing to publish GitHub Pages without an explicit production source branch.",
  },
  {
    file: ".github/workflows/release.yaml",
    snippet: "MARKOUT_HOST_SMOKE_STORAGE_STATE_JSON is required.",
  },
  {
    file: ".github/workflows/release.yaml",
    snippet: "MARKOUT_HOST_SMOKE_RECIPIENT is required.",
  },
  {
    file: ".github/workflows/promote-production.yaml",
    snippet: "production-promotion",
  },
  {
    file: ".github/workflows/promote-production.yaml",
    snippet: "Build and Publish GitHub Pages",
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
  const [packageJson, readme, contributing] = await Promise.all([
    readJson<{ version: string }>("package.json"),
    readText("README.md"),
    readText("CONTRIBUTING.md"),
  ]);
  const { getAllRuntimeChannelConfigs } = (await import(
    new URL("../src/lib/runtime.ts", import.meta.url).href
  )) as {
    getAllRuntimeChannelConfigs: () => RuntimeChannelConfigSnapshot[];
  };

  const contractErrors: string[] = [];
  const expectedManifestVersion = `${packageJson.version}.0`;
  const manifests = await Promise.all(
    getAllRuntimeChannelConfigs().map(async (runtimeChannelConfig) => {
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

      if (snapshot.urls[urlId] !== undefined) {
        const parsedUrl = new URL(snapshot.urls[urlId]);
        const channelQuery = parsedUrl.searchParams.get("channel");

        if (channelQuery !== runtimeChannelConfig.channelId) {
          contractErrors.push(
            `${snapshot.path} ${urlId} must carry \`?channel=${runtimeChannelConfig.channelId}\`.`
          );
        }
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

  checkSnippetPresence(readme, contributing, contractErrors);
  await checkWorkflowSnippets(contractErrors);

  if (contractErrors.length > 0) {
    console.error("MarkOut repository contract check failed:");
    for (const error of contractErrors) {
      console.error(`- ${error}`);
    }
    process.exit(1);
  }

  console.log("MarkOut repository contract check passed.");
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

void main().catch((error: unknown) => {
  console.error("MarkOut repository contract check crashed.", error);
  process.exit(1);
});
