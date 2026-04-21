import { readFile, writeFile } from "fs/promises";

const PAGES_BASE_URL = "https://schoenfeld-solutions.github.io/markout";

export interface PreviewManifestOptions {
  displayName: string;
  previewBaseUrl: string;
}

export function createPreviewManifest(
  sourceManifest: string,
  { displayName, previewBaseUrl }: PreviewManifestOptions
): string {
  const previewUrl = new URL(previewBaseUrl);
  const normalizedBaseUrl = `${previewUrl.origin}${previewUrl.pathname.replace(/\/$/, "")}`;
  const escapedDisplayName = escapeXmlAttribute(displayName);

  if (!sourceManifest.includes(PAGES_BASE_URL)) {
    throw new Error(
      `Preview manifest generation expected the source manifest to contain ${PAGES_BASE_URL}.`
    );
  }

  let manifest = sourceManifest.replaceAll(PAGES_BASE_URL, normalizedBaseUrl);
  manifest = replaceDisplayName(manifest, escapedDisplayName);
  manifest = replaceAppDomains(manifest, previewUrl.origin);

  return manifest;
}

function escapeXmlAttribute(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("'", "&apos;");
}

function replaceDisplayName(manifest: string, displayName: string): string {
  const nextManifest = manifest.replace(
    /<DisplayName DefaultValue="[^"]*" \/>/,
    `<DisplayName DefaultValue="${displayName}" />`
  );

  if (nextManifest === manifest) {
    throw new Error(
      "Preview manifest generation could not update DisplayName."
    );
  }

  return nextManifest;
}

function replaceAppDomains(manifest: string, appDomainOrigin: string): string {
  const nextManifest = manifest.replace(
    /<AppDomains>[\s\S]*?<\/AppDomains>/,
    `  <AppDomains>\n    <AppDomain>${appDomainOrigin}</AppDomain>\n  </AppDomains>`
  );

  if (nextManifest === manifest) {
    throw new Error("Preview manifest generation could not update AppDomains.");
  }

  return nextManifest;
}

function parseCliArguments(argumentsList: string[]): Record<string, string> {
  const parsedArguments: Record<string, string> = {};

  for (let index = 0; index < argumentsList.length; index += 2) {
    const key = argumentsList[index];
    const value = argumentsList[index + 1];

    if (key === undefined || value === undefined || !key.startsWith("--")) {
      throw new Error(
        "Expected CLI arguments in pairs: --source <file> --output <file> --base-url <url> --display-name <label>."
      );
    }

    parsedArguments[key.slice(2)] = value;
  }

  return parsedArguments;
}

async function main(): Promise<void> {
  const parsedArguments = parseCliArguments(process.argv.slice(2));
  const sourcePath = parsedArguments.source;
  const outputPath = parsedArguments.output;
  const previewBaseUrl = parsedArguments["base-url"];
  const displayName = parsedArguments["display-name"];

  if (
    sourcePath === undefined ||
    outputPath === undefined ||
    previewBaseUrl === undefined ||
    displayName === undefined
  ) {
    throw new Error(
      "Missing required arguments. Use --source, --output, --base-url, and --display-name."
    );
  }

  const sourceManifest = await readFile(sourcePath, "utf8");
  const previewManifest = createPreviewManifest(sourceManifest, {
    displayName,
    previewBaseUrl,
  });

  await writeFile(outputPath, previewManifest, "utf8");
}

const isDirectExecution =
  process.argv[1]?.endsWith("generate-preview-manifest.ts") ?? false;

if (isDirectExecution) {
  void main().catch((error: unknown) => {
    console.error("MarkOut preview manifest generation failed.", error);
    process.exitCode = 1;
  });
}
