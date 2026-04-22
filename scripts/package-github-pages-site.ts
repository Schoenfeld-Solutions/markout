import { cp, mkdir, readdir, rm } from "fs/promises";
import path from "path";

export interface PackageGithubPagesSiteOptions {
  betaRoot: string;
  outputRoot: string;
  productionRoot: string;
}

export async function packageGithubPagesSite({
  betaRoot,
  outputRoot,
  productionRoot,
}: PackageGithubPagesSiteOptions): Promise<void> {
  await rm(outputRoot, { force: true, recursive: true });
  await mkdir(outputRoot, { recursive: true });

  await copyRequiredFile(
    path.join(betaRoot, "site", "index.html"),
    path.join(outputRoot, "index.html")
  );
  await copyRequiredFile(
    path.join(betaRoot, "site", "404.html"),
    path.join(outputRoot, "404.html")
  );
  await copyRequiredFile(
    path.join(productionRoot, "manifest.xml"),
    path.join(outputRoot, "manifest.xml")
  );
  await copyRequiredFile(
    path.join(betaRoot, "manifest.beta.xml"),
    path.join(outputRoot, "manifest.beta.xml")
  );
  await copyRequiredDirectory(
    path.join(betaRoot, "assets"),
    path.join(outputRoot, "assets")
  );

  await packageChannelRuntime(productionRoot, path.join(outputRoot, "outlook"));
  await packageChannelRuntime(betaRoot, path.join(outputRoot, "outlook-beta"));
}

async function packageChannelRuntime(
  sourceRoot: string,
  channelRoot: string
): Promise<void> {
  await mkdir(channelRoot, { recursive: true });
  await copyRequiredDirectory(
    path.join(sourceRoot, "assets"),
    path.join(channelRoot, "assets")
  );
  await copyDirectoryContents(path.join(sourceRoot, "dist"), channelRoot);
}

async function copyDirectoryContents(
  sourceDirectory: string,
  targetDirectory: string
): Promise<void> {
  const directoryEntries = await readdir(sourceDirectory, {
    withFileTypes: true,
  });

  for (const directoryEntry of directoryEntries) {
    const sourcePath = path.join(sourceDirectory, directoryEntry.name);
    const targetPath = path.join(targetDirectory, directoryEntry.name);

    await cp(sourcePath, targetPath, {
      force: true,
      recursive: directoryEntry.isDirectory(),
    });
  }
}

async function copyRequiredDirectory(
  sourceDirectory: string,
  targetDirectory: string
): Promise<void> {
  await cp(sourceDirectory, targetDirectory, { force: true, recursive: true });
}

async function copyRequiredFile(
  sourceFile: string,
  targetFile: string
): Promise<void> {
  await cp(sourceFile, targetFile, { force: true });
}

function parseCliArguments(argumentsList: string[]): Record<string, string> {
  const parsedArguments: Record<string, string> = {};

  for (let index = 0; index < argumentsList.length; index += 2) {
    const key = argumentsList[index];
    const value = argumentsList[index + 1];

    if (key === undefined || value === undefined || !key.startsWith("--")) {
      throw new Error(
        "Expected CLI arguments in pairs: --beta-root <dir> --production-root <dir> --output-root <dir>."
      );
    }

    parsedArguments[key.slice(2)] = value;
  }

  return parsedArguments;
}

async function main(): Promise<void> {
  const parsedArguments = parseCliArguments(process.argv.slice(2));
  const betaRoot = parsedArguments["beta-root"];
  const outputRoot = parsedArguments["output-root"];
  const productionRoot = parsedArguments["production-root"];

  if (
    betaRoot === undefined ||
    outputRoot === undefined ||
    productionRoot === undefined
  ) {
    throw new Error(
      "Missing required arguments. Use --beta-root, --production-root, and --output-root."
    );
  }

  await packageGithubPagesSite({
    betaRoot,
    outputRoot,
    productionRoot,
  });
}

const isDirectExecution =
  process.argv[1]?.endsWith("package-github-pages-site.ts") ?? false;

if (isDirectExecution) {
  void main().catch((error: unknown) => {
    console.error("MarkOut GitHub Pages packaging failed.", error);
    process.exitCode = 1;
  });
}
