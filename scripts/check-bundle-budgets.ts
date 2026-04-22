import { readdir, stat } from "fs/promises";
import path from "path";

export const DEFAULT_MAX_JAVASCRIPT_ASSET_BYTES = 244 * 1024;

export interface JavaScriptAsset {
  filePath: string;
  sizeBytes: number;
}

export function findOversizedJavaScriptAssets(
  assets: JavaScriptAsset[],
  maxBytes: number = DEFAULT_MAX_JAVASCRIPT_ASSET_BYTES
): JavaScriptAsset[] {
  return assets.filter((asset) => asset.sizeBytes > maxBytes);
}

async function readJavaScriptAssets(
  directoryPath: string
): Promise<JavaScriptAsset[]> {
  const directoryEntries = await readdir(directoryPath, {
    withFileTypes: true,
  });
  const assets: JavaScriptAsset[] = [];

  for (const directoryEntry of directoryEntries) {
    const resolvedPath = path.join(directoryPath, directoryEntry.name);

    if (directoryEntry.isDirectory()) {
      assets.push(...(await readJavaScriptAssets(resolvedPath)));
      continue;
    }

    if (!directoryEntry.isFile() || !resolvedPath.endsWith(".js")) {
      continue;
    }

    const fileStats = await stat(resolvedPath);
    assets.push({
      filePath: resolvedPath,
      sizeBytes: fileStats.size,
    });
  }

  return assets;
}

async function main(): Promise<void> {
  const distDirectory = process.argv[2] ?? path.join(process.cwd(), "dist");
  const assets = await readJavaScriptAssets(distDirectory);
  const oversizedAssets = findOversizedJavaScriptAssets(assets);

  if (oversizedAssets.length === 0) {
    console.log(
      `MarkOut bundle budgets passed. All JavaScript assets are at or below ${DEFAULT_MAX_JAVASCRIPT_ASSET_BYTES} bytes.`
    );
    return;
  }

  for (const oversizedAsset of oversizedAssets) {
    console.error(
      `${path.relative(process.cwd(), oversizedAsset.filePath)} is ${oversizedAsset.sizeBytes} bytes, which exceeds the ${DEFAULT_MAX_JAVASCRIPT_ASSET_BYTES}-byte budget.`
    );
  }

  process.exitCode = 1;
}

const isDirectExecution =
  process.argv[1]?.endsWith("check-bundle-budgets.ts") ?? false;

if (isDirectExecution) {
  void main().catch((error: unknown) => {
    console.error("MarkOut bundle budget verification failed.", error);
    process.exitCode = 1;
  });
}
