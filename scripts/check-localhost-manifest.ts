import { readFile } from "fs/promises";
import path from "path";

const LOCAL_RUNTIME = {
  addInId: "8a0c8f5a-b1bb-4838-bb0d-dc0732b5d73c",
  commandsUrl: "https://localhost:3000/commands.html?channel=local",
  launcheventUrl: "https://localhost:3000/launchevent.js?channel=local",
  taskpaneUrl: "https://localhost:3000/taskpane.html?channel=local",
};

async function main(): Promise<void> {
  const [manifestText, packageText] = await Promise.all([
    readText("manifest-localhost.xml"),
    readText("package.json"),
  ]);

  const packageJson = JSON.parse(packageText) as { version: string };
  const expectedVersion = `${packageJson.version}.0`;
  const errors: string[] = [];

  const id = readRequiredMatch(manifestText, /<Id>([^<]+)<\/Id>/u, "Id");
  const version = readRequiredMatch(
    manifestText,
    /<Version>([^<]+)<\/Version>/u,
    "Version"
  );
  const sourceLocation = readRequiredMatch(
    manifestText,
    /<SourceLocation DefaultValue="([^"]+)"/u,
    "SourceLocation"
  );
  const appDomains = Array.from(
    manifestText.matchAll(/<AppDomain>([^<]+)<\/AppDomain>/gu),
    (match) => match[1]
  ).filter((value): value is string => value !== undefined);

  const urls = {
    commands: readRequiredMatch(
      manifestText,
      /<bt:Url id="Commands\.Url" DefaultValue="([^"]+)"/u,
      "Commands.Url"
    ),
    jsRuntime: readRequiredMatch(
      manifestText,
      /<bt:Url id="JSRuntime\.Url" DefaultValue="([^"]+)"/u,
      "JSRuntime.Url"
    ),
    taskpane: readRequiredMatch(
      manifestText,
      /<bt:Url id="Taskpane\.Url" DefaultValue="([^"]+)"/u,
      "Taskpane.Url"
    ),
    webViewRuntime: readRequiredMatch(
      manifestText,
      /<bt:Url id="WebViewRuntime\.Url" DefaultValue="([^"]+)"/u,
      "WebViewRuntime.Url"
    ),
  };

  if (id !== LOCAL_RUNTIME.addInId) {
    errors.push("Local manifest add-in ID does not match the runtime config.");
  }

  if (version !== expectedVersion) {
    errors.push(
      `Local manifest version \`${version}\` does not match package.json version \`${expectedVersion}\`.`
    );
  }

  if (sourceLocation !== LOCAL_RUNTIME.taskpaneUrl) {
    errors.push("Local SourceLocation must match the runtime taskpane URL.");
  }

  if (urls.commands !== LOCAL_RUNTIME.commandsUrl) {
    errors.push("Local Commands.Url must match the runtime commands URL.");
  }

  if (urls.jsRuntime !== LOCAL_RUNTIME.launcheventUrl) {
    errors.push("Local JSRuntime.Url must match the runtime launchevent URL.");
  }

  if (urls.taskpane !== LOCAL_RUNTIME.taskpaneUrl) {
    errors.push("Local Taskpane.Url must match the runtime taskpane URL.");
  }

  if (urls.webViewRuntime !== LOCAL_RUNTIME.commandsUrl) {
    errors.push(
      "Local WebViewRuntime.Url must match the runtime commands URL."
    );
  }

  if (!appDomains.includes("https://localhost:3000")) {
    errors.push("Local AppDomains must include https://localhost:3000.");
  }

  if (errors.length > 0) {
    console.error("MarkOut localhost manifest check failed:");
    for (const error of errors) {
      console.error(`- ${error}`);
    }
    process.exit(1);
  }

  console.log("MarkOut localhost manifest check passed.");
}

function readRequiredMatch(
  input: string,
  pattern: RegExp,
  label: string
): string {
  const match = pattern.exec(input);
  if (match?.[1] === undefined) {
    throw new Error(`Could not read ${label} from manifest-localhost.xml.`);
  }

  return match[1];
}

async function readText(relativePath: string): Promise<string> {
  return readFile(path.join(process.cwd(), relativePath), "utf8");
}

void main().catch((error: unknown) => {
  console.error("MarkOut localhost manifest check crashed.", error);
  process.exit(1);
});
