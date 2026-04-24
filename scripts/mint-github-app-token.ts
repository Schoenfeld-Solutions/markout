import { createSign } from "crypto";
import { appendFile } from "fs/promises";

type Environment = Readonly<Record<string, string | undefined>>;

interface GitHubRequestOptions {
  bearerToken: string;
  body?: unknown;
  method: "GET" | "POST";
  path: string;
}

const GITHUB_API_BASE_URL = "https://api.github.com";

async function mintInstallationToken(
  environment: Environment = process.env
): Promise<string> {
  const appId = readRequiredEnvironmentValue(
    environment,
    "MARKOUT_RELEASE_BOT_APP_ID"
  );
  const privateKey = readRequiredEnvironmentValue(
    environment,
    "MARKOUT_RELEASE_BOT_PRIVATE_KEY"
  ).replace(/\\n/gu, "\n");
  const repository = readRepository(environment);
  const jwt = createGitHubAppJwt(appId, privateKey);
  const installation = await githubRequest({
    bearerToken: jwt,
    method: "GET",
    path: `/repos/${encodeURIComponent(repository.owner)}/${encodeURIComponent(
      repository.repo
    )}/installation`,
  });
  const installationId = readNumberProperty(
    installation,
    "id",
    "GitHub App installation id"
  );
  const installationTokenResponse = await githubRequest({
    bearerToken: jwt,
    body: {
      permissions: {
        contents: "write",
      },
    },
    method: "POST",
    path: `/app/installations/${installationId}/access_tokens`,
  });

  return readStringProperty(
    installationTokenResponse,
    "token",
    "GitHub App installation token"
  );
}

function createGitHubAppJwt(appId: string, privateKey: string): string {
  const issuedAtSeconds = Math.floor(Date.now() / 1000) - 60;
  const expiresAtSeconds = issuedAtSeconds + 9 * 60;
  const encodedHeader = base64UrlJson({
    alg: "RS256",
    typ: "JWT",
  });
  const encodedPayload = base64UrlJson({
    exp: expiresAtSeconds,
    iat: issuedAtSeconds,
    iss: appId,
  });
  const unsignedToken = `${encodedHeader}.${encodedPayload}`;
  const signer = createSign("RSA-SHA256");
  signer.update(unsignedToken);
  signer.end();
  return `${unsignedToken}.${signer.sign(privateKey, "base64url")}`;
}

function base64UrlJson(value: unknown): string {
  return Buffer.from(JSON.stringify(value)).toString("base64url");
}

async function githubRequest({
  bearerToken,
  body,
  method,
  path,
}: GitHubRequestOptions): Promise<unknown> {
  const requestInit: RequestInit = {
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
      "User-Agent": "markout-release-bot-token-mint",
      "X-GitHub-Api-Version": "2022-11-28",
    },
    method,
  };

  if (body !== undefined) {
    requestInit.body = JSON.stringify(body);
  }

  const response = await fetch(`${GITHUB_API_BASE_URL}${path}`, requestInit);
  const responseText = await response.text();

  if (!response.ok) {
    throw new Error(
      `GitHub API ${method} ${path} failed with ${response.status}: ${responseText}`
    );
  }

  return JSON.parse(responseText) as unknown;
}

function readRequiredEnvironmentValue(
  environment: Environment,
  variableName: string
): string {
  const value = environment[variableName]?.trim() ?? "";
  if (value.length === 0) {
    throw new Error(`${variableName} is required.`);
  }

  return value;
}

function readRepository(environment: Environment): {
  owner: string;
  repo: string;
} {
  const repository = readRequiredEnvironmentValue(
    environment,
    "GITHUB_REPOSITORY"
  );
  const [owner, repo] = repository.split("/");

  if (!isNonEmptyString(owner) || !isNonEmptyString(repo)) {
    throw new Error("GITHUB_REPOSITORY must be in owner/repo format.");
  }

  return {
    owner,
    repo,
  };
}

function readNumberProperty(
  value: unknown,
  propertyName: string,
  label: string
): number {
  if (!isRecord(value) || typeof value[propertyName] !== "number") {
    throw new Error(`${label} was missing from the GitHub API response.`);
  }

  return value[propertyName];
}

function readStringProperty(
  value: unknown,
  propertyName: string,
  label: string
): string {
  if (!isRecord(value) || !isNonEmptyString(value[propertyName])) {
    throw new Error(`${label} was missing from the GitHub API response.`);
  }

  return value[propertyName];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

async function main(): Promise<void> {
  const token = await mintInstallationToken();
  console.log(`::add-mask::${token}`);

  const githubOutputPath = process.env.GITHUB_OUTPUT;
  if (!isNonEmptyString(githubOutputPath)) {
    throw new Error("GITHUB_OUTPUT is required so the token is not printed.");
  }

  await appendFile(githubOutputPath, `token=${token}\n`, "utf8");
  console.log("Minted a short-lived release bot installation token.");
}

const isDirectExecution =
  process.argv[1]?.endsWith("mint-github-app-token.ts") ?? false;

if (isDirectExecution) {
  void main().catch((error: unknown) => {
    console.error("Could not mint MarkOut release bot token.", error);
    process.exitCode = 1;
  });
}
