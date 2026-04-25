import http from "node:http";
import https from "node:https";
import net from "node:net";
import path from "node:path";
import { spawn, type ChildProcess } from "node:child_process";
import { runTaskpaneUiPlaywright } from "../spec/taskpane-ui-playwright";

const DEFAULT_TIMEOUT_MS = 45_000;

void run().catch((error: unknown) => {
  console.error("MarkOut taskpane UI harness failed.", error);
  process.exitCode = 1;
});

async function run(): Promise<void> {
  const commandLineOptions = parseCommandLineOptions(process.argv.slice(2));
  const port = await resolveTaskpaneUiPort();
  const defaultOrigin = `http://localhost:${port}`;
  const baseUrl =
    process.env.MARKOUT_TASKPANE_UI_URL ??
    `${defaultOrigin}/taskpane-mock.html`;
  const owaHostUrl =
    process.env.MARKOUT_TASKPANE_UI_OWA_HOST_URL ??
    `${defaultOrigin}/owa-taskpane-host.html`;
  const timeoutMs = Number.parseInt(
    process.env.MARKOUT_TASKPANE_UI_TIMEOUT_MS ?? "",
    10
  );
  const headless = resolveHeadlessMode(commandLineOptions.headed);
  const serverProcess = startLocalDevServer(port);

  try {
    console.log(`MarkOut taskpane UI harness starting at ${baseUrl}`);
    console.log(`MarkOut OWA-like taskpane host starting at ${owaHostUrl}`);
    await waitForUrl(
      baseUrl,
      Number.isFinite(timeoutMs) ? timeoutMs : DEFAULT_TIMEOUT_MS
    );
    await waitForUrl(
      owaHostUrl,
      Number.isFinite(timeoutMs) ? timeoutMs : DEFAULT_TIMEOUT_MS
    );
    await runTaskpaneUiPlaywright({ baseUrl, headless, owaHostUrl });
    console.log("MarkOut taskpane UI harness passed.");
  } finally {
    await stopProcess(serverProcess);
  }
}

interface CommandLineOptions {
  headed: boolean;
}

function parseCommandLineOptions(args: string[]): CommandLineOptions {
  let headed = false;

  for (const argument of args) {
    if (argument === "--headed") {
      headed = true;
      continue;
    }

    throw new Error(`Unknown taskpane UI harness argument: ${argument}`);
  }

  return { headed };
}

function resolveHeadlessMode(isHeadedRequested: boolean): boolean {
  if (isHeadedRequested) {
    return false;
  }

  const configuredValue = process.env.MARKOUT_TASKPANE_UI_HEADLESS;

  if (configuredValue === undefined || configuredValue.trim() === "") {
    return true;
  }

  return !["0", "false", "no"].includes(configuredValue.toLowerCase());
}

function startLocalDevServer(port: number): ChildProcess {
  const tsxBin = path.join(
    process.cwd(),
    "node_modules",
    ".bin",
    process.platform === "win32" ? "tsx.cmd" : "tsx"
  );

  return spawn(
    tsxBin,
    [
      "scripts/run-local-dev-server.ts",
      "--http",
      "--taskpane-mock",
      "--port",
      String(port),
    ],
    {
      cwd: process.cwd(),
      stdio: "inherit",
    }
  );
}

async function stopProcess(childProcess: ChildProcess): Promise<void> {
  if (childProcess.exitCode !== null || childProcess.killed) {
    return;
  }

  childProcess.kill("SIGTERM");

  try {
    await waitForProcessExit(childProcess, 5_000);
    return;
  } catch {
    childProcess.kill("SIGKILL");
    await waitForProcessExit(childProcess, 5_000).catch(() => undefined);
  }
}

async function waitForUrl(url: string, timeoutMs: number): Promise<void> {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    const isReady = await pingUrl(url).catch(() => false);

    if (isReady) {
      return;
    }

    await new Promise((resolve) => {
      setTimeout(resolve, 500);
    });
  }

  throw new Error(`Timed out waiting for ${url}.`);
}

async function pingUrl(url: string): Promise<boolean> {
  return new Promise<boolean>((resolve, reject) => {
    const parsedUrl = new URL(url);
    const request =
      parsedUrl.protocol === "https:"
        ? https.get(url, { rejectUnauthorized: false }, (response) => {
            response.resume();
            resolve((response.statusCode ?? 500) < 400);
          })
        : http.get(url, (response) => {
            response.resume();
            resolve((response.statusCode ?? 500) < 400);
          });

    request.on("error", reject);
  });
}

async function resolveTaskpaneUiPort(): Promise<number> {
  const configuredPort = Number.parseInt(
    process.env.MARKOUT_TASKPANE_UI_PORT ?? "",
    10
  );

  if (Number.isFinite(configuredPort) && configuredPort > 0) {
    return configuredPort;
  }

  return new Promise<number>((resolve, reject) => {
    const server = net.createServer();

    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();

      if (address === null || typeof address === "string") {
        server.close(() => {
          reject(
            new Error(
              "Failed to resolve a local port for the taskpane harness."
            )
          );
        });
        return;
      }

      server.close((error) => {
        if (error !== undefined) {
          reject(error);
          return;
        }

        resolve(address.port);
      });
    });
  });
}

async function waitForProcessExit(
  childProcess: ChildProcess,
  timeoutMs: number
): Promise<void> {
  if (childProcess.exitCode !== null) {
    return;
  }

  await new Promise<void>((resolve, reject) => {
    const timeoutId = setTimeout(() => {
      reject(
        new Error("Timed out waiting for the taskpane harness process to exit.")
      );
    }, timeoutMs);

    childProcess.once("exit", () => {
      clearTimeout(timeoutId);
      resolve();
    });
  });
}
