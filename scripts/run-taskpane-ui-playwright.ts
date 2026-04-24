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
  const port = await resolveTaskpaneUiPort();
  const baseUrl =
    process.env.MARKOUT_TASKPANE_UI_URL ??
    `http://localhost:${port}/taskpane-mock.html`;
  const timeoutMs = Number.parseInt(
    process.env.MARKOUT_TASKPANE_UI_TIMEOUT_MS ?? "",
    10
  );
  const serverProcess = startLocalDevServer(port);

  try {
    console.log(`MarkOut taskpane UI harness starting at ${baseUrl}`);
    await waitForUrl(
      baseUrl,
      Number.isFinite(timeoutMs) ? timeoutMs : DEFAULT_TIMEOUT_MS
    );
    await runTaskpaneUiPlaywright({ baseUrl });
    console.log("MarkOut taskpane UI harness passed.");
  } finally {
    await stopProcess(serverProcess);
  }
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
