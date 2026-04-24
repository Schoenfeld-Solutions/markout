import { existsSync } from "node:fs";
import { readFile, stat } from "node:fs/promises";
import http, {
  type IncomingMessage,
  type Server as HttpServer,
  type ServerResponse,
} from "node:http";
import https, { type ServerOptions as HttpsServerOptions } from "node:https";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import webpack, {
  type Configuration,
  type Stats,
  type Watching,
} from "webpack";

type Protocol = "http" | "https";

interface LocalDevServerOptions {
  host: string;
  port: number;
  protocol: Protocol;
  taskpaneMock: boolean;
}

interface BuildState {
  isReady: boolean;
  lastErrorMessage?: string;
}

type WebpackConfigFactory = (
  env: Record<string, string | boolean>,
  options: { mode: "development" | "production" }
) => Configuration | Promise<Configuration>;

const repositoryRoot = process.cwd();
const outputRoot = path.join(repositoryRoot, "dist");
const requireFromScript = createRequire(import.meta.url);
const createWebpackConfig = webpackConfigFactory();

void main().catch((error: unknown) => {
  console.error("MarkOut local dev server failed.", error);
  process.exitCode = 1;
});

async function main(): Promise<void> {
  const options = parseOptions(process.argv.slice(2));
  const buildState: BuildState = { isReady: false };
  const server = await createServer(options, buildState);
  const configuration = await createWebpackConfig(
    {
      taskpaneMock: options.taskpaneMock ? "true" : "false",
    },
    { mode: "development" }
  );
  const compiler = webpack(configuration);
  const watching = compiler.watch({}, (error, stats) => {
    updateBuildState(buildState, error, stats);
  });

  if (watching === undefined) {
    throw new Error("Webpack did not start a watch compiler.");
  }

  await listen(server, options);
  console.log(
    `MarkOut local dev server listening at ${options.protocol}://${options.host}:${options.port}/`
  );

  let isShuttingDown = false;
  const shutdown = async (): Promise<void> => {
    if (isShuttingDown) {
      return;
    }

    isShuttingDown = true;
    await Promise.all([closeServer(server), closeWatching(watching)]);
  };

  process.once("SIGINT", () => {
    void shutdown().finally(() => process.exit(0));
  });
  process.once("SIGTERM", () => {
    void shutdown().finally(() => process.exit(0));
  });
}

function webpackConfigFactory(): WebpackConfigFactory {
  return webpackConfigRequire("../webpack.config.js");
}

function webpackConfigRequire(modulePath: string): WebpackConfigFactory {
  return requireFromScript(modulePath) as WebpackConfigFactory;
}

function parseOptions(args: string[]): LocalDevServerOptions {
  let host = process.env.MARKOUT_DEV_SERVER_HOST ?? "localhost";
  let port =
    parsePositiveInteger(process.env.MARKOUT_DEV_SERVER_PORT) ??
    parsePositiveInteger(process.env.npm_package_config_dev_server_port) ??
    3000;
  let protocol: Protocol = "https";
  let taskpaneMock = false;

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];

    if (argument === undefined) {
      continue;
    }

    if (argument === "--http") {
      protocol = "http";
      continue;
    }

    if (argument === "--https") {
      protocol = "https";
      continue;
    }

    if (argument === "--taskpane-mock") {
      taskpaneMock = true;
      continue;
    }

    if (argument === "--host") {
      const value = args[index + 1];

      if (value === undefined || value.trim() === "") {
        throw new Error("--host requires a non-empty value.");
      }

      host = value;
      index += 1;
      continue;
    }

    if (argument.startsWith("--host=")) {
      const value = argument.slice("--host=".length);

      if (value.trim() === "") {
        throw new Error("--host requires a non-empty value.");
      }

      host = value;
      continue;
    }

    if (argument === "--port") {
      const value = args[index + 1];

      if (value === undefined) {
        throw new Error("--port requires a value.");
      }

      port = parseRequiredPort(value);
      index += 1;
      continue;
    }

    if (argument.startsWith("--port=")) {
      port = parseRequiredPort(argument.slice("--port=".length));
      continue;
    }

    throw new Error(`Unknown local dev server argument: ${argument}`);
  }

  return { host, port, protocol, taskpaneMock };
}

function parsePositiveInteger(value: string | undefined): number | undefined {
  if (value === undefined || value.trim() === "") {
    return undefined;
  }

  const parsedValue = Number.parseInt(value, 10);

  if (!Number.isFinite(parsedValue) || parsedValue <= 0) {
    return undefined;
  }

  return parsedValue;
}

function parseRequiredPort(value: string): number {
  const parsedValue = parsePositiveInteger(value);

  if (parsedValue === undefined || parsedValue > 65_535) {
    throw new Error(`Invalid local dev server port: ${value}`);
  }

  return parsedValue;
}

async function createServer(
  options: LocalDevServerOptions,
  buildState: BuildState
): Promise<HttpServer> {
  const requestHandler = (
    request: IncomingMessage,
    response: ServerResponse
  ): void => {
    void serveRequest(request, response, buildState).catch((error: unknown) => {
      console.error("Failed to serve local dev asset.", error);
      sendText(request, response, 500, "Internal dev server error.");
    });
  };

  if (options.protocol === "https") {
    return https.createServer(await readHttpsServerOptions(), requestHandler);
  }

  return http.createServer(requestHandler);
}

async function readHttpsServerOptions(): Promise<HttpsServerOptions> {
  const defaultCertificateDirectory = path.join(
    os.homedir(),
    ".office-addin-dev-certs"
  );
  const certificatePath =
    process.env.MARKOUT_DEV_TLS_CERT_PATH ??
    path.join(defaultCertificateDirectory, "localhost.crt");
  const keyPath =
    process.env.MARKOUT_DEV_TLS_KEY_PATH ??
    path.join(defaultCertificateDirectory, "localhost.key");

  if (!existsSync(certificatePath) || !existsSync(keyPath)) {
    throw new Error(
      "Local HTTPS requires MARKOUT_DEV_TLS_CERT_PATH and MARKOUT_DEV_TLS_KEY_PATH, or existing ~/.office-addin-dev-certs/localhost.crt and localhost.key files."
    );
  }

  return {
    cert: await readFile(certificatePath),
    key: await readFile(keyPath),
  };
}

async function serveRequest(
  request: IncomingMessage,
  response: ServerResponse,
  buildState: BuildState
): Promise<void> {
  if (request.method !== "GET" && request.method !== "HEAD") {
    sendText(request, response, 405, "Method not allowed.");
    return;
  }

  if (!buildState.isReady) {
    sendText(
      request,
      response,
      503,
      buildState.lastErrorMessage ?? "Webpack build is not ready yet.",
      {
        "Retry-After": "1",
        "X-MarkOut-Build-State": "pending",
      }
    );
    return;
  }

  const filePath = resolveRequestedFilePath(request.url ?? "/");

  if (filePath === undefined) {
    sendText(request, response, 400, "Invalid asset path.");
    return;
  }

  const fileStats = await stat(filePath).catch(() => undefined);

  if (fileStats?.isFile() !== true) {
    sendText(request, response, 404, "Not found.");
    return;
  }

  const headers = {
    ...defaultResponseHeaders(),
    "Content-Length": String(fileStats.size),
    "Content-Type": contentTypeForPath(filePath),
  };

  response.writeHead(200, headers);

  if (request.method === "HEAD") {
    response.end();
    return;
  }

  response.end(await readFile(filePath));
}

function resolveRequestedFilePath(requestUrl: string): string | undefined {
  let pathname: string;

  try {
    pathname = new URL(requestUrl, "http://localhost").pathname;
  } catch {
    return undefined;
  }

  const requestedPath = pathname === "/" ? "/taskpane.html" : pathname;
  let decodedPath: string;

  try {
    decodedPath = decodeURIComponent(requestedPath);
  } catch {
    return undefined;
  }

  if (decodedPath.includes("\0")) {
    return undefined;
  }

  const relativePath = decodedPath.replace(/^\/+/, "");
  const resolvedPath = path.resolve(outputRoot, relativePath);
  const relativeToOutput = path.relative(outputRoot, resolvedPath);
  const isInsideOutput =
    relativeToOutput === "" ||
    (!relativeToOutput.startsWith("..") && !path.isAbsolute(relativeToOutput));

  return isInsideOutput ? resolvedPath : undefined;
}

function sendText(
  request: IncomingMessage,
  response: ServerResponse,
  statusCode: number,
  body: string,
  headers: Record<string, string> = {}
): void {
  response.writeHead(statusCode, {
    ...defaultResponseHeaders(),
    ...headers,
    "Content-Length": String(Buffer.byteLength(body)),
    "Content-Type": "text/plain; charset=utf-8",
  });

  if (request.method === "HEAD") {
    response.end();
    return;
  }

  response.end(body);
}

function defaultResponseHeaders(): Record<string, string> {
  return {
    "Access-Control-Allow-Origin": "*",
    "Cache-Control": "no-store",
  };
}

function contentTypeForPath(filePath: string): string {
  const extension = path.extname(filePath).toLowerCase();
  const contentTypes: Record<string, string> = {
    ".css": "text/css; charset=utf-8",
    ".gif": "image/gif",
    ".html": "text/html; charset=utf-8",
    ".jpeg": "image/jpeg",
    ".jpg": "image/jpeg",
    ".js": "text/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".map": "application/json; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml",
  };

  return contentTypes[extension] ?? "application/octet-stream";
}

function updateBuildState(
  buildState: BuildState,
  error: Error | null | undefined,
  stats: Stats | undefined
): void {
  if (error !== null && error !== undefined) {
    buildState.isReady = false;
    buildState.lastErrorMessage = error.message;
    console.error("MarkOut webpack compilation failed.", error);
    return;
  }

  if (stats === undefined) {
    buildState.isReady = false;
    buildState.lastErrorMessage = "Webpack did not return compilation stats.";
    console.error(buildState.lastErrorMessage);
    return;
  }

  printStats(stats);

  if (stats.hasErrors()) {
    buildState.isReady = false;
    buildState.lastErrorMessage = "Webpack compilation failed.";
    return;
  }

  buildState.isReady = true;
  delete buildState.lastErrorMessage;
  console.log("MarkOut local build ready.");
}

function printStats(stats: Stats): void {
  const output = stats.toString({
    all: false,
    assets: true,
    colors: process.stdout.isTTY,
    errors: true,
    timings: true,
    warnings: true,
  });

  if (output.trim() !== "") {
    console.log(output);
  }
}

async function listen(
  server: HttpServer,
  options: LocalDevServerOptions
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const rejectListen = (error: Error): void => {
      reject(error);
    };

    server.once("error", rejectListen);
    server.listen(options.port, options.host, () => {
      server.off("error", rejectListen);
      resolve();
    });
  });
}

async function closeServer(server: HttpServer): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    server.close((error?: Error) => {
      if (error !== undefined) {
        reject(error);
        return;
      }

      resolve();
    });
  });
}

async function closeWatching(watching: Watching): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    watching.close((error: Error | null) => {
      if (error !== null) {
        reject(error);
        return;
      }

      resolve();
    });
  });
}
