export type ChannelId = "beta" | "local" | "production";

export interface RuntimeChannelConfig {
  addInId: string;
  appBaseUrl: string;
  channelId: ChannelId;
  commandsUrl: string;
  launcheventUrl: string;
  storageNamespace: string;
  supportUrl: string;
  taskpaneUrl: string;
}

export type MarkOutErrorCode =
  | "office-compose-item-missing"
  | "office-selection-api-unavailable"
  | "restore-state-too-large"
  | "unsupported-body-type";

interface MarkOutErrorOptions {
  cause?: unknown;
}

const SUPPORT_URL = "https://github.com/Schoenfeld-Solutions/markout";

function withChannelQuery(url: string, channelId: ChannelId): string {
  const parsedUrl = new URL(url);
  parsedUrl.searchParams.set("channel", channelId);
  return parsedUrl.toString();
}

const RUNTIME_CHANNELS: Record<ChannelId, RuntimeChannelConfig> = {
  beta: {
    addInId: "934e43d2-950c-4cab-8ec0-4ff2808e6c11",
    appBaseUrl: "https://schoenfeld-solutions.github.io/markout/outlook-beta",
    channelId: "beta",
    commandsUrl: withChannelQuery(
      "https://schoenfeld-solutions.github.io/markout/outlook-beta/commands.html",
      "beta"
    ),
    launcheventUrl: withChannelQuery(
      "https://schoenfeld-solutions.github.io/markout/outlook-beta/launchevent.js",
      "beta"
    ),
    storageNamespace: "markout.beta",
    supportUrl: SUPPORT_URL,
    taskpaneUrl: withChannelQuery(
      "https://schoenfeld-solutions.github.io/markout/outlook-beta/taskpane.html",
      "beta"
    ),
  },
  local: {
    addInId: "8a0c8f5a-b1bb-4838-bb0d-dc0732b5d73c",
    appBaseUrl: "https://localhost:3000",
    channelId: "local",
    commandsUrl: withChannelQuery(
      "https://localhost:3000/commands.html",
      "local"
    ),
    launcheventUrl: withChannelQuery(
      "https://localhost:3000/launchevent.js",
      "local"
    ),
    storageNamespace: "markout.local",
    supportUrl: SUPPORT_URL,
    taskpaneUrl: withChannelQuery(
      "https://localhost:3000/taskpane.html",
      "local"
    ),
  },
  production: {
    addInId: "05c2e1c9-3e1d-406e-9a91-e9ac64854143",
    appBaseUrl: "https://schoenfeld-solutions.github.io/markout/outlook",
    channelId: "production",
    commandsUrl: withChannelQuery(
      "https://schoenfeld-solutions.github.io/markout/outlook/commands.html",
      "production"
    ),
    launcheventUrl: withChannelQuery(
      "https://schoenfeld-solutions.github.io/markout/outlook/launchevent.js",
      "production"
    ),
    storageNamespace: "markout.production",
    supportUrl: SUPPORT_URL,
    taskpaneUrl: withChannelQuery(
      "https://schoenfeld-solutions.github.io/markout/outlook/taskpane.html",
      "production"
    ),
  },
};

export class MarkOutError extends Error {
  public readonly code: MarkOutErrorCode;

  public constructor(
    code: MarkOutErrorCode,
    message: string,
    options: MarkOutErrorOptions = {}
  ) {
    super(message);
    this.code = code;
    this.name = "MarkOutError";
    if ("cause" in Error.prototype || options.cause !== undefined) {
      Object.defineProperty(this, "cause", {
        configurable: true,
        enumerable: false,
        value: options.cause,
        writable: true,
      });
    }
  }
}

export function getRuntimeChannelConfig(
  channelId: ChannelId
): RuntimeChannelConfig {
  return RUNTIME_CHANNELS[channelId];
}

export function getAllRuntimeChannelConfigs(): RuntimeChannelConfig[] {
  return Object.values(RUNTIME_CHANNELS);
}

export function resolveRuntimeChannelConfig(
  currentUrl: string | undefined = readCurrentLocationHref()
): RuntimeChannelConfig {
  if (currentUrl === undefined || currentUrl.trim().length === 0) {
    return RUNTIME_CHANNELS.production;
  }

  let parsedUrl: URL;
  try {
    parsedUrl = new URL(currentUrl);
  } catch {
    return RUNTIME_CHANNELS.production;
  }

  const explicitChannelId = parsedUrl.searchParams.get("channel");

  if (
    explicitChannelId === "beta" ||
    explicitChannelId === "local" ||
    explicitChannelId === "production"
  ) {
    return RUNTIME_CHANNELS[explicitChannelId];
  }

  if (parsedUrl.hostname === "localhost") {
    return RUNTIME_CHANNELS.local;
  }

  if (parsedUrl.pathname.includes("/outlook-beta/")) {
    return RUNTIME_CHANNELS.beta;
  }

  return RUNTIME_CHANNELS.production;
}

export function getChannelScopedKey(
  runtimeChannelConfig: Pick<RuntimeChannelConfig, "storageNamespace">,
  keySuffix: string
): string {
  return `${runtimeChannelConfig.storageNamespace}.${keySuffix}`;
}

export function isMarkOutErrorCode(
  error: unknown,
  code: MarkOutErrorCode
): error is MarkOutError {
  return error instanceof MarkOutError && error.code === code;
}

function readCurrentLocationHref(): string | undefined {
  if (
    typeof globalThis.location !== "object" ||
    typeof globalThis.location.href !== "string"
  ) {
    return undefined;
  }

  return globalThis.location.href;
}
