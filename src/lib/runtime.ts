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

export type DiagnosticEventArea =
  | "body-io"
  | "notification"
  | "preview"
  | "render"
  | "restore"
  | "selection";

export type DiagnosticEventLevel = "debug" | "error" | "info" | "warning";

export type DiagnosticEventMetadataValue =
  | boolean
  | null
  | number
  | string
  | undefined;

export type DiagnosticEventMetadata = Record<
  string,
  DiagnosticEventMetadataValue
>;

export interface DiagnosticEventInput {
  area: DiagnosticEventArea;
  code: string;
  level: DiagnosticEventLevel;
  message?: string;
  metadata?: DiagnosticEventMetadata;
}

export interface DiagnosticEventRecord {
  area: DiagnosticEventArea;
  code: string;
  id: number;
  level: DiagnosticEventLevel;
  message?: string;
  metadata: Record<string, boolean | null | number | string>;
  timestamp: string;
}

export interface DiagnosticSink {
  clear(): void;
  record(event: DiagnosticEventInput): DiagnosticEventRecord;
  snapshot(): DiagnosticEventRecord[];
}

interface MarkOutErrorOptions {
  cause?: unknown;
}

const DEFAULT_DIAGNOSTIC_CAPACITY = 50;
const MAX_DIAGNOSTIC_MESSAGE_LENGTH = 160;
const MAX_DIAGNOSTIC_METADATA_VALUE_LENGTH = 120;
const SENSITIVE_DIAGNOSTIC_METADATA_KEY_PATTERN =
  /(auth|body|cookie|draft|email|html|mail|markdown|message|password|recipient|secret|selection|session|state|storage|text|token)/i;
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
    commandsUrl:
      "https://schoenfeld-solutions.github.io/markout/outlook-beta/commands.html",
    launcheventUrl:
      "https://schoenfeld-solutions.github.io/markout/outlook-beta/launchevent.js",
    storageNamespace: "markout.beta",
    supportUrl: SUPPORT_URL,
    taskpaneUrl:
      "https://schoenfeld-solutions.github.io/markout/outlook-beta/taskpane.html",
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
    commandsUrl:
      "https://schoenfeld-solutions.github.io/markout/outlook/commands.html",
    launcheventUrl:
      "https://schoenfeld-solutions.github.io/markout/outlook/launchevent.js",
    storageNamespace: "markout.production",
    supportUrl: SUPPORT_URL,
    taskpaneUrl:
      "https://schoenfeld-solutions.github.io/markout/outlook/taskpane.html",
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

export function createInMemoryDiagnosticSink(
  capacity = DEFAULT_DIAGNOSTIC_CAPACITY,
  clock: () => Date = () => new Date()
): DiagnosticSink {
  return new InMemoryDiagnosticSink(capacity, clock);
}

export function getErrorDiagnosticMetadata(
  error: unknown
): DiagnosticEventMetadata {
  if (error instanceof MarkOutError) {
    return {
      errorCode: error.code,
      errorName: error.name,
    };
  }

  if (error instanceof Error) {
    return {
      errorName: error.name,
    };
  }

  return {
    errorName: "UnknownError",
  };
}

class InMemoryDiagnosticSink implements DiagnosticSink {
  private readonly clock: () => Date;
  private events: DiagnosticEventRecord[] = [];
  private nextId = 1;
  private readonly capacity: number;

  public constructor(capacity: number, clock: () => Date) {
    this.capacity = Math.max(1, Math.floor(capacity));
    this.clock = clock;
  }

  public clear(): void {
    this.events = [];
  }

  public record(event: DiagnosticEventInput): DiagnosticEventRecord {
    const record: DiagnosticEventRecord = {
      area: event.area,
      code: event.code,
      id: this.nextId,
      level: event.level,
      metadata: sanitizeDiagnosticMetadata(event.metadata),
      timestamp: this.clock().toISOString(),
    };

    if (event.message !== undefined) {
      record.message = truncateDiagnosticString(
        event.message,
        MAX_DIAGNOSTIC_MESSAGE_LENGTH
      );
    }

    this.nextId += 1;
    this.events = [...this.events, record].slice(-this.capacity);
    return record;
  }

  public snapshot(): DiagnosticEventRecord[] {
    return this.events.map((event) => ({
      ...event,
      metadata: { ...event.metadata },
    }));
  }
}

function sanitizeDiagnosticMetadata(
  metadata: DiagnosticEventMetadata | undefined
): Record<string, boolean | null | number | string> {
  if (metadata === undefined) {
    return {};
  }

  const normalizedMetadata: Record<string, boolean | null | number | string> =
    {};

  for (const [key, value] of Object.entries(metadata)) {
    if (value === undefined) {
      continue;
    }

    if (SENSITIVE_DIAGNOSTIC_METADATA_KEY_PATTERN.test(key)) {
      normalizedMetadata[key] = "[redacted]";
      continue;
    }

    normalizedMetadata[key] =
      typeof value === "string"
        ? truncateDiagnosticString(value, MAX_DIAGNOSTIC_METADATA_VALUE_LENGTH)
        : value;
  }

  return normalizedMetadata;
}

function truncateDiagnosticString(value: string, maxLength: number): string {
  const normalizedValue = value.replaceAll(/\s+/g, " ").trim();

  if (normalizedValue.length <= maxLength) {
    return normalizedValue;
  }

  return `${normalizedValue.slice(0, maxLength - 3).trimEnd()}...`;
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
