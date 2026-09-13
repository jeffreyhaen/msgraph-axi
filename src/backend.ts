import { spawn } from "cross-spawn";
import { AxiError } from "axi-sdk-js";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { hostTimeZone } from "./time.js";

export interface RunResult {
  stdout: string;
  code: number;
}

export interface M365Status {
  connectionName?: string;
  connectedAs?: string;
  authType?: string;
  appId?: string;
  appTenant?: string;
  cloudType?: string;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** cmd.exe refuses anything longer; the real limit is 8191 characters. */
const CMD_LINE_LIMIT = 8000;

/**
 * On Windows an npm-installed CLI is spawned through its `.cmd` shim, so the
 * arguments travel via cmd.exe, which truncates a value at the first line break
 * and caps the whole command line — both silently. Refuse those values instead
 * of sending corrupted data to Graph.
 */
export function assertCmdlineSafeArgs(
  args: string[],
  platform: string = process.platform,
): void {
  if (platform !== "win32") {
    return;
  }
  for (const arg of args) {
    if (/[\r\n]/.test(arg)) {
      throw new AxiError(
        "An argument contains a line break; on Windows it would be cut off at the first line break",
        "VALIDATION_ERROR",
        [
          "Pass the payload as a file: --body @payload.json (the backend reads @file itself)",
          "JSON payloads this CLI builds itself are sent through a temp file automatically",
        ],
      );
    }
    if (arg.length > CMD_LINE_LIMIT) {
      throw new AxiError(
        `An argument is ${arg.length} characters, beyond the ${CMD_LINE_LIMIT}-character Windows limit`,
        "VALIDATION_ERROR",
        [
          "Pass the payload as a file: --body @payload.json (the backend reads @file itself)",
          "JSON payloads this CLI builds itself are sent through a temp file automatically",
        ],
      );
    }
  }
}

export interface BodyArgument {
  /** The argument to pass as `--body`: the value itself, or `@<tempfile>`. */
  value: string;
  cleanup: () => void;
}

/**
 * A `--body` value that survives the platform. On Windows every payload travels
 * as a file: cmd.exe mangles an argument that mixes quotes with shell
 * metacharacters (`<`, `>`, `&`, `|`, `^`, `%`, `!`) or a line break, and JSON
 * always carries quotes. Elsewhere argv is passed verbatim.
 *
 * `bodyDir` keeps the payload on disk for inspection (tests, debugging);
 * without it the file is temporary and removed after the request.
 */
export function bodyArgument(
  contents: string,
  platform: string = process.platform,
  bodyDir?: string,
): BodyArgument {
  if (platform !== "win32") {
    return { value: contents, cleanup: () => {} };
  }
  if (bodyDir !== undefined) {
    const file = join(bodyDir, "body.json");
    writeFileSync(file, contents, "utf8");
    return { value: `@${file}`, cleanup: () => {} };
  }
  const dir = mkdtempSync(join(tmpdir(), "msgraph-axi-body-"));
  const file = join(dir, "body.json");
  writeFileSync(file, contents, "utf8");
  return {
    value: `@${file}`,
    cleanup: () => rmSync(dir, { recursive: true, force: true }),
  };
}

export interface BackendOptions {
  /**
   * Keep payload files in this directory instead of a temp dir. Used by tests
   * and debugging: the payload stays readable after the request finished.
   */
  bodyDir?: string;
}

export class PnpCliBackend {
  private userCache: Promise<string | undefined> | undefined;
  private timeZoneCache = new Map<string, Promise<string>>();

  constructor(
    private readonly bin: string = process.env.MSGRAPH_AXI_M365_BIN ?? "m365",
    private readonly prefix: string[] = [],
    private readonly env: Record<string, string> | undefined = undefined,
    private readonly options: BackendOptions = {},
  ) {}

  async run(args: string[]): Promise<RunResult> {
    const full = [...this.prefix, ...args, "--output", "json"];
    assertCmdlineSafeArgs(full);
    const spawnEnv =
      this.env === undefined
        ? undefined
        : { ...process.env, ...this.env };
    return new Promise<RunResult>((resolve, reject) => {
      let settled = false;
      const fail = (error: unknown): void => {
        if (!settled) {
          settled = true;
          reject(error);
        }
      };
      const child = spawn(this.bin, full, {
        stdio: ["ignore", "pipe", "pipe"],
        env: spawnEnv,
      });
      let out = "";
      let err = "";
      child.stdout?.on("data", (d: Buffer) => {
        out += d.toString();
      });
      child.stderr?.on("data", (d: Buffer) => {
        err += d.toString();
      });
      child.on("error", (e: NodeJS.ErrnoException) => {
        fail(
          new AxiError(
            `m365 could not be started: ${e.message}`,
            "M365_NOT_FOUND",
            [
              "Install the CLI for Microsoft 365: npm i -g @pnp/cli-microsoft365",
              "Or point MSGRAPH_AXI_M365_BIN at the binary",
            ],
          ),
        );
      });
      child.on("close", (code) => {
        if (settled) {
          return;
        }
        if (code !== 0) {
          fail(m365Error(err, out));
          return;
        }
        settled = true;
        resolve({ stdout: out, code: code ?? 0 });
      });
    });
  }

  /**
   * Run a Graph request with an optional JSON payload. Payloads too long for a
   * Windows command line travel as a `@file` argument instead.
   */
  async request<T>(method: string, url: string, payload?: unknown): Promise<T> {
    const args = ["request", "--method", method, "--url", url];
    const argument =
      payload === undefined
        ? undefined
        : bodyArgument(JSON.stringify(payload), process.platform, this.options.bodyDir);
    if (argument !== undefined) {
      args.push("--body", argument.value, "--content-type", "application/json");
    }
    try {
      return await this.runJson<T>(args);
    } finally {
      argument?.cleanup();
    }
  }

  async runJsonArray<T>(args: string[]): Promise<T[]> {
    const result = await this.run(args);
    const parsed = JSON.parse(result.stdout) as unknown;
    if (!Array.isArray(parsed)) {
      throw new AxiError(
        `Expected a JSON array from m365 ${args[0]}, got ${typeof parsed}`,
        "M365_UNEXPECTED_OUTPUT",
        ["Inspect the m365 CLI output for this command"],
      );
    }
    return parsed as T[];
  }

  async runJson<T>(args: string[]): Promise<T> {
    const result = await this.run(args);
    return JSON.parse(result.stdout) as T;
  }

  /**
   * Interactive login (device code flow prints the code to the terminal). The
   * SDK's no-interactive-prompts rule intentionally leaves room for this: it is
   * an explicit user action, not an agent-driven prompt.
   */
  async login(authType: string, extraArgs: string[] = []): Promise<void> {
    const { spawn } = await import("cross-spawn");
    const result = await new Promise<number>((resolve, reject) => {
      const child = spawn(
        this.bin,
        [...this.prefix, "login", "--authType", authType, ...extraArgs, "--output", "json"],
        { stdio: "inherit" },
      );
      child.on("error", (e: NodeJS.ErrnoException) => {
        reject(
          new AxiError(
            `m365 could not be started: ${e.message}`,
            "M365_NOT_FOUND",
            ["Install: npm i -g @pnp/cli-microsoft365"],
          ),
        );
      });
      child.on("close", (code) => resolve(code ?? 0));
    });
    if (result !== 0) {
      throw new AxiError(
        "m365 login did not complete",
        "M365_ERROR",
        ["Run `msgraph-axi auth login` again and complete the flow"],
      );
    }
  }

  async status(): Promise<M365Status | undefined> {
    let result: RunResult;
    try {
      result = await this.run(["status"]);
    } catch (error) {
      if (error instanceof AxiError && error.code === "M365_ERROR") {
        return undefined;
      }
      throw error;
    }
    const parsed = JSON.parse(result.stdout) as unknown;
    if (typeof parsed === "string") {
      return undefined;
    }
    return parsed as M365Status;
  }

  currentUser(): Promise<string | undefined> {
    if (!this.userCache) {
      this.userCache = this.status().then((s) => s?.connectedAs);
    }
    return this.userCache;
  }

  /**
   * Resolve the time zone a command should work in: an explicit `--timezone`
   * wins, then the mailbox setting from Graph, then this machine's zone. The
   * mailbox lookup needs `MailboxSettings.Read`; without it the machine zone is
   * used so times are never silently read or written as UTC.
   */
  timeZone(userFlag: string | undefined, explicit?: string): Promise<string> {
    if (explicit !== undefined && explicit.trim() !== "") {
      return Promise.resolve(explicit.trim());
    }
    const key = userFlag ?? "";
    let cached = this.timeZoneCache.get(key);
    if (!cached) {
      cached = this.fetchMailboxTimeZone(userFlag).catch(() => hostTimeZone());
      this.timeZoneCache.set(key, cached);
    }
    return cached;
  }

  private async fetchMailboxTimeZone(userFlag: string | undefined): Promise<string> {
    const user = (await this.userArgs(userFlag))[1] ?? "me";
    const settings = await this.runJson<{ timeZone?: string }>([
      "request",
      "--method",
      "get",
      "--url",
      `@graph/users/${user}/mailboxSettings?$select=timeZone`,
    ]);
    const zone = settings.timeZone?.trim();
    return zone ? zone : hostTimeZone();
  }

  /**
   * Build the `--userName`/`--userId` pair: an explicit `--user` that looks
   * like a GUID is passed as `--userId`, anything else as `--userName`; without
   * `--user` the signed-in account is resolved from `m365 status`.
   */
  async userArgs(userFlag: string | undefined): Promise<string[]> {
    if (userFlag !== undefined) {
      const def = GUID_RE.test(userFlag) ? "--userId" : "--userName";
      return [def, userFlag];
    }
    const user = await this.currentUser();
    if (user) {
      return ["--userName", user];
    }
    throw new AxiError(
      "No signed-in account; a mailbox context is required",
      "NOT_SIGNED_IN",
      [
        "Run `msgraph-axi auth login` to sign in",
        "Or pass --user <upn> explicitly",
      ],
    );
  }
}

/**
 * Resolve a body value: `@<path>` reads the file at `<path>` (Git Bash note:
 * prefix with MSYS_NO_PATHCONV=1 on Windows when the path starts with `/`).
 */
export function resolveBody(body: string | undefined): string | undefined {
  if (body === undefined) {
    return undefined;
  }
  if (body.startsWith("@")) {
    return readFileSync(body.slice(1), "utf8");
  }
  return body;
}

const SIGNED_OUT_RE = /logged out|not signed in|no valid (connection|login)/i;
const PERMISSION_RE = /403|forbidden|access denied|insufficient|does not have permission|consent/i;
const AUTH_RE = /401|unauthor|invalid token|token (has )?expired|authentication|interactive/i;
const NOT_FOUND_RE = /404|not found|itemnotfound|does not exist/i;
const INPUT_RE = /not a valid|invalid|must be|unrecognized|unexpected|unknown (flag|command)|parameter|expects/i;

/**
 * Map a backend failure to an actionable error. The hint must match the cause:
 * telling the caller to check its connection after a bad `--start` value sends
 * agents down the wrong path.
 */
function m365Error(stderr: string, stdout: string): AxiError {
  const message =
    extractError(stderr) ??
    extractError(stdout) ??
    (stderr.trim() || stdout.trim());
  if (SIGNED_OUT_RE.test(message)) {
    return new AxiError(
      "m365: not signed in",
      "NOT_SIGNED_IN",
      ["Run `msgraph-axi auth login` to sign in"],
    );
  }
  if (PERMISSION_RE.test(message)) {
    return new AxiError(`m365: ${message}`, "M365_ERROR", [
      "The signed-in account is missing a Graph permission for this call",
      "See README.md for the required scope, then re-consent the app",
    ]);
  }
  if (AUTH_RE.test(message)) {
    return new AxiError(`m365: ${message}`, "M365_ERROR", [
      "The session looks expired: run `msgraph-axi auth login`",
      "Run `msgraph-axi auth status` to check the connection",
    ]);
  }
  if (NOT_FOUND_RE.test(message)) {
    return new AxiError(`m365: ${message}`, "M365_ERROR", [
      "Check that the id or upn exists and is visible to the signed-in account",
    ]);
  }
  if (INPUT_RE.test(message)) {
    return new AxiError(`m365: ${message}`, "M365_ERROR", [
      "This is an input error, not a connection problem: check the flag values",
      "Run `msgraph-axi <command> --help` for the accepted format",
    ]);
  }
  return new AxiError(`m365: ${message}`, "M365_ERROR", [
    "Inspect the failure with the m365 CLI directly for the raw response",
    "Run `msgraph-axi <command> --help` for the accepted flags",
  ]);
}

function extractError(text: string): string | undefined {
  const match = text.match(/(?:^|\n)error:\s*(.+)$/im);
  if (match) {
    return match[1].trim();
  }
  try {
    const json = JSON.parse(text) as Record<string, unknown> | undefined;
    if (json && typeof json.error === "string") {
      return json.error;
    }
  } catch {
    // not JSON — fall through to the raw text
  }
  return undefined;
}