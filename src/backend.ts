import { spawn } from "cross-spawn";
import { AxiError } from "axi-sdk-js";
import { readFileSync } from "node:fs";

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

export class PnpCliBackend {
  private userCache: Promise<string | undefined> | undefined;

  constructor(
    private readonly bin: string = process.env.MSGRAPH_AXI_M365_BIN ?? "m365",
    private readonly prefix: string[] = [],
    private readonly env: Record<string, string> | undefined = undefined,
  ) {}

  async run(args: string[]): Promise<RunResult> {
    const full = [...this.prefix, ...args, "--output", "json"];
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

function m365Error(stderr: string, stdout: string): AxiError {
  const message =
    extractError(stderr) ??
    extractError(stdout) ??
    (stderr.trim() || stdout.trim());
  if (/logged out|not signed in|no valid connection/i.test(message)) {
    return new AxiError(
      "m365: not signed in",
      "NOT_SIGNED_IN",
      ["Run `msgraph-axi auth login` to sign in"],
    );
  }
  return new AxiError(
    `m365: ${message}`,
    "M365_ERROR",
    ["Run `msgraph-axi auth status` to verify the connection"],
  );
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