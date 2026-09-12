import { AxiError } from "axi-sdk-js";
import { PnpCliBackend, resolveBody } from "../backend.js";
import {
  parseFlags,
  flagString,
  flagBool,
  type FlagDef,
} from "../flags.js";

export interface RawContext {
  m365: PnpCliBackend;
  user?: string;
}

const RAW_FLAGS: Record<string, FlagDef> = {
  method: { type: "string", aliases: ["m"] },
  body: { type: "string", aliases: ["b"] },
  "content-type": { type: "string" },
  prefer: { type: "string" },
  query: { type: "string", aliases: ["q"] },
  execute: { type: "boolean" },
};

const METHODS = new Set(["get", "post", "put", "patch", "delete", "head", "options"]);
const WRITE_METHODS = new Set(["post", "put", "patch", "delete"]);

export async function raw(
  args: string[],
  context: RawContext,
): Promise<string | Record<string, unknown>> {
  const parsed = parseFlags(args, RAW_FLAGS);
  if (parsed.positionals.length !== 1) {
    throw new AxiError(
      "raw expects exactly one Graph path",
      "VALIDATION_ERROR",
      [
        'Example: msgraph-axi raw me/mailFolders/inbox?$select=unreadItemCount',
        "Or with a method: msgraph-axi raw me/events --method post --body '{...}'",
      ],
    );
  }
  const method = (flagString(parsed, "method", RAW_FLAGS) ?? "get").toLowerCase();
  if (!METHODS.has(method)) {
    throw new AxiError(
      `Unsupported method: ${method}`,
      "VALIDATION_ERROR",
      [`Valid methods: ${[...METHODS].join(", ")}`],
    );
  }
  const body = resolveBody(flagString(parsed, "body", RAW_FLAGS));
  if (body !== undefined && (method === "get" || method === "head" || method === "options")) {
    throw new AxiError(
      `--body is not supported for method ${method}`,
      "VALIDATION_ERROR",
      ["Use a write method (post, put, patch, delete) with --body"],
    );
  }
  if (WRITE_METHODS.has(method) && !flagBool(parsed, "execute", RAW_FLAGS)) {
    return {
      destructive: true,
      method,
      path: parsed.positionals[0],
      execute: false,
      help: ["Run with --execute to send the write request"],
    };
  }

  const path = parsed.positionals[0];
  const query = flagString(parsed, "query", RAW_FLAGS);
  const url = `@graph/${path}${query ? `?${query}` : ""}`;
  const requestArgs = ["request", "--method", method, "--url", url];
  if (body !== undefined) {
    requestArgs.push(
      "--body",
      body,
      "--content-type",
      flagString(parsed, "content-type", RAW_FLAGS) ?? "application/json",
    );
  }
  const prefer = flagString(parsed, "prefer", RAW_FLAGS);
  if (prefer !== undefined) {
    requestArgs.push("--prefer", prefer);
  }

  const result = await context.m365.run(requestArgs);
  if (result.stdout.trim() === "") {
    return { status: "ok", method, path };
  }
  const parsedJson = JSON.parse(result.stdout) as unknown;
  return parsedJson as string | Record<string, unknown>;
}