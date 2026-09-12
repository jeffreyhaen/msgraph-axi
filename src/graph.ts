import { AxiError } from "axi-sdk-js";
import { PnpCliBackend } from "./backend.js";

export interface GraphContext {
  m365: PnpCliBackend;
  user?: string;
}

/**
 * Resolve the signed-in user (or the explicit --user) and return the
 * `@graph/users/<upn>` prefix used by the raw request bridge.
 */
export async function graphPrefix(
  context: GraphContext,
  userFlag: string | undefined,
): Promise<string> {
  const userArgs = await context.m365.userArgs(userFlag);
  const user = userArgs[1] ?? "me";
  return `@graph/users/${user}`;
}

/**
 * GET a Graph collection that returns `{ value: [...] }` (or a bare array) and
 * normalize it to an array.
 */
export async function getGraphValue<T>(
  context: GraphContext,
  userFlag: string | undefined,
  pathAndQuery: string,
): Promise<T[]> {
  const prefix = await graphPrefix(context, userFlag);
  const response = await context.m365.runJson<unknown>([
    "request",
    "--method",
    "get",
    "--url",
    `${prefix}${pathAndQuery}`,
  ]);
  if (Array.isArray(response)) {
    return response as T[];
  }
  const value = (response as { value?: unknown }).value;
  if (!Array.isArray(value)) {
    throw new AxiError(
      "Expected a value[] collection from Graph",
      "M365_UNEXPECTED_OUTPUT",
      ["Inspect the Graph response for this endpoint"],
    );
  }
  return value as T[];
}

/** URL-encode a single query-string value (e.g. an OData $search expression). */
export function odataValue(value: string): string {
  return encodeURIComponent(value);
}

/** Build a `?$search="<q>"&$top=N&$select=...` suffix from parts. */
export function searchQuery(search: string, top: number, select: string[]): string {
  const params = [
    `$search=${odataValue(`"${search}"`)}`,
    `$top=${top}`,
    `$select=${select.join(",")}`,
  ];
  return `?${params.join("&")}`;
}