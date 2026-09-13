import { AxiError } from "axi-sdk-js";
import { PnpCliBackend } from "../backend.js";
import { odataValue } from "../graph.js";
import { parseFlags, flagNumber, type FlagDef } from "../flags.js";
import { cell } from "../toon.js";

export interface UserContext {
  m365: PnpCliBackend;
  user?: string;
}

const GET_FLAGS: Record<string, FlagDef> = {
  full: { type: "boolean" },
};

const SEARCH_FLAGS: Record<string, FlagDef> = {
  limit: { type: "number" },
  full: { type: "boolean" },
};

interface GraphUser {
  displayName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mail?: string;
  userPrincipalName?: string;
  businessPhones?: string[];
  mobilePhone?: string | null;
}

interface GraphUserRow {
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
}

/**
 * Directory search by display-name, first/last-name, mail or upn prefix.
 * Read-only. Graph `$search` needs a `ConsistencyLevel` header that the request
 * bridge does not send, so this uses `startswith` filters.
 */
export async function userSearch(
  args: string[],
  context: UserContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, SEARCH_FLAGS);
  if (parsed.positionals.length !== 1) {
    throw new AxiError(
      "user search expects exactly one search term",
      "VALIDATION_ERROR",
      [
        'Example: msgraph-axi user search alex   (a name, mail or upn prefix)',
        "Search terms with spaces: try just the first or last name",
      ],
    );
  }
  const term = parsed.positionals[0].trim();
  if (term === "") {
    throw new AxiError("user search needs a non-empty term", "VALIDATION_ERROR", [
      'Example: msgraph-axi user search alex',
    ]);
  }
  const full = parsed.flags["full"] === true;
  const limit = Math.floor(flagNumber(parsed, "limit", SEARCH_FLAGS, 15));
  const escaped = term.replace(/'/g, "''");
  const filter = [
    `startswith(displayName,'${escaped}')`,
    `startswith(givenName,'${escaped}')`,
    `startswith(surname,'${escaped}')`,
    `startswith(mail,'${escaped}')`,
    `startswith(userPrincipalName,'${escaped}')`,
  ].join(" or ");
  const select = "displayName,mail,userPrincipalName,jobTitle,department";
  const response = await context.m365.runJson<
    GraphUserRow[] | { value?: GraphUserRow[] }
  >([
    "request",
    "--method",
    "get",
    "--url",
    `@graph/users?$filter=${odataValue(filter)}&$select=${select}&$top=${limit}`,
  ]);
  const list = Array.isArray(response) ? response : (response.value ?? []);
  const users = list.map((user) => ({
    name: cell(user.displayName ?? "", full),
    upn: user.userPrincipalName ?? "",
    mail: user.mail ?? "",
    jobTitle: cell(user.jobTitle ?? "", full),
  }));
  const out: Record<string, unknown> = { users, count: users.length };
  if (users.length === 0) {
    out.help = [
      `No directory user starts with "${term}"`,
      "Try fewer letters, or the mail address / upn prefix",
    ];
  } else if (users.length >= limit) {
    out.truncated = true;
    out.help = [`Showing the first ${limit} matches; raise --limit for more`];
  }
  return out;
}

/** People and organizational lookup via Graph `/users`. Read-only. */
export async function userGet(
  args: string[],
  context: UserContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, GET_FLAGS);
  const positionals = parsed.positionals;
  if (positionals.length !== 1) {
    throw new AxiError(
      "user get expects exactly one user principal name",
      "VALIDATION_ERROR",
      ["Run `msgraph-axi user get <upn>`"],
    );
  }
  const upn = positionals[0];
  const full = parsed.flags["full"] === true;
  const select =
    "displayName,jobTitle,department,officeLocation,mail,userPrincipalName,businessPhones,mobilePhone";
  const user = await context.m365.runJson<GraphUser>([
    "request",
    "--method",
    "get",
    "--url",
    `@graph/users/${upn}?$select=${select}`,
  ]);
  const phones = [
    ...(user.businessPhones ?? []),
    ...(user.mobilePhone ? [user.mobilePhone] : []),
  ];
  const out: Record<string, unknown> = {
    name: user.displayName ?? "",
    jobTitle: cell(user.jobTitle ?? "", full),
    department: cell(user.department ?? "", full),
    location: cell(user.officeLocation ?? "", full),
    mail: user.mail ?? user.userPrincipalName ?? "",
    upn: user.userPrincipalName ?? upn,
    phone: phones[0] ?? "",
  };
  let manager: { displayName?: string; mail?: string } | undefined;
  try {
    manager = await context.m365.runJson<{ displayName?: string; mail?: string }>([
      "request",
      "--method",
      "get",
      "--url",
      `@graph/users/${upn}/manager?$select=displayName,mail`,
    ]);
  } catch {
    manager = undefined;
  }
  out.manager = manager
    ? manager.displayName
      ? `${manager.displayName} <${manager.mail ?? ""}>`
      : (manager.mail ?? "")
    : "";
  return out;
}