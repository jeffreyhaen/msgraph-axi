import { AxiError } from "axi-sdk-js";
import { PnpCliBackend } from "../backend.js";
import { graphPrefix } from "../graph.js";
import { parseFlags, type FlagDef } from "../flags.js";
import { cell } from "../toon.js";

export interface UserContext {
  m365: PnpCliBackend;
  user?: string;
}

const GET_FLAGS: Record<string, FlagDef> = {
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