import { AxiError } from "axi-sdk-js";
import { PnpCliBackend, type M365Status } from "../backend.js";
import { parseFlags, type FlagDef } from "../flags.js";

export interface AuthContext {
  m365: PnpCliBackend;
}

const AUTH_FLAGS: Record<string, FlagDef> = {
  "auth-type": { type: "string" },
};

export async function authStatus(
  args: string[],
  context: AuthContext,
): Promise<Record<string, unknown>> {
  parseFlags(args, {});
  const status = await context.m365.status();
  if (!status) {
    return {
      signedIn: false,
      help: ["Run `msgraph-axi auth login` to sign in"],
    };
  }
  return {
    signedIn: true,
    connectedAs: status.connectedAs ?? "",
    authType: status.authType ?? "",
    appTenant: status.appTenant ?? "",
    appId: status.appId ?? "",
    cloud: status.cloudType ?? "",
  };
}

export async function authLogin(
  args: string[],
  context: AuthContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, AUTH_FLAGS);
  const authType =
    typeof parsed.flags["auth-type"] === "string"
      ? (parsed.flags["auth-type"] as string)
      : "deviceCode";
  await context.m365.login(authType);
  const status: M365Status | undefined = await context.m365.status();
  return {
    signedIn: true,
    connectedAs: status?.connectedAs ?? "",
    authType,
  };
}

export async function authLogout(
  args: string[],
  context: AuthContext,
): Promise<Record<string, unknown>> {
  parseFlags(args, {});
  try {
    await context.m365.run(["logout"]);
  } catch (error) {
    if (
      error instanceof AxiError &&
      /logged out|not signed in/i.test(error.message)
    ) {
      return { signedOut: true, wasSignedIn: false };
    }
    throw error;
  }
  return { signedOut: true, wasSignedIn: true };
}