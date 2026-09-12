import { AxiError } from "axi-sdk-js";
import { PnpCliBackend, type M365Status } from "../backend.js";
import { parseFlags, type FlagDef } from "../flags.js";

export interface AuthContext {
  m365: PnpCliBackend;
}

const AUTH_FLAGS: Record<string, FlagDef> = {
  "auth-type": { type: "string" },
  "app-id": { type: "string" },
  tenant: { type: "string" },
  "cert-file": { type: "string" },
  "cert-base64": { type: "string" },
  thumbprint: { type: "string" },
  secret: { type: "string" },
  "user-name": { type: "string" },
  password: { type: "string" },
};

const PNP_LOGIN_FLAG: Record<string, string> = {
  "app-id": "--appId",
  tenant: "--tenant",
  "cert-file": "--certificateFile",
  "cert-base64": "--certificateBase64Encoded",
  thumbprint: "--thumbprint",
  secret: "--secret",
  "user-name": "--userName",
  password: "--password",
};

function loginExtraArgs(
  flags: Record<string, string | number | boolean>,
): string[] {
  const extra: string[] = [];
  for (const [axi, pnp] of Object.entries(PNP_LOGIN_FLAG)) {
    const value = flags[axi];
    if (typeof value === "string") {
      extra.push(pnp, value);
    }
  }
  return extra;
}

function requireLoginFlags(
  authType: string,
  flags: Record<string, string | number | boolean>,
  required: string[],
): void {
  const missing = required.filter((f) => typeof flags[f] !== "string");
  if (missing.length > 0) {
    throw new AxiError(
      `auth login --auth-type ${authType} requires ${missing
        .map((f) => `--${f}`)
        .join(", ")}`,
      "VALIDATION_ERROR",
      [
        `Example: msgraph-axi auth login --auth-type ${authType} ${required
          .map((f) => `--${f} <value>`)
          .join(" ")}`,
      ],
    );
  }
}

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

const REQUIRED_BY_TYPE: Record<string, string[]> = {
  certificate: ["app-id", "tenant"],
  secret: ["app-id", "tenant", "secret"],
  password: ["app-id", "tenant", "user-name", "password"],
};

export async function authLogin(
  args: string[],
  context: AuthContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, AUTH_FLAGS);
  const flags = parsed.flags;
  const authType =
    typeof flags["auth-type"] === "string"
      ? (flags["auth-type"] as string)
      : "deviceCode";
  const required = REQUIRED_BY_TYPE[authType];
  if (required) {
    requireLoginFlags(authType, flags, required);
    if (authType === "certificate") {
      const hasCert =
        typeof flags["cert-file"] === "string" ||
        typeof flags["cert-base64"] === "string";
      if (!hasCert) {
        throw new AxiError(
          "auth login --auth-type certificate requires --cert-file or --cert-base64",
          "VALIDATION_ERROR",
          ["Pass --cert-file <path.pem> or --cert-base64 <value>"],
        );
      }
    }
  }
  const extra = loginExtraArgs(flags);
  await context.m365.login(authType, extra);
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