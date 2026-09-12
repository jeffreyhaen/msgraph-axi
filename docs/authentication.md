# Authentication

`msgraph-axi` delegates authentication to the CLI for Microsoft 365 (`m365`) and
reads the sign-in state through `msgraph-axi auth status`. The m365 CLI stores
the acquired tokens itself, so a login survives across msgraph-axi invocations —
no credentials are ever stored by msgraph-axi itself.

The default flow is **device code** (interactive, works with MFA):

```sh
msgraph-axi auth login
```

`auth login --auth-type <type>` accepts every `m365 login` auth type. All flags
of `m365 login` are available with a `--` prefix translation (`--appId` →
`--app-id`, `--certificateFile` → `--cert-file`, `--secret`, `--userName` →
`--user-name`, `--password`).

| Flow | Command | Use when |
|---|---|---|
| Device code | `msgraph-axi auth login` | a human at a terminal; MFA OK (default) |
| Browser | `msgraph-axi auth login --auth-type browser` | a human; interactive popup |
| Client secret | `msgraph-axi auth login --auth-type secret --app-id <id> --tenant <id> --secret <s>` | headless service principal |
| Certificate | `msgraph-axi auth login --auth-type certificate --app-id <id> --tenant <id> --cert-file <path.pem>` | headless service principal, key in a certificate |
| Certificate (inline) | `msgraph-axi auth login --auth-type certificate --app-id <id> --tenant <id> --cert-base64 <value> [--thumbprint <t>]` | certificates injected via secret stores |
| Username/password | `msgraph-axi auth login --auth-type password --app-id <id> --tenant <id> --user-name <upn> --password <p>` | automation with a regular account; **no MFA** (ROPC) |
| Managed identity | `msgraph-axi auth login --auth-type identity` | Azure VM/App Service / Functions with a system-assigned identity |
| Federated identity | `msgraph-axi auth login --auth-type federatedIdentity` | workload identity federation (GitHub Actions, AKS) |

## Notes

- **Service principal flows need an app registration**: create one in Entra ID,
  grant it the Graph permissions the commands need (delegated or application),
  and use its application (client) id as `--app-id`.
- `--tenant` accepts the tenant id or a verified domain name
  (`contoso.onmicrosoft.com`).
- Client-secret and certificate secrets should come from a secret store /
  environment — never paste them into prompts or shared shells:
  ```sh
  msgraph-axi auth login --auth-type secret --app-id $APP_ID --tenant $TENANT --secret "$(cat .secret)"
  ```
  Use `MSYS_NO_PATHCONV=1` before the command on Git Bash when `--cert-file`
  points at a path starting with `/`.
- Check the applied flow with `msgraph-axi auth status`; switch accounts with
  `auth logout` followed by a new `auth login`.
- The exact behavior of each flow (including token storage locations) is defined
  by the [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/cmd/login/).