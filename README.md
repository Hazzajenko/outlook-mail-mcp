# outlook-mail-mcp

Query personal Outlook via Microsoft Graph. CLI + MCP server, sharing one core. Read-only.

## Stack

TypeScript / Node 22+ / pnpm. Auth via `@azure/msal-node` with device code flow, a public client, and no secret. Graph access via hand-rolled `fetch` wrapper with Zod-validated responses. Tests: Vitest + MSW.

## Setup

### 1. Install

Install the package globally to get the `outlook-mail` CLI:

```
npm install -g @hazzajenko/outlook-mail-mcp
```

To run the CLI with no install, put `npx -p @hazzajenko/outlook-mail-mcp`
before each command. For example:
`npx -p @hazzajenko/outlook-mail-mcp outlook-mail setup`. The MCP server
needs no install. Your MCP host starts it with `npx`, as shown in the MCP
section below.

### 2. Run guided setup

```
outlook-mail setup
```

The command walks you through registering the Entra app. Microsoft does not
allow this to be automated for personal accounts. The API returns
`403 not authorized to create apps using consumer identity`, see
`docs/adr/0002`. The command prints the portal steps, asks for the resulting
**Application (client) ID**, writes `.env` in the current directory without
touching other keys, prints the values, and then **verifies** the
registration with a device-code sign-in. A successful verification caches the
token, so setup doubles as first-run auth. No separate `auth` step is needed.
If verification fails, the error tells you which portal step to revisit.

Flags:

- `--personal-only`: you picked **Personal Microsoft accounts only** in step 3
  of the walkthrough. It writes `OUTLOOK_MAIL_TENANT_ID=consumers` to match. The
  tenant value is always derived from your audience choice, so the
  `AADSTS9002331` mismatch cannot happen on this path.
- `--client-id <id>`: skip the prompt, for non-interactive or agent use.
- `--no-verify`: skip the verification sign-in. Run `outlook-mail auth` later.

Env can also live in your shell or an MCP `env` block instead of `.env`. The
values are always printed for copy-paste. The CLI auto-loads `.env` from cwd
via Node's built-in `process.loadEnvFile()`. The MCP server does too, but
the MCP host starts it from its own cwd. Put the values in the `env` block of
your MCP config, as shown in the MCP section below.

### Manual setup

Use this if the guided flow fails. The step numbers below match the walkthrough and its error messages.

1. <https://entra.microsoft.com> → Identity → Applications → App registrations → **New registration**.
2. Name: anything, for example `outlook-mail`.
3. Supported account types: **Accounts in any org + personal** (default), or **Personal Microsoft accounts only**. With the second, use `--personal-only` or `OUTLOOK_MAIL_TENANT_ID=consumers`.
4. Redirect URI: leave blank. Device code flow needs none. Register.
5. On the new app's **Overview** page, copy the **Application (client) ID**.
6. **Authentication** → **Allow public client flows** → **Yes** → Save.
7. **API permissions** → Add → Microsoft Graph → Delegated permissions → check `Mail.Read` and `offline_access` → Add. Personal accounts have no admin consent. You accept the consent prompt on first sign-in. Work or school tenants can click **Grant admin consent**.

Then create `.env` in the directory where you run the CLI, or export the values in your shell. Example:

```
OUTLOOK_MAIL_CLIENT_ID=<paste app id>
# Optional. Default 'common' (personal + work). Use 'consumers' to lock to personal.
OUTLOOK_MAIL_TENANT_ID=common
```

**If you chose "Personal Microsoft accounts only" in step 3, you must set
`OUTLOOK_MAIL_TENANT_ID=consumers`.** Such apps reject the default `common`
authority with `AADSTS9002331: Application is configured for use by Microsoft
Account users only. Please use the /consumers endpoint`.

Finally authenticate:

```
outlook-mail auth
```

Opens a device-code message in stderr. Visit the URL, enter the code, sign in. Token cached at `~/.config/outlook-mail/msal-cache.json` (mode 0600). Subsequent runs are silent.

## CLI

```
outlook-mail search --from example.com --since -30d
outlook-mail search -q "interview" --unread --top 100
outlook-mail search --folder Jobs --has-attachment --json
outlook-mail get <message-id>
outlook-mail folders
```

Flags: `-q/--query`, `--from`, `--to`, `--subject`, `--body`, `--since`, `--until`, `--has-attachment`, `--unread` / `--read`, `--folder`, `--importance`, `--inference-classification` (`focused`/`other`), `--top`, `--json`. Dates: ISO (`2026-05-01`) or relative (`-7d`, `-2w`, `-3h`, `-30m`).

`--inference-classification` needs a `--since`/`--until` alongside it. Results are sorted by `receivedDateTime`, and Graph rejects an unbounded inference filter combined with that sort (400 `InefficientFilter`). Same applies to the `list_emails_brief` MCP tool. `count_emails` doesn't sort and is exempt. Omitting the bound raises an error that says so. The raw Graph 400 names neither the cause nor the fix.

## MCP

Run `outlook-mail setup` first. The server reads the token that setup caches,
and it cannot show a device-code prompt itself.

For Claude Code:

```
claude mcp add outlook-mail --env OUTLOOK_MAIL_CLIENT_ID=<your-client-id> -- npx -y @hazzajenko/outlook-mail-mcp
```

For other hosts, such as Claude Desktop, add this to the MCP config:

```json
{
  "mcpServers": {
    "outlook-mail": {
      "command": "npx",
      "args": ["-y", "@hazzajenko/outlook-mail-mcp"],
      "env": {
        "OUTLOOK_MAIL_CLIENT_ID": "<your-client-id>"
      }
    }
  }
}
```

If you used `--personal-only`, also set `OUTLOOK_MAIL_TENANT_ID` to
`consumers` in the `env` block. On Windows, some hosts cannot start `npx`
directly. Use `"command": "cmd"` with `"args": ["/c", "npx", "-y", "@hazzajenko/outlook-mail-mcp"]`.

Tools exposed: `search_emails`, `get_email`, `get_conversation`, `count_emails`, `list_emails_brief`, `list_folders`. Same schema as the CLI for the search filter.

## Troubleshooting

**`AADSTS9002331` during auth.** Your app registration is "Personal Microsoft
accounts only" but the tenant is `common`. Set `OUTLOOK_MAIL_TENANT_ID=consumers`,
or re-run `outlook-mail setup --personal-only`, which derives it for you.

**MCP tools don't show up in the client.** Fully quit and relaunch the host,
then check the host's MCP logs. Claude Desktop does not hot-reload config. To
smoke-test the server directly, run:

```
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"x","version":"0"}}}' | npx -y @hazzajenko/outlook-mail-mcp
```

A JSON response should print within a few seconds.

**`OUTLOOK_MAIL_CLIENT_ID env var not set` at tool-call time.** The host's
working directory doesn't contain your `.env`. Pass the variable via the MCP
config's `env` block.

**Token expired or 401 from Graph.** Re-run `outlook-mail auth`.
Silent refresh usually works. If it fails, you get a device-code prompt.

## Development

To run from source, clone the repo and build it:

```
pnpm install
pnpm build
pnpm exec outlook-mail setup
```

Point the MCP config at the build with `"command": "node"` and
`"args": ["/abs/path/to/outlook-mail-mcp/dist/mcp.js"]`. Run `pnpm build`
after each change so the host loads the new code.

```
pnpm test         # vitest run
pnpm test:watch
pnpm typecheck
pnpm fix          # biome check --write (lint + format + organize imports)
pnpm check        # fix + typecheck + test (full local CI)
pnpm smoke        # live Graph smoke test (real mailbox, read-only)
```

### Smoke test

`pnpm check` only proves the code agrees with its own fixtures. It cannot catch
Graph returning a shape our schemas reject. One example is a `$select` that
stops requesting a field a Zod schema still marks required. That breaks a tool
completely while the suite stays green.

`pnpm smoke` calls every exposed operation against the real mailbox and parses
each response through its declared result schema (`SearchResultSchema`,
`FullMessageSchema`, …). It needs `OUTLOOK_MAIL_CLIENT_ID` and a cached token.
It is read-only and exits non-zero on any failure. Cases that need data the mailbox
doesn't have report `SKIP` rather than passing silently.

Known Graph limitation it encodes: `inference_classification` combined with
`$orderby receivedDateTime desc` returns **400 InefficientFilter** unless the
filter also bounds `receivedDateTime`. Pass `--since`/`--until` alongside
`--inference-classification`.

## Releasing

Releases use [release-please](https://github.com/googleapis/release-please).
Each push to main updates one open release PR. That PR bumps the version in
`package.json` and adds the new commits to `CHANGELOG.md`. Commit subjects set
the bump: `fix:` gives a patch, `feat:` gives a minor, and a breaking change
gives a minor while the version is below 1.0.

To release, merge the release PR. release-please then tags `vX.Y.Z`, creates
the GitHub Release, and the `publish` job in
`.github/workflows/release-please.yml` publishes to npm with Trusted
Publishing.

## Layout

```
src/
  core/            query engine, auth, http, schemas, mappers
  cli.ts           commander entrypoint
  cli-render.ts    pretty-print helpers
  mcp.ts           MCP server entrypoint
tests/
  unit/            pure function tests against fakes
  integration/     HttpGraphClient via MSW
  fixtures/        Graph response samples
  helpers/         FakeGraphClient
```
