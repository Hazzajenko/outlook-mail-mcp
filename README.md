# outlook-query

Query personal Outlook via Microsoft Graph. CLI + MCP server, sharing one core. Read-only.

## Stack

TypeScript / Node 22+ / pnpm. Auth via `@azure/msal-node` with device code flow, a public client, and no secret. Graph access via hand-rolled `fetch` wrapper with Zod-validated responses. Tests: Vitest + MSW.

## Setup

### 1. Install

```
pnpm install
pnpm build
```

### 2. Run guided setup

```
pnpm exec outlook-query setup
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
  of the walkthrough. It writes `OUTLOOK_QUERY_TENANT_ID=consumers` to match. The
  tenant value is always derived from your audience choice, so the
  `AADSTS9002331` mismatch cannot happen on this path.
- `--client-id <id>`: skip the prompt, for non-interactive or agent use.
- `--no-verify`: skip the verification sign-in. Run `outlook-query auth` later.

Env can also live in your shell or an MCP `env` block instead of `.env`. The
values are always printed for copy-paste. The CLI auto-loads `.env` from cwd
via Node's built-in `process.loadEnvFile()`. The MCP server does too, but
Claude Code launches it from its own cwd. Set env via the `.mcp.json` `env`
block instead, as shown in the MCP section below.

### Manual setup

Use this if the guided flow fails. The step numbers below match the walkthrough and its error messages.

1. <https://entra.microsoft.com> → Identity → Applications → App registrations → **New registration**.
2. Name: anything, for example `outlook-query`.
3. Supported account types: **Accounts in any org + personal** (default), or **Personal Microsoft accounts only**. With the second, use `--personal-only` or `OUTLOOK_QUERY_TENANT_ID=consumers`.
4. Redirect URI: leave blank. Device code flow needs none. Register.
5. On the new app's **Overview** page, copy the **Application (client) ID**.
6. **Authentication** → **Allow public client flows** → **Yes** → Save.
7. **API permissions** → Add → Microsoft Graph → Delegated permissions → check `Mail.Read` and `offline_access` → Add. Personal accounts have no admin consent. You accept the consent prompt on first sign-in. Work or school tenants can click **Grant admin consent**.

Then create `.env` in the project root, or export the values in your shell. Git ignores `.env`. Example:

```
OUTLOOK_QUERY_CLIENT_ID=<paste app id>
# Optional. Default 'common' (personal + work). Use 'consumers' to lock to personal.
OUTLOOK_QUERY_TENANT_ID=common
```

**If you chose "Personal Microsoft accounts only" in step 3, you must set
`OUTLOOK_QUERY_TENANT_ID=consumers`.** Such apps reject the default `common`
authority with `AADSTS9002331: Application is configured for use by Microsoft
Account users only. Please use the /consumers endpoint`.

Finally authenticate:

```
pnpm exec outlook-query auth
```

Opens a device-code message in stderr. Visit the URL, enter the code, sign in. Token cached at `~/.config/outlook-query/msal-cache.json` (mode 0600). Subsequent runs are silent.

## CLI

```
outlook-query search --from example.com --since -30d
outlook-query search -q "interview" --unread --top 100
outlook-query search --folder Jobs --has-attachment --json
outlook-query get <message-id>
outlook-query folders
```

Flags: `-q/--query`, `--from`, `--to`, `--subject`, `--body`, `--since`, `--until`, `--has-attachment`, `--unread` / `--read`, `--folder`, `--importance`, `--inference-classification` (`focused`/`other`), `--top`, `--json`. Dates: ISO (`2026-05-01`) or relative (`-7d`, `-2w`, `-3h`, `-30m`).

`--inference-classification` needs a `--since`/`--until` alongside it. Results are sorted by `receivedDateTime`, and Graph rejects an unbounded inference filter combined with that sort (400 `InefficientFilter`). Same applies to the `list_emails_brief` MCP tool. `count_emails` doesn't sort and is exempt. Omitting the bound raises an error that says so. The raw Graph 400 names neither the cause nor the fix.

## MCP (Claude Code)

After `pnpm build`, register in your `.mcp.json` or `~/.claude/mcp_servers.json`:

```json
{
  "mcpServers": {
    "outlook-query": {
      "command": "node",
      "args": ["/abs/path/to/outlook-query/dist/mcp.js"],
      "env": {
        "OUTLOOK_QUERY_CLIENT_ID": "<your-client-id>"
      }
    }
  }
}
```

Tools exposed: `search_emails`, `get_email`, `get_conversation`, `count_emails`, `list_emails_brief`, `list_folders`. Same schema as the CLI for the search filter.

## Troubleshooting

**`AADSTS9002331` during auth.** Your app registration is "Personal Microsoft
accounts only" but the tenant is `common`. Set `OUTLOOK_QUERY_TENANT_ID=consumers`,
or re-run `outlook-query setup --personal-only`, which derives it for you.

**MCP tools don't show up in the client.** Fully quit and relaunch the host,
then check the host's MCP logs. Claude Desktop does not hot-reload config. To
smoke-test the server directly, run:

```
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"x","version":"0"}}}' | node dist/mcp.js
```

A JSON response should print within a few seconds.

**`OUTLOOK_QUERY_CLIENT_ID env var not set` at tool-call time.** The host's
working directory doesn't contain your `.env`. Pass the variable via the MCP
config's `env` block, or `cd` into the project dir in the launch command.

**Token expired or 401 from Graph.** Re-run `pnpm exec outlook-query auth`.
Silent refresh usually works. If it fails, you get a device-code prompt.

## Development

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
`FullMessageSchema`, …). It needs `OUTLOOK_QUERY_CLIENT_ID` and a cached token.
It is read-only and exits non-zero on any failure. Cases that need data the mailbox
doesn't have report `SKIP` rather than passing silently.

Known Graph limitation it encodes: `inference_classification` combined with
`$orderby receivedDateTime desc` returns **400 InefficientFilter** unless the
filter also bounds `receivedDateTime`. Pass `--since`/`--until` alongside
`--inference-classification`.

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
