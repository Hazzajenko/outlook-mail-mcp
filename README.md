# outlook-query

Query personal Outlook via Microsoft Graph. CLI + MCP server, sharing one core. Read-only.

## Stack

TypeScript / Node 22+ / pnpm. Auth via `@azure/msal-node` (device code flow, public client, no secret). Graph access via hand-rolled `fetch` wrapper with Zod-validated responses. Tests: Vitest + MSW.

## Setup

### 1. Install

```
pnpm install
pnpm build
```

### 2. Register an Azure app

1. <https://entra.microsoft.com> → Identity → Applications → App registrations → **New registration**.
2. Name: anything (e.g. `outlook-query`).
3. Supported account types: **Personal Microsoft accounts only** (or *Accounts in any org + personal* if you have both).
4. Redirect URI: leave blank (device code flow needs none).
5. Create.
6. On the new app's **Overview** page, copy the **Application (client) ID**.
7. **Authentication** → **Allow public client flows** → **Yes** → Save.
8. **API permissions** → Add → Microsoft Graph → Delegated permissions → check `Mail.Read` and `offline_access` → Add. Click **Grant admin consent** (if available) or accept consent on first run.

### 3. Configure env

Create `.env` in the project root (gitignored), or export in your shell:

```
OUTLOOK_QUERY_CLIENT_ID=<paste app id>
# Optional. Default 'common' (personal + work). Use 'consumers' to lock to personal.
OUTLOOK_QUERY_TENANT_ID=common
```

The CLI auto-loads `.env` from cwd via Node's built-in `process.loadEnvFile()`. The MCP server does too, but Claude Code launches it from its own cwd — set env via `.mcp.json` `env` block instead (see below).

### 4. First-run auth

```
pnpm exec outlook-query auth
```

Opens a device-code message in stderr — visit the URL, enter the code, sign in. Token cached at `~/.config/outlook-query/msal-cache.json` (mode 0600). Subsequent runs are silent.

## CLI

```
outlook-query search --from goldman.com --since -30d
outlook-query search -q "interview" --unread --top 100
outlook-query search --folder Jobs --has-attachment --json
outlook-query get <message-id>
outlook-query folders
```

Flags: `-q/--query`, `--from`, `--to`, `--subject`, `--body`, `--since`, `--until`, `--has-attachment`, `--unread` / `--read`, `--folder`, `--importance`, `--inference-classification` (`focused`/`other`), `--top`, `--json`. Dates: ISO (`2026-05-01`) or relative (`-7d`, `-2w`, `-3h`, `-30m`).

`--inference-classification` needs a `--since`/`--until` alongside it: results are sorted by `receivedDateTime`, and Graph rejects an unbounded inference filter combined with that sort (400 `InefficientFilter`). Same applies to the `list_emails_brief` MCP tool; `count_emails` doesn't sort and is exempt. Omitting the bound raises an error that says so — the raw Graph 400 names neither the cause nor the fix.

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
Graph returning a shape our schemas reject — e.g. a `$select` that stops
requesting a field a Zod schema still marks required, which takes a tool down
completely while the suite stays green.

`pnpm smoke` calls every exposed operation against the real mailbox and parses
each response through its declared result schema (`SearchResultSchema`,
`FullMessageSchema`, …). Requires `OUTLOOK_QUERY_CLIENT_ID` and a cached token;
read-only; exits non-zero on any failure. Cases that need data the mailbox
doesn't have report `SKIP` rather than passing silently.

Known Graph limitation it encodes: `inference_classification` combined with
`$orderby receivedDateTime desc` returns **400 InefficientFilter** unless the
filter also bounds `receivedDateTime` — so pass `--since`/`--until` alongside
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
