# CLAUDE.md

## Stack

TypeScript / Node 22+ / pnpm 11. ESM only (`type: module`). Single project; `pnpm-workspace.yaml` exists only to declare `allowBuilds` for esbuild + msw.

## Source imports use `.ts` extensions

`tsconfig.json` has `rewriteRelativeImportExtensions: true`. Write `import x from "./foo.ts"` — tsc rewrites to `./foo.js` on emit.

## Scripts

- `pnpm fix` — `biome check --write` (lint + format + organize imports). There is **no** `pnpm format` script.
- `pnpm check` — fix + typecheck + vitest. Run after non-trivial edits.
- `pnpm dev:cli`, `pnpm dev:mcp` — tsx runners. ~6s cold-start under WSL filesystem; affects MCP smoke tests.

## tsconfig strictness — gotchas

- `exactOptionalPropertyTypes` — for `x?: T`, never assign `x: undefined`. Conditionally spread instead: `...(cond ? { x: val } : {})`.
- `noUncheckedIndexedAccess` — `array[i]` is `T | undefined`. Guard accesses.
- `verbatimModuleSyntax` — type-only imports must use `import type {...}` (Biome enforces).

## Zod 4 (not 3)

Strict objects: `z.strictObject({...})` (not `.strict()` chain). Inferred types: `z.infer<typeof X>` for output, `z.input<typeof X>` for pre-transform input. `since`/`until` in `SearchParamsSchema` use a `dateInput` transform — input is string, output is `ParsedDateInput` (`{ date: Date, dateOnly: boolean }`). Date-only `until` is inclusive of the whole day — query-builder rolls to next-day-start and uses `lt` (OData) / `<` (KQL).

## GraphClient is a Protocol — inject, don't import the HTTP impl in tests

`src/core/graph-client.ts` defines the `GraphClient` interface. `HttpGraphClient` is the prod impl; tests use `FakeGraphClient` (`tests/helpers/`) for unit logic and MSW for HTTP-layer integration tests. New core functions accept `GraphClient` as a parameter.

## Query builder has two paths — don't combine

`buildGraphQuery` switches on whether any free-text field is present (`query` / `from` / `to` / `subject_contains` / `body_contains`):
- With text → KQL `$search`; **no** `$filter`, **no** `$orderby` (Graph rejects the combo).
- Without text → OData `$filter` + `$orderby receivedDateTime desc`.

`inference_classification` is OData-path only: combining it with text **throws** (KQL has no inference keyword). Other structured filters fold into KQL.

## Auth scope discipline

Scopes are `Mail.Read` + `offline_access` only — keep it that way. The `auth` CLI subcommand acquires a token via `getToken()` directly; it does not call `/me`. Stick to Graph endpoints these two scopes cover (`User.Read` endpoints like `/me` are out).

## Env loading

Both entry points (`cli.ts`, `mcp.ts`) call `process.loadEnvFile()` at startup (Node 22 built-in). `.env` from cwd is loaded silently if present. Required: `OUTLOOK_QUERY_CLIENT_ID`. Optional: `OUTLOOK_QUERY_TENANT_ID` (defaults to `common`).

## Agent workflow

- Issues live in GitHub Issues, via `gh` — conventions in `docs/agents/issue-tracker.md`.
- Triage label vocabulary — `docs/agents/triage-labels.md`.
- Domain docs (ADRs in `docs/adr/`) — how to consume them: `docs/agents/domain.md`.

## Run `pnpm build` after changes that affect the MCP server

The MCP server is launched by Claude Code from `dist/mcp.js`. After changes to anything under `src/core/` or `src/mcp.ts`, run `pnpm build` before ending the turn so the next MCP session loads the updated code. Skip the build for changes that don't affect the MCP runtime (CLI-only files, tests, README, this file).
