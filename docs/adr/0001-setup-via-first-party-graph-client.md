# Setup command authenticates through a Microsoft first-party client

Status: superseded by ADR-0002

The `outlook-mail setup` subcommand automates the Entra app registration that
users would otherwise do by hand in the portal. To call Graph
(`POST /applications`) before our own app exists, it must authenticate as an
already-registered client — a chicken-and-egg problem. We resolve it by doing a
device-code sign-in through the **Microsoft Graph Command Line Tools**
first-party public client (`14d82eec-204b-4c2f-b7e8-296a70dab67e`), requesting
the delegated scope `Application.ReadWrite.All` for the setup session only.

## Considered options

- **Azure CLI (`az ad app create`)** — rejected: ~1 GB toolchain the target
  user (a personal Outlook account holder) does not have, weak support for
  tenant-less personal Microsoft accounts, and its first-party client is being
  progressively restricted from Graph scopes.
- **Microsoft Graph PowerShell (`New-MgApplication`)** — rejected: requires
  `Install-Module` (slow, prompts) and pwsh is not a given outside Windows.
- **In-repo Node script via a first-party client** — chosen: zero new
  dependencies (users already need Node 22), same `@azure/msal-node` device-code
  flow the project ships, cross-platform, works from `npx` because it ships in
  `dist` as a CLI subcommand.

## Consequences

- The consent screen shows "Microsoft Graph Command Line Tools", not
  "outlook-mail". The README warns about this so users are not spooked.
- The `Application.ReadWrite.All` token is held **in memory only** and is never
  written to the msal cache file. The on-disk cache holds only the
  `Mail.Read` + `offline_access` token from our own app, so the scope
  discipline in CLAUDE.md ("this tool only ever has Mail.Read") stays true for
  everything persisted.
- Our own app registration is unchanged: `Mail.Read` + `offline_access` only.
  The elevated scope exists only on the transient first-party session.
- We depend on Microsoft continuing to allow dynamic consent of
  `Application.ReadWrite.All` on that first-party client. If that changes, or a
  tenant policy blocks app registration, the documented manual portal steps in
  the README remain the fallback and setup error messages point there.
