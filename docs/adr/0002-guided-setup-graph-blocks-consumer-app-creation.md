# Guided setup: Graph blocks consumer-identity app creation

Status: accepted

We wanted `outlook-query setup` to create the Entra app registration
automatically (ADR-0001). A live spike (issue #2, results on issue #1) proved
this impossible for the target audience: Microsoft Graph refuses
`POST /applications` for personal Microsoft accounts with
`403 "Client <id> is not authorized to create apps using consumer identity"`,
regardless of which client requests it — tested with the Graph Command Line
Tools first-party client, the Azure CLI first-party client (blocked earlier,
at consent), and a consumer-audience third-party app. Consent of
`Application.ReadWrite.All` succeeds; the create is refused server-side. The
Graph docs' "Delegated (personal Microsoft account): supported" row for this
API did not match observed behaviour in August 2026. The Entra portal is the
only registration path for consumer accounts.

Decision: `outlook-query setup` is a **guided, verified** flow instead. It
walks the user through the portal steps, accepts the resulting client ID,
writes `.env` with a tenant value derived from the chosen audience (making the
`AADSTS9002331` mismatch unrepresentable), then verifies the registration by
running the normal device-code auth and translating known failures
(`AADSTS9002331`, public-client flows disabled, app not found, consent
declined) into "go back and fix step N" guidance.

## Consequences

- No elevated Graph scopes anywhere: the first-party bootstrap client, the
  in-memory `Application.ReadWrite.All` token, and the setup-only get/post/patch
  Graph protocol from ADR-0001 are all unnecessary. The codebase keeps a single
  GET-only Graph protocol and the `Mail.Read` + `offline_access` scope
  discipline holds everywhere, with no exceptions to explain.
- Successful verification caches a token, so setup doubles as first-run auth.
- If Microsoft ever re-opens consumer app creation (the docs suggest it once
  worked), ADR-0001's design is the recorded starting point.
