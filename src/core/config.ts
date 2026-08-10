/**
 * Resolution of auth configuration from the environment, shared by every
 * entry point (CLI, MCP server, smoke script).
 */

export class ConfigError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ConfigError";
  }
}

export interface ResolvedAuthConfig {
  clientId: string;
  tenantId?: string;
}

export function resolveAuthConfig(
  env: Record<string, string | undefined> = process.env,
): ResolvedAuthConfig {
  const clientId = env.OUTLOOK_QUERY_CLIENT_ID;
  if (!clientId) {
    throw new ConfigError(
      "OUTLOOK_QUERY_CLIENT_ID is not set. Run `npx outlook-query setup` in a terminal for a guided walkthrough, or set the env var to your Azure app registration's client ID.",
    );
  }
  // Empty string would produce authority ".../" — treat like unset (→ common).
  const tenantId = env.OUTLOOK_QUERY_TENANT_ID;
  return {
    clientId,
    ...(tenantId ? { tenantId } : {}),
  };
}
