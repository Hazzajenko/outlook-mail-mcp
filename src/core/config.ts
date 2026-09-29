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
  const clientId = env.OUTLOOK_MAIL_CLIENT_ID;
  if (!clientId) {
    throw new ConfigError(
      "OUTLOOK_MAIL_CLIENT_ID is not set. Run `npx -p @hazzajenko/outlook-mail-mcp outlook-mail setup` in a terminal for a guided walkthrough, or set the env var to your Azure app registration's client ID.",
    );
  }
  // Empty string would produce authority ".../" — treat like unset (→ common).
  const tenantId = env.OUTLOOK_MAIL_TENANT_ID;
  return {
    clientId,
    ...(tenantId ? { tenantId } : {}),
  };
}
