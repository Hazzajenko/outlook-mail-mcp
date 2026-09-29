import { chmod, mkdir, readFile, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { type Configuration, type ICachePlugin, PublicClientApplication } from "@azure/msal-node";

const SCOPES = ["Mail.Read", "offline_access"];

export interface AuthConfig {
  clientId: string;
  tenantId?: string;
  cachePath?: string;
  /**
   * When false, never start the device-code flow: a missing or unrefreshable
   * cached token throws NotAuthenticatedError instead. The MCP server uses
   * this — a device-code prompt inside a tool call goes to stderr, which the
   * user never sees. Defaults to true (CLI behavior).
   */
  interactive?: boolean;
}

export class NotAuthenticatedError extends Error {
  constructor() {
    super(
      "Not authenticated. Run `npx -p @hazzajenko/outlook-mail-mcp outlook-mail auth` in a terminal, then retry.",
    );
    this.name = "NotAuthenticatedError";
  }
}

export interface TokenProvider {
  getToken(): Promise<string>;
}

export function defaultCachePath(): string {
  const xdg = process.env.XDG_CONFIG_HOME ?? join(homedir(), ".config");
  return join(xdg, "outlook-mail", "msal-cache.json");
}

export function makeFileCachePlugin(path: string): ICachePlugin {
  return {
    async beforeCacheAccess(ctx) {
      try {
        const data = await readFile(path, "utf-8");
        ctx.tokenCache.deserialize(data);
      } catch (e: unknown) {
        if ((e as NodeJS.ErrnoException).code !== "ENOENT") throw e;
      }
    },
    async afterCacheAccess(ctx) {
      if (!ctx.cacheHasChanged) return;
      await mkdir(dirname(path), { recursive: true });
      await writeFile(path, ctx.tokenCache.serialize(), { mode: 0o600 });
      await chmod(path, 0o600);
    },
  };
}

export function createTokenProvider(config: AuthConfig): TokenProvider {
  const cachePath = config.cachePath ?? defaultCachePath();
  const msalConfig: Configuration = {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId ?? "common"}`,
    },
    cache: { cachePlugin: makeFileCachePlugin(cachePath) },
  };
  const pca = new PublicClientApplication(msalConfig);

  return {
    async getToken() {
      const accounts = await pca.getTokenCache().getAllAccounts();
      const account = accounts[0];
      if (account) {
        try {
          const result = await pca.acquireTokenSilent({ account, scopes: SCOPES });
          if (result?.accessToken) return result.accessToken;
        } catch {
          // silent failed; fall through to interactive
        }
      }

      if (config.interactive === false) {
        throw new NotAuthenticatedError();
      }

      const result = await pca.acquireTokenByDeviceCode({
        scopes: SCOPES,
        deviceCodeCallback: (response) => {
          process.stderr.write(`${response.message}\n`);
        },
      });
      if (!result?.accessToken) {
        throw new Error("Failed to acquire access token via device code flow");
      }
      return result.accessToken;
    },
  };
}
