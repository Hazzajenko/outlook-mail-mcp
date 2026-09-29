import { describe, expect, it } from "vitest";
import { ConfigError, resolveAuthConfig } from "../../src/core/config.ts";

describe("resolveAuthConfig", () => {
  it("throws a ConfigError naming the setup command when client ID is missing", () => {
    expect(() => resolveAuthConfig({})).toThrowError(ConfigError);
    expect(() => resolveAuthConfig({})).toThrowError(/OUTLOOK_MAIL_CLIENT_ID/);
    expect(() => resolveAuthConfig({})).toThrowError(/outlook-mail setup/);
  });

  it("treats an empty client ID as missing", () => {
    expect(() => resolveAuthConfig({ OUTLOOK_MAIL_CLIENT_ID: "" })).toThrowError(ConfigError);
  });

  it("resolves client ID with no tenant when tenant is unset", () => {
    const config = resolveAuthConfig({ OUTLOOK_MAIL_CLIENT_ID: "client-123" });
    expect(config).toEqual({ clientId: "client-123" });
    expect("tenantId" in config).toBe(false);
  });

  it("treats an empty tenant ID as unset", () => {
    const config = resolveAuthConfig({
      OUTLOOK_MAIL_CLIENT_ID: "client-123",
      OUTLOOK_MAIL_TENANT_ID: "",
    });
    expect("tenantId" in config).toBe(false);
  });

  it("resolves tenant ID when set", () => {
    const config = resolveAuthConfig({
      OUTLOOK_MAIL_CLIENT_ID: "client-123",
      OUTLOOK_MAIL_TENANT_ID: "consumers",
    });
    expect(config).toEqual({ clientId: "client-123", tenantId: "consumers" });
  });
});
