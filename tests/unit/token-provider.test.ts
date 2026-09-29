import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  getAllAccounts: vi.fn(),
  acquireTokenSilent: vi.fn(),
  acquireTokenByDeviceCode: vi.fn(),
}));

vi.mock("@azure/msal-node", () => ({
  PublicClientApplication: class {
    getTokenCache() {
      return { getAllAccounts: mocks.getAllAccounts };
    }
    acquireTokenSilent = mocks.acquireTokenSilent;
    acquireTokenByDeviceCode = mocks.acquireTokenByDeviceCode;
  },
}));

import { createTokenProvider, NotAuthenticatedError } from "../../src/core/auth.ts";

const ACCOUNT = { homeAccountId: "acct" };

beforeEach(() => {
  vi.resetAllMocks();
});

describe("createTokenProvider (non-interactive)", () => {
  const provider = () => createTokenProvider({ clientId: "c", interactive: false });

  it("rejects with NotAuthenticatedError when no account is cached", async () => {
    mocks.getAllAccounts.mockResolvedValue([]);
    await expect(provider().getToken()).rejects.toThrowError(NotAuthenticatedError);
    await expect(provider().getToken()).rejects.toThrowError(/outlook-mail auth/);
    expect(mocks.acquireTokenByDeviceCode).not.toHaveBeenCalled();
  });

  it("rejects with NotAuthenticatedError when silent refresh fails", async () => {
    mocks.getAllAccounts.mockResolvedValue([ACCOUNT]);
    mocks.acquireTokenSilent.mockRejectedValue(new Error("refresh token expired"));
    await expect(provider().getToken()).rejects.toThrowError(NotAuthenticatedError);
    expect(mocks.acquireTokenByDeviceCode).not.toHaveBeenCalled();
  });

  it("returns the token when silent refresh succeeds", async () => {
    mocks.getAllAccounts.mockResolvedValue([ACCOUNT]);
    mocks.acquireTokenSilent.mockResolvedValue({ accessToken: "tok-1" });
    await expect(provider().getToken()).resolves.toBe("tok-1");
    expect(mocks.acquireTokenByDeviceCode).not.toHaveBeenCalled();
  });
});

describe("createTokenProvider (interactive, default)", () => {
  it("falls through to the device-code flow when no account is cached", async () => {
    mocks.getAllAccounts.mockResolvedValue([]);
    mocks.acquireTokenByDeviceCode.mockResolvedValue({ accessToken: "tok-2" });
    const provider = createTokenProvider({ clientId: "c" });
    await expect(provider.getToken()).resolves.toBe("tok-2");
    expect(mocks.acquireTokenByDeviceCode).toHaveBeenCalledOnce();
  });

  it("falls through to the device-code flow when silent refresh fails", async () => {
    mocks.getAllAccounts.mockResolvedValue([ACCOUNT]);
    mocks.acquireTokenSilent.mockRejectedValue(new Error("expired"));
    mocks.acquireTokenByDeviceCode.mockResolvedValue({ accessToken: "tok-3" });
    const provider = createTokenProvider({ clientId: "c" });
    await expect(provider.getToken()).resolves.toBe("tok-3");
    expect(mocks.acquireTokenByDeviceCode).toHaveBeenCalledOnce();
  });
});
