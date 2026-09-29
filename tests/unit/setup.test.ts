import { describe, expect, it } from "vitest";
import {
  buildEnvUpdates,
  buildSetupSteps,
  isValidClientId,
  translateVerifyError,
} from "../../src/core/setup.ts";

describe("isValidClientId", () => {
  it("accepts a GUID", () => {
    expect(isValidClientId("81d9df6a-510c-4990-8876-a9c66899f3d7")).toBe(true);
  });
  it("accepts uppercase GUIDs", () => {
    expect(isValidClientId("81D9DF6A-510C-4990-8876-A9C66899F3D7")).toBe(true);
  });
  it("rejects non-GUID input", () => {
    expect(isValidClientId("not-a-guid")).toBe(false);
    expect(isValidClientId("")).toBe(false);
    expect(isValidClientId("81d9df6a510c49908876a9c66899f3d7")).toBe(false);
  });
});

describe("buildEnvUpdates", () => {
  it("writes the client ID and deletes any stale tenant by default", () => {
    expect(buildEnvUpdates("abc", false)).toEqual({
      OUTLOOK_MAIL_CLIENT_ID: "abc",
      OUTLOOK_MAIL_TENANT_ID: null,
    });
  });
  it("derives consumers tenant in personal-only mode", () => {
    expect(buildEnvUpdates("abc", true)).toEqual({
      OUTLOOK_MAIL_CLIENT_ID: "abc",
      OUTLOOK_MAIL_TENANT_ID: "consumers",
    });
  });
});

describe("buildSetupSteps", () => {
  it("names the dual audience by default", () => {
    const steps = buildSetupSteps(false);
    expect(steps).toContain("Accounts in any organizational directory");
    expect(steps).not.toMatch(/Personal Microsoft accounts only(?!.*personal\))/);
  });
  it("names personal-only audience with --personal-only", () => {
    expect(buildSetupSteps(true)).toContain("Personal Microsoft accounts only");
  });
  it("numbers the public-client and permissions steps", () => {
    const steps = buildSetupSteps(false);
    expect(steps).toMatch(/6\..*public client flows/i);
    expect(steps).toMatch(/7\..*Mail\.Read/);
  });
});

describe("translateVerifyError", () => {
  it("maps AADSTS9002331 to the audience/tenant fix", () => {
    const g = translateVerifyError(
      "AADSTS9002331: Application is configured for use by Microsoft Account users only.",
      false,
    );
    expect(g).toContain("--personal-only");
    expect(g).toContain("step 3");
  });
  it("maps AADSTS7000218 to the public client flows step", () => {
    const g = translateVerifyError(
      "AADSTS7000218: request body must contain client_assertion",
      false,
    );
    expect(g).toContain("step 6");
  });
  it("maps AADSTS700016 to wrong client ID guidance", () => {
    const g = translateVerifyError("AADSTS700016: Application not found in the directory", false);
    expect(g).toContain("step 5");
  });
  it("maps AADSTS65004 to declined consent", () => {
    const g = translateVerifyError("AADSTS65004: User declined to consent", false);
    expect(g.toLowerCase()).toContain("consent");
  });
  it("falls back to the manual section for unknown errors", () => {
    const g = translateVerifyError("something exploded", false);
    expect(g).toContain("Manual setup");
  });
  it("suggests removing --personal-only when mismatch is the other way", () => {
    const g = translateVerifyError(
      "AADSTS50194: Application is not configured as a multi-tenant application",
      true,
    );
    expect(g).toContain("without --personal-only");
  });
});
