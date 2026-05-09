import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { defaultCachePath, makeFileCachePlugin } from "../../src/core/auth.ts";

describe("defaultCachePath", () => {
  const origXdg = process.env.XDG_CONFIG_HOME;
  const origHome = process.env.HOME;

  afterEach(() => {
    if (origXdg === undefined) delete process.env.XDG_CONFIG_HOME;
    else process.env.XDG_CONFIG_HOME = origXdg;
    if (origHome !== undefined) process.env.HOME = origHome;
  });

  it("uses XDG_CONFIG_HOME when set", () => {
    process.env.XDG_CONFIG_HOME = "/custom/cfg";
    expect(defaultCachePath()).toBe("/custom/cfg/outlook-query/msal-cache.json");
  });

  it("falls back to ~/.config when XDG_CONFIG_HOME unset", () => {
    delete process.env.XDG_CONFIG_HOME;
    process.env.HOME = "/home/test";
    expect(defaultCachePath()).toBe("/home/test/.config/outlook-query/msal-cache.json");
  });
});

describe("makeFileCachePlugin", () => {
  let tmpDir: string;
  let cachePath: string;

  beforeEach(async () => {
    tmpDir = await mkdtemp(join(tmpdir(), "outlook-query-"));
    cachePath = join(tmpDir, "nested", "cache.json");
  });

  afterEach(async () => {
    await rm(tmpDir, { recursive: true, force: true });
  });

  it("beforeCacheAccess no-ops when file missing", async () => {
    const plugin = makeFileCachePlugin(cachePath);
    let deserializeCalled = false;
    const ctx = makeContext({
      deserialize: () => {
        deserializeCalled = true;
      },
    });
    await plugin.beforeCacheAccess(ctx);
    expect(deserializeCalled).toBe(false);
  });

  it("beforeCacheAccess deserializes existing file", async () => {
    await writeFile(cachePath.replace("/nested/cache.json", "/cache-existing.json"), "DATA");
    const plugin = makeFileCachePlugin(
      cachePath.replace("/nested/cache.json", "/cache-existing.json"),
    );
    let deserialized = "";
    const ctx = makeContext({ deserialize: (d) => (deserialized = d) });
    await plugin.beforeCacheAccess(ctx);
    expect(deserialized).toBe("DATA");
  });

  it("afterCacheAccess writes serialized data and creates parent dirs", async () => {
    const plugin = makeFileCachePlugin(cachePath);
    const ctx = makeContext({ serialize: () => "SERIALIZED" });
    await plugin.afterCacheAccess(ctx);

    const written = await readFile(cachePath, "utf-8");
    expect(written).toBe("SERIALIZED");
  });

  it("afterCacheAccess sets 0600 file permissions", async () => {
    const plugin = makeFileCachePlugin(cachePath);
    const ctx = makeContext({ serialize: () => "data" });
    await plugin.afterCacheAccess(ctx);

    const st = await stat(cachePath);
    const mode = st.mode & 0o777;
    expect(mode).toBe(0o600);
  });

  it("afterCacheAccess skips write when cacheHasChanged is false", async () => {
    const plugin = makeFileCachePlugin(cachePath);
    const ctx = makeContext({ serialize: () => "data", cacheHasChanged: false });
    await plugin.afterCacheAccess(ctx);

    await expect(readFile(cachePath)).rejects.toThrow();
  });
});

function makeContext(opts: {
  serialize?: () => string;
  deserialize?: (d: string) => void;
  cacheHasChanged?: boolean;
}) {
  return {
    cacheHasChanged: opts.cacheHasChanged ?? true,
    tokenCache: {
      serialize: opts.serialize ?? (() => ""),
      deserialize: opts.deserialize ?? (() => {}),
    },
  } as unknown as Parameters<
    NonNullable<
      ReturnType<typeof import("../../src/core/auth.ts").makeFileCachePlugin>
    >["beforeCacheAccess"]
  >[0];
}
