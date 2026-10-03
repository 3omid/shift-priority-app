import { describe, it, expect } from "vitest";
import { decideSync, buildSharedFile, isSharedFile, fetchSharedFile, publishSharedFile, SHARED_REPO, SHARED_RAW_URL } from "../src/sharedData.js";

const shared = (updatedAt, map = { "21@NMK": "MOR" }) => buildSharedFile(map, updatedAt);

describe("shared data: who follows the shared copy", () => {
  it("a regular device always takes a newer shared copy, even with old local types", () => {
    expect(decideSync({ sync: { syncedAt: 0, dirty: false }, shared: shared(100), isAdmin: false, localCount: 5 })).toEqual({ apply: true, markDirty: false });
    expect(decideSync({ sync: { syncedAt: 100, dirty: true }, shared: shared(200), isAdmin: false, localCount: 5 }).apply).toBe(true);
    expect(decideSync({ sync: { syncedAt: 100, dirty: false }, shared: shared(100), isAdmin: false, localCount: 5 }).apply).toBe(false);
  });

  it("an Admin device keeps its unshared edits, and otherwise follows the shared copy", () => {
    expect(decideSync({ sync: { syncedAt: 100, dirty: true }, shared: shared(200), isAdmin: true, localCount: 3 }).apply).toBe(false);
    expect(decideSync({ sync: { syncedAt: 100, dirty: false }, shared: shared(200), isAdmin: true, localCount: 3 }).apply).toBe(true);
  });

  it("an Admin device with types set before sharing existed marks them as unshared", () => {
    expect(decideSync({ sync: { syncedAt: 0, dirty: false }, shared: null, isAdmin: true, localCount: 4 })).toEqual({ apply: false, markDirty: true });
    expect(decideSync({ sync: { syncedAt: 0, dirty: false }, shared: null, isAdmin: true, localCount: 0 })).toEqual({ apply: false, markDirty: false });
    expect(decideSync({ sync: { syncedAt: 0, dirty: false }, shared: null, isAdmin: false, localCount: 4 })).toEqual({ apply: false, markDirty: false });
  });

  it("ignores a malformed shared file", () => {
    expect(isSharedFile({ foo: 1 })).toBe(false);
    expect(decideSync({ sync: null, shared: { foo: 1 }, isAdmin: false, localCount: 0 }).apply).toBe(false);
  });
});

function fakeFetch(routes) {
  const calls = [];
  const impl = async (url, opts = {}) => {
    calls.push({ url, method: opts.method || "GET", body: opts.body ? JSON.parse(opts.body) : null, headers: opts.headers });
    const r = routes.shift();
    if (r instanceof Error) throw r;
    return { ok: r.status >= 200 && r.status < 300, status: r.status, json: async () => r.json };
  };
  return { impl, calls };
}
const API = `https://api.github.com/repos/${SHARED_REPO.owner}/${SHARED_REPO.repo}`;

describe("shared data: download", () => {
  it("reads the data branch's raw file, skipping the cache", async () => {
    const good = shared(5);
    const f = fakeFetch([{ status: 200, json: good }]);
    expect(await fetchSharedFile(f.impl, 123)).toEqual(good);
    expect(f.calls[0].url).toBe(`${SHARED_RAW_URL}?t=123`);
    expect(SHARED_RAW_URL).toContain("/shared-data/");
  });
  it("returns null when nothing is shared yet or the phone is offline", async () => {
    expect(await fetchSharedFile(fakeFetch([{ status: 404 }]).impl)).toBeNull();
    expect(await fetchSharedFile(fakeFetch([new Error("offline")]).impl)).toBeNull();
  });
});

describe("shared data: publish (never touches main)", () => {
  it("first save: creates the data branch from main, then the file", async () => {
    const f = fakeFetch([
      { status: 404 }, // data branch missing
      { status: 200, json: { object: { sha: "mainsha" } } }, // main
      { status: 201, json: {} }, // branch created
      { status: 404 }, // file missing
      { status: 201, json: {} }, // file created
    ]);
    const file = shared(7, { "21@NMK": "MOR", "ران@*": "OR" });
    await publishSharedFile({ token: "tok", file, fetchImpl: f.impl });
    expect(f.calls[2]).toMatchObject({ url: `${API}/git/refs`, method: "POST", body: { ref: "refs/heads/shared-data", sha: "mainsha" } });
    const put = f.calls[4];
    expect(put).toMatchObject({ url: `${API}/contents/shared-data.json`, method: "PUT" });
    expect(put.body.branch).toBe("shared-data");
    expect(put.body.sha).toBeUndefined();
    expect(put.headers.Authorization).toBe("Bearer tok");
    const decoded = new TextDecoder().decode(Uint8Array.from(atob(put.body.content), (c) => c.charCodeAt(0)));
    expect(JSON.parse(decoded)).toEqual(file);
    // Nothing is ever written to main.
    expect(f.calls.filter((c) => c.method !== "GET").every((c) => !JSON.stringify(c.body).includes('"main"'))).toBe(true);
  });

  it("later saves update the file with its sha, retrying once on a conflict", async () => {
    const f = fakeFetch([
      { status: 200, json: {} }, // branch exists
      { status: 200, json: { sha: "old" } }, { status: 409 },
      { status: 200, json: { sha: "newer" } }, { status: 200, json: {} },
    ]);
    await publishSharedFile({ token: "t", file: shared(1), fetchImpl: f.impl });
    expect(f.calls[2].body.sha).toBe("old");
    expect(f.calls[4].body.sha).toBe("newer");
  });

  it("reports a wrong token, missing access and no network", async () => {
    await expect(publishSharedFile({ token: "t", file: shared(1), fetchImpl: fakeFetch([{ status: 401 }]).impl })).rejects.toMatchObject({ kind: "badToken" });
    await expect(publishSharedFile({ token: "t", file: shared(1), fetchImpl: fakeFetch([{ status: 403 }]).impl })).rejects.toMatchObject({ kind: "noAccess" });
    await expect(publishSharedFile({ token: "t", file: shared(1), fetchImpl: fakeFetch([new Error("x")]).impl })).rejects.toMatchObject({ kind: "network" });
  });
});
