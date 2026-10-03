// Data the Admin sets once for EVERYONE (today: the OR / MOR / OR+MOR type
// of each shift). The app has no server, so the shared copy is a JSON file
// on its own branch ("shared-data") of the app's GitHub repository:
//   - every device reads it from raw.githubusercontent.com (public, no login);
//   - the Admin's device writes it through the GitHub API, using a personal
//     access token that is kept only on that device.
// The branch is only data: the site is built from "main", so saving here
// never rebuilds or redeploys the app.

export const SHARED_REPO = { owner: "3omid", repo: "shift-priority-app", branch: "shared-data", baseBranch: "main", path: "shared-data.json" };
export const SHARED_FORMAT = "shift-priority-shared";
const API = `https://api.github.com/repos/${SHARED_REPO.owner}/${SHARED_REPO.repo}`;
export const SHARED_RAW_URL = `https://raw.githubusercontent.com/${SHARED_REPO.owner}/${SHARED_REPO.repo}/${SHARED_REPO.branch}/${SHARED_REPO.path}`;

export function isSharedFile(x) {
  return !!x && x.format === SHARED_FORMAT && typeof x.updatedAt === "number" && !!x.shiftServiceTypes && typeof x.shiftServiceTypes === "object";
}

export function buildSharedFile(shiftServiceTypes, now = Date.now()) {
  return { format: SHARED_FORMAT, version: 1, updatedAt: now, shiftServiceTypes: { ...shiftServiceTypes } };
}

// Should this device take the shared copy?
//   sync = { syncedAt, dirty } kept on this device.
// Everyone follows the shared copy, except an Admin device that has edits
// not shared yet (dirty) — those stay until they're published. An Admin
// device that already had types before sharing existed (nothing shared yet,
// but types set locally) counts as having unshared edits.
export function decideSync({ sync, shared, isAdmin, localCount }) {
  const s = sync || { syncedAt: 0, dirty: false };
  if (isAdmin && s.dirty) return { apply: false, markDirty: false };
  if (!isSharedFile(shared)) {
    return { apply: false, markDirty: !!(isAdmin && s.syncedAt === 0 && localCount > 0) };
  }
  return { apply: shared.updatedAt !== s.syncedAt, markDirty: false };
}

// The shared copy, or null (offline, nothing shared yet, or unreadable).
// The time query string skips GitHub's few-minute cache.
export async function fetchSharedFile(fetchImpl = fetch, now = Date.now()) {
  try {
    const res = await fetchImpl(`${SHARED_RAW_URL}?t=${now}`, { cache: "no-store" });
    if (!res.ok) return null;
    const json = await res.json();
    return isSharedFile(json) ? json : null;
  } catch {
    return null;
  }
}

function toBase64Utf8(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

export class PublishError extends Error {
  constructor(kind, status) { super(kind); this.kind = kind; this.status = status; }
}
function errorFor(status) {
  if (status === 401) return new PublishError("badToken", status);
  if (status === 403 || status === 404) return new PublishError("noAccess", status);
  return new PublishError("other", status);
}

// Saves the shared file. On the very first save it creates the
// "shared-data" branch (from main). Retries once if the file changed in
// between (another Admin device saved at the same moment).
export async function publishSharedFile({ token, file, fetchImpl = fetch }) {
  const { branch, baseBranch, path } = SHARED_REPO;
  const headers = { Authorization: `Bearer ${token}`, Accept: "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28" };
  const call = async (url, opts = {}) => {
    try { return await fetchImpl(url, { cache: "no-store", ...opts, headers: { ...headers, ...(opts.body ? { "Content-Type": "application/json" } : {}) } }); }
    catch { throw new PublishError("network"); }
  };

  // 1) Make sure the data branch exists.
  let res = await call(`${API}/git/ref/heads/${branch}`);
  if (res.status === 404) {
    const base = await call(`${API}/git/ref/heads/${baseBranch}`);
    if (!base.ok) throw errorFor(base.status);
    const baseSha = (await base.json()).object.sha;
    const created = await call(`${API}/git/refs`, { method: "POST", body: JSON.stringify({ ref: `refs/heads/${branch}`, sha: baseSha }) });
    if (!created.ok && created.status !== 422) throw errorFor(created.status); // 422 = created meanwhile
  } else if (!res.ok) {
    throw errorFor(res.status);
  }

  // 2) Write the file (with its current sha when it already exists).
  const body = JSON.stringify(file, null, 2) + "\n";
  for (let attempt = 0; attempt < 2; attempt++) {
    res = await call(`${API}/contents/${path}?ref=${branch}`);
    let sha;
    if (res.ok) sha = (await res.json()).sha;
    else if (res.status !== 404) throw errorFor(res.status);
    res = await call(`${API}/contents/${path}`, {
      method: "PUT",
      body: JSON.stringify({ message: "Update shared shift types (from Admin)", content: toBase64Utf8(body), branch, ...(sha ? { sha } : {}) }),
    });
    if (res.ok) return { ok: true };
    if ((res.status === 409 || res.status === 422) && attempt === 0) continue;
    throw errorFor(res.status);
  }
  throw new PublishError("other");
}
