// Minimal promise wrapper around IndexedDB — no external dependency.
// One database, opened once; every store/version change lives in schema.js.

export function openDB(name, version, onUpgrade) {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(name, version);
    req.onupgradeneeded = (e) => onUpgrade(req.result, e.oldVersion, e.newVersion);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
    req.onblocked = () => reject(new Error("IndexedDB upgrade blocked by another open tab"));
  });
}

export function tx(db, storeNames, mode, fn) {
  return new Promise((resolve, reject) => {
    const t = db.transaction(storeNames, mode);
    let result;
    t.oncomplete = () => resolve(result);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error || new Error("Transaction aborted"));
    Promise.resolve(fn(t)).then((r) => { result = r; }).catch(reject);
  });
}

export function reqAsPromise(req) {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export function getAll(store) { return reqAsPromise(store.getAll()); }
export function get(store, key) { return reqAsPromise(store.get(key)); }
export function put(store, value) { return reqAsPromise(store.put(value)); }
export function del(store, key) { return reqAsPromise(store.delete(key)); }
export function clearStore(store) { return reqAsPromise(store.clear()); }
