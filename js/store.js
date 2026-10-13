/*
 * Lokale opslag op het toestel (IndexedDB). Er gaat niets naar een server:
 * het register en het originele Excel-bestand blijven op de telefoon.
 */
(function (root) {
  'use strict';
  const DB = 'actieregister';
  const STORE = 'kv';
  let dbPromise;

  function open() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        const req = indexedDB.open(DB, 1);
        req.onupgradeneeded = () => req.result.createObjectStore(STORE);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    return dbPromise;
  }

  function tx(mode, fn) {
    return open().then((db) => new Promise((resolve, reject) => {
      const t = db.transaction(STORE, mode);
      const req = fn(t.objectStore(STORE));
      t.oncomplete = () => resolve(req && req.result);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    }));
  }

  root.Store = {
    get: (key) => tx('readonly', (s) => s.get(key)),
    set: (key, value) => tx('readwrite', (s) => s.put(value, key)),
    clear: () => tx('readwrite', (s) => s.clear()),
    pref(key, value) {
      const k = 'actieregister.' + key;
      try {
        if (value === undefined) return JSON.parse(localStorage.getItem(k));
        localStorage.setItem(k, JSON.stringify(value));
      } catch (e) { return null; }
      return value;
    },
  };
})(self);
