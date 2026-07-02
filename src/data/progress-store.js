import { DB_NAME, DB_VERSION } from "../config.js";

function result(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export async function openProgressStore(factory = indexedDB) {
  const request = factory.open(DB_NAME, DB_VERSION);
  request.onupgradeneeded = () => {
    if (!request.result.objectStoreNames.contains("progress")) {
      request.result.createObjectStore("progress", { keyPath: "questionId" });
    }
  };
  const database = await result(request);
  const store = mode => database.transaction("progress", mode).objectStore("progress");
  return {
    all: () => result(store("readonly").getAll()),
    get: id => result(store("readonly").get(id)),
    put: record => result(store("readwrite").put(record)),
    clear: () => result(store("readwrite").clear()),
    replaceAll: records => new Promise((resolve, reject) => {
      const transaction = database.transaction("progress", "readwrite");
      const target = transaction.objectStore("progress");
      target.clear();
      records.forEach(record => target.put(record));
      transaction.oncomplete = resolve;
      transaction.onerror = () => reject(transaction.error);
      transaction.onabort = () => reject(transaction.error);
    })
  };
}

export function createMemoryStore() {
  const values = new Map();
  return {
    all: async () => [...values.values()],
    get: async id => values.get(id),
    put: async record => values.set(record.questionId, record),
    clear: async () => values.clear(),
    replaceAll: async records => { values.clear(); records.forEach(record => values.set(record.questionId, record)); }
  };
}
