/**
 * Base de datos local IndexedDB: MineriaProduccionDB
 * Tablas:
 *  - registros: 'id, codigo_equipo, fecha, turno, estado, creado_en'
 *  - equipos:   'codigo, ultimo_uso, creado_en'
 *  - areas:     'nombre, ultimo_uso, creado_en'
 */

const DB_NAME = 'MineriaProduccionDB';
const DB_VERSION = 1;

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event) => {
      const db = event.target.result;

      if (!db.objectStoreNames.contains('registros')) {
        const regStore = db.createObjectStore('registros', { keyPath: 'id' });
        regStore.createIndex('codigo_equipo', 'codigo_equipo', { unique: false });
        regStore.createIndex('fecha', 'fecha', { unique: false });
        regStore.createIndex('turno', 'turno', { unique: false });
        regStore.createIndex('estado', 'estado', { unique: false });
        regStore.createIndex('creado_en', 'creado_en', { unique: false });
      }

      if (!db.objectStoreNames.contains('equipos')) {
        const eqStore = db.createObjectStore('equipos', { keyPath: 'codigo' });
        eqStore.createIndex('ultimo_uso', 'ultimo_uso', { unique: false });
      }

      if (!db.objectStoreNames.contains('areas')) {
        const areaStore = db.createObjectStore('areas', { keyPath: 'nombre' });
        areaStore.createIndex('ultimo_uso', 'ultimo_uso', { unique: false });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function createTableHelper(storeName) {
  return {
    async toArray() {
      const db = await openDatabase();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readonly');
        const store = tx.objectStore(storeName);
        const req = store.getAll();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => reject(req.error);
      });
    },

    async get(key) {
      const db = await openDatabase();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readonly');
        const store = tx.objectStore(storeName);
        const req = store.get(key);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      });
    },

    async put(item) {
      const db = await openDatabase();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readwrite');
        const store = tx.objectStore(storeName);
        const req = store.put(item);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    },

    async delete(key) {
      const db = await openDatabase();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readwrite');
        const store = tx.objectStore(storeName);
        const req = store.delete(key);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
      });
    },

    async count() {
      const db = await openDatabase();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readonly');
        const store = tx.objectStore(storeName);
        const req = store.count();
        req.onsuccess = () => resolve(req.result || 0);
        req.onerror = () => reject(req.error);
      });
    },

    async clear() {
      const db = await openDatabase();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readwrite');
        const store = tx.objectStore(storeName);
        const req = store.clear();
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
      });
    }
  };
}

export const localDb = {
  registros: createTableHelper('registros'),
  equipos: createTableHelper('equipos'),
  areas: createTableHelper('areas')
};
