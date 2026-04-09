import { Injectable } from '@angular/core';

export interface DocumentMeta {
  id: string;
  name: string;
  createdAt: number;
  updatedAt: number;
}

export interface StoredDocument {
  id: string;
  name: string;
  createdAt: number;
  updatedAt: number;
  /** Serialized IWorkbookData JSON */
  data: string;
}

const DB_NAME = 'UniversSheetsDB';
const DB_VERSION = 2;
const STORE_NAME = 'documents';

@Injectable({ providedIn: 'root' })
export class DbService {
  private dbPromise: Promise<IDBDatabase>;

  constructor() {
    this.dbPromise = this.openDb();
  }

  private openDb(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          const store = db.createObjectStore(STORE_NAME, { keyPath: 'id' });
          store.createIndex('updatedAt', 'updatedAt', { unique: false });
        }
        if (!db.objectStoreNames.contains('permissions')) {
          db.createObjectStore('permissions', { keyPath: 'documentId' });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async listDocuments(): Promise<DocumentMeta[]> {
    const db = await this.dbPromise;
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const store = tx.objectStore(STORE_NAME);
      const request = store.getAll();
      request.onsuccess = () => {
        const docs: StoredDocument[] = request.result;
        const metas: DocumentMeta[] = docs
          .map(d => ({ id: d.id, name: d.name, createdAt: d.createdAt, updatedAt: d.updatedAt }))
          .sort((a, b) => b.updatedAt - a.updatedAt);
        resolve(metas);
      };
      request.onerror = () => reject(request.error);
    });
  }

  async getDocument(id: string): Promise<StoredDocument | undefined> {
    const db = await this.dbPromise;
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const request = tx.objectStore(STORE_NAME).get(id);
      request.onsuccess = () => resolve(request.result ?? undefined);
      request.onerror = () => reject(request.error);
    });
  }

  async saveDocument(doc: StoredDocument): Promise<void> {
    const db = await this.dbPromise;
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put(doc);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async deleteDocument(id: string): Promise<void> {
    const db = await this.dbPromise;
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).delete(id);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async renameDocument(id: string, name: string): Promise<void> {
    const doc = await this.getDocument(id);
    if (doc) {
      doc.name = name;
      doc.updatedAt = Date.now();
      await this.saveDocument(doc);
    }
  }
}
