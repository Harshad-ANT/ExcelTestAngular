import { Injectable } from '@angular/core';
import { UserRole } from './user.service';

export type SheetAccess = 'edit' | 'view' | 'none';

export interface SheetPermission {
  sheetId: string;
  /** Roles allowed to edit this sheet */
  editRoles: UserRole[];
  /** Roles allowed to view this sheet */
  viewRoles: UserRole[];
}

export interface DocumentPermissions {
  documentId: string;
  /** Owner role always has full access. Per-sheet overrides below. */
  sheets: SheetPermission[];
}

const DB_NAME = 'UniversSheetsDB';
const DB_VERSION = 2;
const PERM_STORE = 'permissions';
const DOC_STORE = 'documents';

@Injectable({ providedIn: 'root' })
export class SheetPermissionService {
  private dbPromise: Promise<IDBDatabase>;

  constructor() {
    this.dbPromise = this.openDb();
  }

  private openDb(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(DOC_STORE)) {
          const store = db.createObjectStore(DOC_STORE, { keyPath: 'id' });
          store.createIndex('updatedAt', 'updatedAt', { unique: false });
        }
        if (!db.objectStoreNames.contains(PERM_STORE)) {
          db.createObjectStore(PERM_STORE, { keyPath: 'documentId' });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async getPermissions(documentId: string): Promise<DocumentPermissions | undefined> {
    const db = await this.dbPromise;
    return new Promise((resolve, reject) => {
      const tx = db.transaction(PERM_STORE, 'readonly');
      const req = tx.objectStore(PERM_STORE).get(documentId);
      req.onsuccess = () => resolve(req.result ?? undefined);
      req.onerror = () => reject(req.error);
    });
  }

  async savePermissions(perms: DocumentPermissions): Promise<void> {
    const db = await this.dbPromise;
    return new Promise((resolve, reject) => {
      const tx = db.transaction(PERM_STORE, 'readwrite');
      tx.objectStore(PERM_STORE).put(perms);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async getSheetAccess(documentId: string, sheetId: string, role: UserRole): Promise<SheetAccess> {
    if (role === 'admin') return 'edit';

    const perms = await this.getPermissions(documentId);
    if (!perms) return role === 'viewer' ? 'view' : 'edit'; // default: editors can edit, viewers can view

    const sheetPerm = perms.sheets.find(s => s.sheetId === sheetId);
    if (!sheetPerm) return role === 'viewer' ? 'view' : 'edit'; // no rule = default

    if (sheetPerm.editRoles.includes(role)) return 'edit';
    if (sheetPerm.viewRoles.includes(role)) return 'view';
    return 'none';
  }

  async setSheetPermission(
    documentId: string,
    sheetId: string,
    editRoles: UserRole[],
    viewRoles: UserRole[],
  ): Promise<void> {
    let perms = await this.getPermissions(documentId);
    if (!perms) {
      perms = { documentId, sheets: [] };
    }

    const idx = perms.sheets.findIndex(s => s.sheetId === sheetId);
    const entry: SheetPermission = { sheetId, editRoles, viewRoles };

    if (idx >= 0) {
      perms.sheets[idx] = entry;
    } else {
      perms.sheets.push(entry);
    }

    await this.savePermissions(perms);
  }

  /** Initialize default permissions for a new sheet */
  async initSheetDefaults(documentId: string, sheetId: string): Promise<void> {
    await this.setSheetPermission(
      documentId,
      sheetId,
      ['admin', 'editor'],
      ['viewer'],
    );
  }
}
