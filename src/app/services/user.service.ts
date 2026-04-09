import { Injectable, signal, computed } from '@angular/core';

export type UserRole = 'admin' | 'editor' | 'viewer';

export interface AppUser {
  id: string;
  name: string;
  role: UserRole;
}

const DEFAULT_USERS: AppUser[] = [
  { id: 'admin-1', name: 'Admin User', role: 'admin' },
  { id: 'editor-1', name: 'Editor User', role: 'editor' },
  { id: 'viewer-1', name: 'Viewer User', role: 'viewer' },
];

const STORAGE_KEY = 'activeUserId';

@Injectable({ providedIn: 'root' })
export class UserService {
  readonly users = signal<AppUser[]>(DEFAULT_USERS);
  readonly activeUserId = signal<string>(localStorage.getItem(STORAGE_KEY) || DEFAULT_USERS[0].id);

  readonly activeUser = computed(() =>
    this.users().find(u => u.id === this.activeUserId()) ?? this.users()[0]
  );

  switchUser(userId: string) {
    this.activeUserId.set(userId);
    localStorage.setItem(STORAGE_KEY, userId);
  }

  isAdmin(): boolean {
    return this.activeUser().role === 'admin';
  }

  isViewer(): boolean {
    return this.activeUser().role === 'viewer';
  }
}
