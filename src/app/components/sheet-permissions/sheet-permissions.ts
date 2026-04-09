import { Component, input, output, signal, OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { SheetPermissionService, SheetPermission } from '../../services/sheet-permission.service';
import { UserRole } from '../../services/user.service';

export interface SheetInfo {
  sheetId: string;
  sheetName: string;
}

@Component({
  selector: 'app-sheet-permissions',
  imports: [FormsModule],
  template: `
    <div class="perm-panel">
      <div class="perm-header">
        <h3>Sheet Permissions</h3>
        <button class="btn-close" (click)="closed.emit()">&#10005;</button>
      </div>

      <div class="perm-body">
        @for (sheet of sheets(); track sheet.sheetId) {
          <div class="sheet-section">
            <h4>{{ sheet.sheetName }}</h4>
            <div class="role-grid">
              @for (role of roles; track role) {
                <div class="role-row">
                  <span class="role-label" [attr.data-role]="role">{{ role }}</span>
                  <select
                    [ngModel]="getAccess(sheet.sheetId, role)"
                    (ngModelChange)="setAccess(sheet.sheetId, role, $event)"
                  >
                    <option value="edit">Can Edit</option>
                    <option value="view">View Only</option>
                    <option value="none">No Access</option>
                  </select>
                </div>
              }
            </div>
          </div>
        }
      </div>
    </div>
  `,
  styles: [`
    .perm-panel {
      position: absolute; top: 0; right: 0; bottom: 0;
      width: 320px; background: #fff; border-left: 1px solid #e0e0e0;
      z-index: 200; display: flex; flex-direction: column;
      box-shadow: -4px 0 16px rgba(0,0,0,0.08);
    }

    .perm-header {
      display: flex; align-items: center; justify-content: space-between;
      padding: 16px; border-bottom: 1px solid #eee;
      h3 { margin: 0; font-size: 16px; font-weight: 600; color: #1a1a2e; }
    }

    .btn-close {
      background: none; border: none; font-size: 18px; color: #999;
      cursor: pointer; padding: 4px 8px; border-radius: 4px;
      &:hover { background: #f0f0f0; color: #333; }
    }

    .perm-body { flex: 1; overflow-y: auto; padding: 16px; }

    .sheet-section {
      margin-bottom: 20px;
      h4 {
        margin: 0 0 10px; font-size: 14px; font-weight: 600; color: #333;
        padding-bottom: 6px; border-bottom: 1px solid #f0f0f0;
      }
    }

    .role-grid { display: flex; flex-direction: column; gap: 8px; }

    .role-row {
      display: flex; align-items: center; justify-content: space-between; gap: 12px;

      select {
        padding: 4px 8px; border: 1px solid #ddd; border-radius: 6px;
        font-size: 13px; background: #fff; cursor: pointer;
        &:focus { outline: none; border-color: #34a853; }
      }
    }

    .role-label {
      font-size: 13px; font-weight: 600; text-transform: capitalize;
      padding: 2px 8px; border-radius: 4px;
      &[data-role="admin"] { background: #ffebee; color: #c62828; }
      &[data-role="editor"] { background: #e3f2fd; color: #1565c0; }
      &[data-role="viewer"] { background: #f1f8e9; color: #558b2f; }
    }
  `],
})
export class SheetPermissionsComponent implements OnInit {
  documentId = input.required<string>();
  sheets = input.required<SheetInfo[]>();
  closed = output();
  permissionsChanged = output();

  roles: UserRole[] = ['editor', 'viewer'];

  /** Local cache: sheetId -> role -> access */
  private accessMap = signal<Record<string, Record<string, string>>>({});

  constructor(private permService: SheetPermissionService) {}

  async ngOnInit() {
    await this.loadPermissions();
  }

  private async loadPermissions() {
    const map: Record<string, Record<string, string>> = {};
    for (const sheet of this.sheets()) {
      map[sheet.sheetId] = {};
      for (const role of this.roles) {
        map[sheet.sheetId][role] = await this.permService.getSheetAccess(
          this.documentId(),
          sheet.sheetId,
          role,
        );
      }
    }
    this.accessMap.set(map);
  }

  getAccess(sheetId: string, role: UserRole): string {
    return this.accessMap()[sheetId]?.[role] ?? 'edit';
  }

  async setAccess(sheetId: string, role: UserRole, access: string) {
    // Update local cache
    const current = { ...this.accessMap() };
    if (!current[sheetId]) current[sheetId] = {};
    current[sheetId] = { ...current[sheetId], [role]: access };
    this.accessMap.set(current);

    // Gather all roles for this sheet
    const editRoles: UserRole[] = ['admin']; // admin always has edit
    const viewRoles: UserRole[] = [];

    for (const r of this.roles) {
      const a = current[sheetId][r] ?? 'edit';
      if (a === 'edit') editRoles.push(r);
      else if (a === 'view') viewRoles.push(r);
    }

    await this.permService.setSheetPermission(
      this.documentId(),
      sheetId,
      editRoles,
      viewRoles,
    );

    this.permissionsChanged.emit();
  }
}
