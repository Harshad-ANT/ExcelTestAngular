import {
  Component,
  OnInit,
  OnDestroy,
  ElementRef,
  ViewChild,
  signal,
  computed,
  NgZone,
} from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { FormsModule } from '@angular/forms';
import { DbService } from '../../services/db.service';
import { UserService } from '../../services/user.service';
import { SheetPermissionService, SheetAccess } from '../../services/sheet-permission.service';
import { UserSwitcherComponent } from '../../components/user-switcher/user-switcher';
import { SheetPermissionsComponent, SheetInfo } from '../../components/sheet-permissions/sheet-permissions';

import { createUniver, LocaleType, IWorkbookData, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';

import '@univerjs/preset-sheets-core/lib/index.css';

@Component({
  selector: 'app-spreadsheet-editor',
  imports: [FormsModule, UserSwitcherComponent, SheetPermissionsComponent],
  templateUrl: './spreadsheet-editor.html',
  styleUrl: './spreadsheet-editor.scss',
})
export class SpreadsheetEditorComponent implements OnInit, OnDestroy {
  @ViewChild('univerContainer', { static: true }) containerRef!: ElementRef<HTMLDivElement>;

  docName = signal('');
  isSaving = signal(false);
  lastSaved = signal<Date | null>(null);
  isEditing = signal(false);
  showPermPanel = signal(false);
  sheetList = signal<SheetInfo[]>([]);
  accessMessage = signal('');

  isAdmin = computed(() => this.userService.isAdmin());

  documentId = '';
  private univerAPI: any;
  private univer: any;
  private autoSaveInterval: ReturnType<typeof setInterval> | null = null;

  constructor(
    private route: ActivatedRoute,
    private router: Router,
    private db: DbService,
    private ngZone: NgZone,
    public userService: UserService,
    private permService: SheetPermissionService,
  ) {}

  async ngOnInit() {
    this.documentId = this.route.snapshot.paramMap.get('id')!;
    const stored = await this.db.getDocument(this.documentId);

    if (!stored) {
      this.router.navigate(['/']);
      return;
    }

    this.docName.set(stored.name);
    const workbookData: IWorkbookData = JSON.parse(stored.data);

    this.initUniver(workbookData);
    await this.applyPermissions();
    this.startAutoSave();
  }

  private initUniver(workbookData: IWorkbookData) {
    const container = this.containerRef.nativeElement;

    const { univer, univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: mergeLocales(sheetsCoreEnUS),
      presets: [
        UniverSheetsCorePreset({
          container,
        }),
      ],
    });

    this.univer = univer;
    this.univerAPI = univerAPI;

    univerAPI.createWorkbook(workbookData);
  }

  async applyPermissions() {
    const workbook = this.univerAPI?.getActiveWorkbook();
    if (!workbook) return;

    const permission = workbook.getPermission();
    const unitId = workbook.getId();
    const sheets = workbook.getSheets();
    const role = this.userService.activeUser().role;

    const sheetInfos: SheetInfo[] = [];

    for (const sheet of sheets) {
      const sheetId = sheet.getSheetId();
      const sheetName = sheet.getSheetName();
      sheetInfos.push({ sheetId, sheetName });

      const access: SheetAccess = await this.permService.getSheetAccess(
        this.documentId,
        sheetId,
        role,
      );

      if (access === 'edit') {
        // Full access — remove any existing protection
        permission.removeWorksheetPermission(unitId, sheetId);
      } else if (access === 'view') {
        // View only — protect the sheet so it can't be edited
        await permission.addWorksheetBasePermission(unitId, sheetId);
        await permission.setWorksheetPermissionPoint(
          unitId, sheetId,
          permission.permissionPointsDefinition.WorksheetEditPermission,
          false,
        );
        await permission.setWorksheetPermissionPoint(
          unitId, sheetId,
          permission.permissionPointsDefinition.WorksheetSetCellValuePermission,
          false,
        );
        await permission.setWorksheetPermissionPoint(
          unitId, sheetId,
          permission.permissionPointsDefinition.WorksheetSetCellStylePermission,
          false,
        );
        await permission.setWorksheetPermissionPoint(
          unitId, sheetId,
          permission.permissionPointsDefinition.WorksheetInsertRowPermission,
          false,
        );
        await permission.setWorksheetPermissionPoint(
          unitId, sheetId,
          permission.permissionPointsDefinition.WorksheetInsertColumnPermission,
          false,
        );
        await permission.setWorksheetPermissionPoint(
          unitId, sheetId,
          permission.permissionPointsDefinition.WorksheetDeleteRowPermission,
          false,
        );
        await permission.setWorksheetPermissionPoint(
          unitId, sheetId,
          permission.permissionPointsDefinition.WorksheetDeleteColumnPermission,
          false,
        );
      }
    }

    this.sheetList.set(sheetInfos);

    // Show access message for non-admin users
    if (role !== 'admin') {
      const activeSheet = workbook.getActiveSheet();
      const activeAccess = await this.permService.getSheetAccess(
        this.documentId,
        activeSheet.getSheetId(),
        role,
      );
      this.accessMessage.set(
        activeAccess === 'view' ? 'View Only — You cannot edit this sheet' : '',
      );
    } else {
      this.accessMessage.set('');
    }
  }

  private startAutoSave() {
    this.ngZone.runOutsideAngular(() => {
      this.autoSaveInterval = setInterval(() => {
        this.ngZone.run(() => this.saveDocument());
      }, 30000);
    });
  }

  async saveDocument() {
    if (!this.univerAPI || this.isSaving()) return;

    this.isSaving.set(true);
    try {
      const workbook = this.univerAPI.getActiveWorkbook();
      if (!workbook) return;

      const snapshot = workbook.getSnapshot();
      const now = Date.now();

      await this.db.saveDocument({
        id: this.documentId,
        name: this.docName(),
        createdAt: 0,
        updatedAt: now,
        data: JSON.stringify(snapshot),
      });

      const existing = await this.db.getDocument(this.documentId);
      if (existing && existing.createdAt === 0) {
        existing.createdAt = now;
        await this.db.saveDocument(existing);
      }

      this.lastSaved.set(new Date(now));
    } finally {
      this.isSaving.set(false);
    }
  }

  startEditName() {
    this.isEditing.set(true);
  }

  finishEditName() {
    this.isEditing.set(false);
    const name = this.docName().trim();
    if (name) {
      this.db.renameDocument(this.documentId, name);
    }
  }

  goBack() {
    this.saveDocument().then(() => this.router.navigate(['/']));
  }

  async addSheet() {
    const workbook = this.univerAPI?.getActiveWorkbook();
    if (!workbook) return;
    const sheets = workbook.getSheets();
    const name = `Sheet${sheets.length + 1}`;
    const newSheet = workbook.create(name, 100, 26);
    const sheetId = newSheet.getSheetId();

    // Initialize default permissions for the new sheet
    await this.permService.initSheetDefaults(this.documentId, sheetId);

    // Refresh sheet list and permissions
    await this.applyPermissions();
  }

  async onUserChanged() {
    // Re-apply permissions when user role changes
    await this.applyPermissions();
  }

  async onPermissionsChanged() {
    await this.applyPermissions();
  }

  togglePermPanel() {
    // Refresh sheet list before showing
    const workbook = this.univerAPI?.getActiveWorkbook();
    if (workbook) {
      const sheets = workbook.getSheets();
      this.sheetList.set(
        sheets.map((s: any) => ({ sheetId: s.getSheetId(), sheetName: s.getSheetName() })),
      );
    }
    this.showPermPanel.update(v => !v);
  }

  ngOnDestroy() {
    if (this.autoSaveInterval) {
      clearInterval(this.autoSaveInterval);
    }
    this.saveDocument();
    if (this.univer) {
      this.univer.dispose();
    }
  }
}
