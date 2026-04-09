import {
  Component,
  OnInit,
  OnDestroy,
  ElementRef,
  ViewChild,
  signal,
  NgZone,
} from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { FormsModule } from '@angular/forms';
import { DbService, StoredDocument } from '../../services/db.service';

import { createUniver, LocaleType, IWorkbookData, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';

import '@univerjs/preset-sheets-core/lib/index.css';

@Component({
  selector: 'app-spreadsheet-editor',
  imports: [FormsModule],
  templateUrl: './spreadsheet-editor.html',
  styleUrl: './spreadsheet-editor.scss',
})
export class SpreadsheetEditorComponent implements OnInit, OnDestroy {
  @ViewChild('univerContainer', { static: true }) containerRef!: ElementRef<HTMLDivElement>;

  docName = signal('');
  isSaving = signal(false);
  lastSaved = signal<Date | null>(null);
  isEditing = signal(false);

  private documentId = '';
  private univerAPI: any;
  private univer: any;
  private autoSaveInterval: ReturnType<typeof setInterval> | null = null;

  constructor(
    private route: ActivatedRoute,
    private router: Router,
    private db: DbService,
    private ngZone: NgZone,
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
        createdAt: 0, // will be preserved by put
        updatedAt: now,
        data: JSON.stringify(snapshot),
      });

      // Preserve createdAt
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
    workbook.create(name, 100, 26);
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
