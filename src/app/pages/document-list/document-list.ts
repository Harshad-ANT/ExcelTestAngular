import { Component, OnInit, signal } from '@angular/core';
import { Router } from '@angular/router';
import { DatePipe } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DbService, DocumentMeta } from '../../services/db.service';
import { TemplateService, DocumentTemplate } from '../../services/template.service';
import { UserSwitcherComponent } from '../../components/user-switcher/user-switcher';

@Component({
  selector: 'app-document-list',
  imports: [DatePipe, FormsModule, UserSwitcherComponent],
  templateUrl: './document-list.html',
  styleUrl: './document-list.scss',
})
export class DocumentListComponent implements OnInit {
  documents = signal<DocumentMeta[]>([]);
  newDocName = signal('');
  selectedTemplate = signal('bcs-cam-note');
  templates: DocumentTemplate[] = [];
  isCreating = signal(false);

  constructor(
    private db: DbService,
    private router: Router,
    private templateService: TemplateService,
  ) {
    this.templates = this.templateService.getTemplates();
  }

  async ngOnInit() {
    await this.loadDocuments();
  }

  async loadDocuments() {
    this.documents.set(await this.db.listDocuments());
  }

  async createDocument() {
    if (this.isCreating()) return;
    this.isCreating.set(true);

    try {
      const name = this.newDocName().trim() || 'Untitled Spreadsheet';
      const id = crypto.randomUUID();
      const now = Date.now();
      const templateId = this.selectedTemplate();

      let workbookData: any;

      if (templateId === 'blank') {
        workbookData = {
          id,
          name,
          appVersion: '0.1.0',
          locale: 'EN_US',
          styles: {},
          sheetOrder: ['sheet-01'],
          sheets: {
            'sheet-01': {
              id: 'sheet-01',
              name: 'Sheet1',
              rowCount: 100,
              columnCount: 26,
              cellData: {},
            },
          },
        };
      } else {
        workbookData = await this.templateService.getTemplateData(templateId);
        workbookData.id = id;
        workbookData.name = name;
      }

      await this.db.saveDocument({
        id,
        name,
        createdAt: now,
        updatedAt: now,
        data: JSON.stringify(workbookData),
      });

      this.newDocName.set('');
      this.router.navigate(['/editor', id]);
    } finally {
      this.isCreating.set(false);
    }
  }

  openDocument(id: string) {
    this.router.navigate(['/editor', id]);
  }

  async deleteDocument(event: Event, id: string) {
    event.stopPropagation();
    if (confirm('Delete this document?')) {
      await this.db.deleteDocument(id);
      await this.loadDocuments();
    }
  }
}
