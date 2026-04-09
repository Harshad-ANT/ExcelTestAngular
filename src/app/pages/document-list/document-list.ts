import { Component, OnInit, signal } from '@angular/core';
import { Router } from '@angular/router';
import { DatePipe } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DbService, DocumentMeta } from '../../services/db.service';

@Component({
  selector: 'app-document-list',
  imports: [DatePipe, FormsModule],
  templateUrl: './document-list.html',
  styleUrl: './document-list.scss',
})
export class DocumentListComponent implements OnInit {
  documents = signal<DocumentMeta[]>([]);
  newDocName = signal('');

  constructor(
    private db: DbService,
    private router: Router,
  ) {}

  async ngOnInit() {
    await this.loadDocuments();
  }

  async loadDocuments() {
    this.documents.set(await this.db.listDocuments());
  }

  async createDocument() {
    const name = this.newDocName().trim() || 'Untitled Spreadsheet';
    const id = crypto.randomUUID();
    const now = Date.now();

    const defaultData = {
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

    await this.db.saveDocument({
      id,
      name,
      createdAt: now,
      updatedAt: now,
      data: JSON.stringify(defaultData),
    });

    this.newDocName.set('');
    this.router.navigate(['/editor', id]);
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
