import { Routes } from '@angular/router';
import { DocumentListComponent } from './pages/document-list/document-list';
import { SpreadsheetEditorComponent } from './pages/spreadsheet-editor/spreadsheet-editor';

export const routes: Routes = [
  { path: '', component: DocumentListComponent },
  { path: 'editor/:id', component: SpreadsheetEditorComponent },
  { path: '**', redirectTo: '' },
];
