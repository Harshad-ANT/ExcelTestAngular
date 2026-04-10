import { Injectable } from '@angular/core';

export interface DocumentTemplate {
  id: string;
  name: string;
  description: string;
}

export const AVAILABLE_TEMPLATES: DocumentTemplate[] = [
  {
    id: 'blank',
    name: 'Blank Spreadsheet',
    description: 'Empty spreadsheet with one sheet',
  },
  {
    id: 'bcs-cam-note',
    name: 'BCS CAM Note',
    description: 'Credit Appraisal Memorandum — 21 sheets with formulas',
  },
];

@Injectable({ providedIn: 'root' })
export class TemplateService {
  private templateCache: Record<string, any> = {};

  getTemplates(): DocumentTemplate[] {
    return AVAILABLE_TEMPLATES;
  }

  async getTemplateData(templateId: string): Promise<any> {
    if (templateId === 'blank') {
      return null; // Caller handles blank template
    }

    if (this.templateCache[templateId]) {
      return structuredClone(this.templateCache[templateId]);
    }

    const url = this.getTemplateUrl(templateId);
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Failed to load template: ${templateId}`);
    }

    const data = await response.json();
    this.templateCache[templateId] = data;
    return structuredClone(data);
  }

  private getTemplateUrl(templateId: string): string {
    const urls: Record<string, string> = {
      'bcs-cam-note': '/assets/bcs-cam-note-template.json',
    };
    return urls[templateId] ?? '';
  }
}
