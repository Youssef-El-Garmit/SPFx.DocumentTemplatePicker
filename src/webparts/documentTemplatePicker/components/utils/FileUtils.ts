/**
 * Utility functions for file operations
 */

export class FileUtils {
  /**
   * Get file extension from filename
   */
  static getFileExtension(fileName: string): string {
    return fileName.split('.').pop()?.toLowerCase() || '';
  }

  /**
   * Convert template extension to document extension.
   * When creating a document from a template, the final file should have
   * the document extension, not the template extension.
   * E.g., .dotx -> .docx, .xltx -> .xlsx, .potx -> .pptx
   */
  static getDocumentExtension(templateExtension: string): string {
    const ext = templateExtension.toLowerCase();
    const templateToDocumentMap: Record<string, string> = {
      'dotx': 'docx',
      'dotm': 'docm',
      'dot': 'doc',
      'xltx': 'xlsx',
      'xltm': 'xlsm',
      'xlt': 'xls',
      'potx': 'pptx',
      'potm': 'pptm',
      'pot': 'ppt'
    };
    return templateToDocumentMap[ext] || ext;
  }

  /**
   * Check if a file extension is a template format that should be converted
   * to a document format when creating a new file from it.
   */
  static isTemplateExtension(extension: string): boolean {
    const ext = extension.toLowerCase();
    const templateExtensions = ['dotx', 'dotm', 'dot', 'xltx', 'xltm', 'xlt', 'potx', 'potm', 'pot'];
    return templateExtensions.indexOf(ext) !== -1;
  }

  /**
   * Get Fluent UI icon name based on file extension (Word/Excel/PPT families).
   */
  static getFileIcon(fileName: string): string {
    const ext = this.getFileExtension(fileName);
    if (['doc', 'docx', 'dot', 'dotx', 'dotm', 'docm'].indexOf(ext) !== -1) return 'WordDocument';
    if (['xls', 'xlsx', 'xlt', 'xltx', 'xltm', 'xlsm'].indexOf(ext) !== -1) return 'ExcelDocument';
    if (['ppt', 'pptx', 'pot', 'potx', 'potm', 'pptm'].indexOf(ext) !== -1) return 'PowerPointDocument';
    if (ext === 'pdf') return 'PDF';
    if (ext === 'txt') return 'TextDocument';
    return 'Document';
  }

  /**
   * Get Office protocol for opening document in desktop app.
   * Single source of truth: all Word/Excel/PowerPoint formats (including templates).
   */
  static getOfficeProtocol(fileName: string): string | undefined {
    const ext = this.getFileExtension(fileName);
    const word = ['doc', 'docx', 'dot', 'dotx', 'dotm', 'docm'];
    const excel = ['xls', 'xlsx', 'xlt', 'xltx', 'xltm', 'xlsm'];
    const powerpoint = ['ppt', 'pptx', 'pot', 'potx', 'potm', 'pptm'];
    if (word.indexOf(ext) !== -1) return 'ms-word:ofe|u|';
    if (excel.indexOf(ext) !== -1) return 'ms-excel:ofe|u|';
    if (powerpoint.indexOf(ext) !== -1) return 'ms-powerpoint:ofe|u|';
    return undefined;
  }

  /**
   * Check if file is an Office document (no separate list – derived from protocol).
   */
  static isOfficeDocument(fileName: string): boolean {
    return this.getOfficeProtocol(fileName) !== undefined;
  }
}

