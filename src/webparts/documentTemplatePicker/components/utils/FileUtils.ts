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

