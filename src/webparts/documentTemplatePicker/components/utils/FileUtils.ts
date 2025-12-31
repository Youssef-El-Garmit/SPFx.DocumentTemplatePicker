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
   * Get Fluent UI icon name based on file extension
   */
  static getFileIcon(fileName: string): string {
    const extension = this.getFileExtension(fileName);
    switch (extension) {
      case 'docx':
      case 'doc':
        return 'WordDocument';
      case 'xlsx':
      case 'xls':
        return 'ExcelDocument';
      case 'pptx':
      case 'ppt':
        return 'PowerPointDocument';
      case 'pdf':
        return 'PDF';
      case 'txt':
        return 'TextDocument';
      default:
        return 'Document';
    }
  }

  /**
   * Check if file extension is an Office document
   */
  static isOfficeDocument(fileName: string): boolean {
    const extension = this.getFileExtension(fileName);
    const officeExtensions = ['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt'];
    return officeExtensions.indexOf(extension) !== -1;
  }

  /**
   * Get Office protocol for opening document in desktop app
   */
  static getOfficeProtocol(fileName: string): string | undefined {
    const extension = this.getFileExtension(fileName);
    switch (extension) {
      case 'docx':
      case 'doc':
        return 'ms-word:ofe|u|';
      case 'xlsx':
      case 'xls':
        return 'ms-excel:ofe|u|';
      case 'pptx':
      case 'ppt':
        return 'ms-powerpoint:ofe|u|';
      default:
        return undefined;
    }
  }
}

