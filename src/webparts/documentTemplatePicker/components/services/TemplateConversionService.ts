import * as JSZip from 'jszip';

/**
 * Service for converting Office template files to document files.
 * Handles the OOXML content type conversion needed when creating documents from templates.
 * 
 * When a .dotx file is simply copied and renamed to .docx, Word may report corruption
 * because the internal Content_Types.xml still references the template MIME type.
 * This service properly converts the internal content types to match the document format.
 */
export class TemplateConversionService {
  /**
   * Content type mappings for template to document conversion.
   * Maps template MIME types to their corresponding document MIME types.
   */
  private static readonly CONTENT_TYPE_MAPPINGS: Record<string, string> = {
    // Word templates
    'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml':
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    'application/vnd.ms-word.template.macroEnabled.main+xml':
      'application/vnd.ms-word.document.macroEnabled.main+xml',
    // Excel templates
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml':
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    'application/vnd.ms-excel.template.macroEnabled.main+xml':
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    // PowerPoint templates
    'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml':
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'application/vnd.ms-powerpoint.template.macroEnabled.main+xml':
      'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml'
  };

  /**
   * Template extensions that require conversion
   */
  private static readonly TEMPLATE_EXTENSIONS = ['dotx', 'dotm', 'xltx', 'xltm', 'potx', 'potm'];

  /**
   * Check if a file extension is a template format that needs conversion
   */
  public static isTemplateExtension(extension: string): boolean {
    return this.TEMPLATE_EXTENSIONS.includes(extension.toLowerCase());
  }

  /**
   * Convert a template file (dotx, xltx, potx, etc.) to a document file (docx, xlsx, pptx).
   * This modifies the [Content_Types].xml inside the OOXML package to change
   * the content type from template to document format.
   * 
   * @param templateBlob - The template file as a Blob
   * @returns The converted document as a Blob
   */
  public static async convertTemplateToDocument(templateBlob: Blob): Promise<Blob> {
    // Load the ZIP archive
    const zip = await JSZip.loadAsync(templateBlob);

    // Get the [Content_Types].xml file
    const contentTypesFile = zip.file('[Content_Types].xml');
    if (!contentTypesFile) {
      throw new Error('Invalid Office document: [Content_Types].xml not found');
    }

    // Read and parse the content types XML
    let contentTypesXml = await contentTypesFile.async('text');

    // Replace all template content types with document content types
    // Using string split/join instead of RegExp for safety (no dynamic regex construction)
    for (const [templateType, documentType] of Object.entries(this.CONTENT_TYPE_MAPPINGS)) {
      contentTypesXml = contentTypesXml.split(templateType).join(documentType);
    }

    // Update the [Content_Types].xml in the archive
    zip.file('[Content_Types].xml', contentTypesXml);

    // Generate the new document blob
    const documentBlob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    return documentBlob;
  }
}
