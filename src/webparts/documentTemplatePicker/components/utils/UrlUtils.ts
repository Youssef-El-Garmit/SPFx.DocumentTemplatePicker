/**
 * Utility functions for URL operations
 */

export class UrlUtils {
  /**
   * Get thumbnail URL using SharePoint getpreview.ashx API
   * @param serverRelativeUrl Server-relative URL of the file
   * @param webUrl Absolute URL of the web
   * @param resolution Resolution: 0=300px, 1=480px, 2=750px, 3=1024px, 4=1600px, 5=2560px, 6=Original
   */
  static getThumbnailUrl(serverRelativeUrl: string, webUrl: string, resolution: number = 1): string {
    const encodedPath = encodeURIComponent(serverRelativeUrl);
    return `${webUrl}/_layouts/15/getpreview.ashx?path=${encodedPath}&resolution=${resolution}`;
  }

  /**
   * Get preview URL for document
   * Uses UniqueId with embed.aspx for Office documents, fallback to getpreview.ashx for images
   */
  static getPreviewUrl(uniqueId: string | undefined, serverRelativeUrl: string, fileName: string, webUrl: string): string | undefined {
    // Use UniqueId with embed.aspx like DocumentPreview.tsx
    if (uniqueId) {
      // Add parameters to control iframe behavior
      // action=embedview forces fit-to-width behavior
      return `${webUrl}/_layouts/15/embed.aspx?UniqueId=${uniqueId}&action=embedview`;
    }
    
    // Fallback: if no UniqueId, try getpreview.ashx for images
    const fileExtension = fileName.split('.').pop()?.toLowerCase() || '';
    const imageTypes = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'tif'];
    if (imageTypes.indexOf(fileExtension) !== -1) {
      // Use resolution=3 (1024px) for image previews in dialog
      return this.getThumbnailUrl(serverRelativeUrl, webUrl, 3);
    }
    
    // No preview available
    return undefined;
  }

  /**
   * Build full URL from server-relative URL, avoiding duplication
   */
  static buildFullUrl(serverRelativeUrl: string, webUrl: string): string {
    const webUrlPath = new URL(webUrl).pathname;
    let relativePath = serverRelativeUrl;
    
    // If serverRelativeUrl already starts with the web path, remove it
    if (relativePath.startsWith(webUrlPath)) {
      relativePath = relativePath.substring(webUrlPath.length);
    }
    
    // Ensure relativePath starts with /
    if (!relativePath.startsWith('/')) {
      relativePath = '/' + relativePath;
    }
    
    return `${webUrl}${relativePath}`;
  }

  /**
   * Normalize folder path for ParentLink comparison
   */
  static normalizeFolderUrl(url: string): string {
    return url.replace(/\/Forms\/AllItems\.aspx$/i, '').replace(/\/$/, '').trim().toLowerCase();
  }

  /**
   * Build folder ServerRelativeUrl from parent and folder name
   */
  static buildFolderPath(parentPath: string, folderName: string): string {
    const endsWithSlash = parentPath.lastIndexOf('/') === parentPath.length - 1;
    return endsWithSlash 
      ? `${parentPath}${folderName}`
      : `${parentPath}/${folderName}`;
  }
}

