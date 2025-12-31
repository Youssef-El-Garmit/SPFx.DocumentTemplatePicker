import { IBreadcrumbItem } from '@fluentui/react';
import { UrlUtils } from './UrlUtils';

/**
 * Utility functions for breadcrumb operations
 */

export interface IBreadcrumbConfig {
  folderPath: string;
  libraryRootUrl: string;
  libraryTitle: string;
  onRootClick: () => void;
  onFolderClick: (folderPath: string) => void;
}

export class BreadcrumbUtils {
  /**
   * Build breadcrumb items from folder path
   * Only builds breadcrumb if not at library root
   * Last item (current location) is not clickable
   */
  static buildBreadcrumb(config: IBreadcrumbConfig): IBreadcrumbItem[] {
    const { folderPath, libraryRootUrl, libraryTitle, onRootClick, onFolderClick } = config;
    const items: IBreadcrumbItem[] = [];

    // Only build breadcrumb if we're not at the library root
    if (folderPath && folderPath !== libraryRootUrl) {
      // Get the relative path from library root
      const relativePath = folderPath.replace(libraryRootUrl, '').replace(/^\//, '');
      
      if (relativePath) {
        // Add root item
        items.push({ 
          text: libraryTitle, 
          key: 'root', 
          onClick: onRootClick
        });

        // Split the relative path and add each folder
        const pathParts = relativePath.split('/').filter((p: string) => p);
        let currentPath = libraryRootUrl;
        
        pathParts.forEach((part, index) => {
          // Normalize path: ensure single slash between parts
          const endsWithSlash = currentPath.lastIndexOf('/') === currentPath.length - 1;
          currentPath = endsWithSlash 
            ? currentPath + part 
            : currentPath + '/' + part;
          const isLast = index === pathParts.length - 1;
          // Capture the current path value for this iteration
          const pathForThisItem = currentPath;
          items.push({
            text: part,
            key: `folder-${index}`,
            // Last item (current location) should not be clickable
            onClick: isLast ? undefined : () => {
              // Use the captured path for this specific breadcrumb item
              onFolderClick(pathForThisItem);
            }
          });
        });
      }
    }

    return items;
  }

  /**
   * Build simplified breadcrumb for success step (only last folder with ellipsis)
   */
  static buildSuccessBreadcrumb(
    destinationFolderPath: string,
    destinationLibraryRootUrl: string,
    libraryTitle: string,
    webUrl: string
  ): IBreadcrumbItem[] {
    const fullBreadcrumbItems = this.buildBreadcrumb({
      folderPath: destinationFolderPath,
      libraryRootUrl: destinationLibraryRootUrl,
      libraryTitle,
      onRootClick: () => {},
      onFolderClick: () => {}
    });
    
    const successBreadcrumbItems: IBreadcrumbItem[] = [];
    if (fullBreadcrumbItems.length > 0) {
      const lastItem = fullBreadcrumbItems[fullBreadcrumbItems.length - 1];
      
      // Build folder URL correctly - avoid duplicating the site path
      const folderUrl = UrlUtils.buildFullUrl(destinationFolderPath, webUrl);
      
      // Add ellipsis if there are multiple items
      if (fullBreadcrumbItems.length > 1) {
        successBreadcrumbItems.push({
          text: '...',
          key: 'ellipsis',
          isCurrentItem: false
        });
      }
      
      // Add the last item (the destination folder) as clickable
      successBreadcrumbItems.push({
        text: lastItem.text,
        key: lastItem.key,
        isCurrentItem: true,
        onClick: () => {
          window.open(folderUrl, '_blank');
        }
      });
    }
    
    return successBreadcrumbItems;
  }
}

