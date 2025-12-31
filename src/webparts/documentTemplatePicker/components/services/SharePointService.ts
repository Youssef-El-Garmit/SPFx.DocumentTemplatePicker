import { spfi, SPFx } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import "@pnp/sp/search";
import type { ITemplateItem, IFolderItem } from '../IDocumentTemplatePickerProps';
import { FileUtils } from '../utils/FileUtils';
import { UrlUtils } from '../utils/UrlUtils';

/**
 * Service for SharePoint operations
 */
export class SharePointService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Load templates (files and folders) from a library
   */
  async loadTemplates(libraryId: string, folderPath: string = ''): Promise<{
    items: ITemplateItem[];
    folders: IFolderItem[];
    libraryRootUrl: string;
  }> {
    const sp = spfi().using(SPFx(this.context));

    const list = sp.web.lists.getById(libraryId);
    const rootFolder = await list.rootFolder();
    const libraryRootUrl = rootFolder.ServerRelativeUrl;
    const folderServerRelativeUrl = folderPath || libraryRootUrl;
    
    // Get folders (excluding hidden folders)
    const folder = sp.web.getFolderByServerRelativePath(folderServerRelativeUrl);
    const allFolders = await folder.folders.select('Name', 'ServerRelativeUrl', 'Properties')();
    
    // Filter out hidden folders
    const folders = allFolders.filter((f: any) => {
      const name = f.Name || '';
      return !name.startsWith('.') && name.toLowerCase() !== 'forms';
    });
    
    // Get files (excluding hidden files) with thumbnail and preview URLs
    const allFiles = await folder.files
      .select('Name', 'ServerRelativeUrl', 'TimeLastModified', 'Length', 'ModifiedBy/Title', 'ServerRedirectedEmbedUrl', 'ListItemAllFields/Thumbnail', 'ListItemAllFields/PictureThumbnailURL', 'ListItemAllFields/UniqueId')
      .expand('ModifiedBy', 'ListItemAllFields')();
    
    // Filter out hidden files
    const files = allFiles.filter((f: any) => {
      const name = f.Name || '';
      return !name.startsWith('.');
    });

    const folderItems: IFolderItem[] = folders.map((f: any) => ({
      key: f.ServerRelativeUrl,
      name: f.Name,
      serverRelativeUrl: f.ServerRelativeUrl,
      path: f.ServerRelativeUrl
    }));

    const webUrl = this.context.pageContext.web.absoluteUrl;
    const fileItems: ITemplateItem[] = files.map((f: any) => {
      const fileUrl = f.ServerRedirectedEmbedUrl || `${webUrl}${f.ServerRelativeUrl}`;
      const uniqueId = f.ListItemAllFields?.UniqueId;
      const thumbnailUrl = UrlUtils.getThumbnailUrl(f.ServerRelativeUrl, webUrl);
      const previewUrl = UrlUtils.getPreviewUrl(uniqueId, f.ServerRelativeUrl, f.Name, webUrl);
      
      return {
        key: f.ServerRelativeUrl,
        name: f.Name,
        fileType: FileUtils.getFileExtension(f.Name),
        fileIcon: FileUtils.getFileIcon(f.Name),
        fileUrl: fileUrl,
        serverRelativeUrl: f.ServerRelativeUrl,
        uniqueId: uniqueId,
        modified: new Date(f.TimeLastModified),
        modifiedBy: f.ModifiedBy?.Title || '',
        size: f.Length || 0,
        isFolder: false,
        folderPath: folderServerRelativeUrl,
        thumbnailUrl: thumbnailUrl,
        previewUrl: previewUrl
      };
    });

    const folderTemplateItems: ITemplateItem[] = folderItems.map((f) => ({
      key: f.serverRelativeUrl,
      name: f.name,
      fileType: '',
      fileIcon: 'Folder',
      fileUrl: '',
      serverRelativeUrl: f.serverRelativeUrl,
      modified: new Date(),
      modifiedBy: '',
      size: 0,
      isFolder: true,
      folderPath: folderServerRelativeUrl
    }));

    const allItems = [...folderTemplateItems, ...fileItems].sort((a, b) => {
      if (a.isFolder !== b.isFolder) {
        return a.isFolder ? -1 : 1;
      }
      return a.name.localeCompare(b.name);
    });

    return {
      items: allItems,
      folders: folderItems,
      libraryRootUrl
    };
  }

  /**
   * Load destination folders with pagination and search support
   */
  async loadDestinationFolders(
    libraryId: string,
    folderPath: string = '',
    pageSize: number = 50,
    skip: number = 0,
    searchQuery: string = ''
  ): Promise<{
    folders: IFolderItem[];
    hasMore: boolean;
    libraryRootUrl: string;
  }> {
    const sp = spfi().using(SPFx(this.context));

    const list = sp.web.lists.getById(libraryId);
    const rootFolder = await list.rootFolder();
    const folderServerRelativeUrl = folderPath || rootFolder.ServerRelativeUrl;
    
    let folders: IFolderItem[] = [];
    let hasMore = false;
    
    // If searching, use SharePoint Search API with ParentLink filter
    if (searchQuery) {
      folders = await this.searchFolders(folderServerRelativeUrl, searchQuery, pageSize, skip);
      hasMore = folders.length === pageSize;
    } else {
      // Use REST API for normal folder loading
      const folder = sp.web.getFolderByServerRelativePath(folderServerRelativeUrl);
      
      let query = folder.folders
        .select('Name', 'ServerRelativeUrl')
        .orderBy('Name')
        .top(pageSize);
      
      if (skip > 0) {
        query = query.skip(skip);
      }
      
      const allFolders = await query();
      
      // Filter out hidden folders
      folders = allFolders
        .filter((f: any) => {
          const name = f.Name || '';
          return !name.startsWith('.') && name.toLowerCase() !== 'forms';
        })
        .map((f: any) => ({
          key: f.ServerRelativeUrl,
          name: f.Name,
          serverRelativeUrl: f.ServerRelativeUrl,
          path: f.ServerRelativeUrl
        }));
        
      hasMore = allFolders.length === pageSize;
    }

    return {
      folders,
      hasMore,
      libraryRootUrl: rootFolder.ServerRelativeUrl
    };
  }

  /**
   * Search folders using SharePoint Search API
   */
  private async searchFolders(
    folderServerRelativeUrl: string,
    searchQuery: string,
    pageSize: number,
    skip: number
  ): Promise<IFolderItem[]> {
    const sp = spfi().using(SPFx(this.context));
    
    // Build search query
    const escapedQuery = searchQuery.replace(/"/g, '\\"');
    const searchQueryText = `*${escapedQuery}*`;
    
    // Get web URLs
    const webAbsoluteUrl = this.context.pageContext.web.absoluteUrl;
    const webServerRelativeUrl = this.context.pageContext.web.serverRelativeUrl;
    
    // Build ParentLink filter
    let folderRelativePath = folderServerRelativeUrl;
    if (folderServerRelativeUrl.toLowerCase().startsWith(webServerRelativeUrl.toLowerCase())) {
      folderRelativePath = folderServerRelativeUrl.substring(webServerRelativeUrl.length);
      if (folderRelativePath && !folderRelativePath.startsWith('/')) {
        folderRelativePath = '/' + folderRelativePath;
      }
    }
    const folderFullUrl = `${webAbsoluteUrl}${folderRelativePath}`;
    const parentLinkFilter = `ParentLink:"${folderFullUrl}"`;
    
    // Build search query object
    const searchQueryObj: any = {
      Querytext: `${searchQueryText} AND ${parentLinkFilter} AND ContentType:Folder`,
      RowLimit: pageSize,
      SelectProperties: ['Title', 'Path', 'ParentLink', 'FileRef'],
      TrimDuplicates: true
    };
    
    if (skip > 0) {
      searchQueryObj.StartRow = skip;
    }
    
    const searchResults = await sp.search(searchQueryObj);
    
    // Filter and convert search results
    const searchFolders = (searchResults.PrimarySearchResults || [])
      .filter((result: any) => {
        // Filter out hidden folders
        const path = result.Path || result.FileRef || '';
        const pathParts = path.split('/');
        const folderName = pathParts[pathParts.length - 1];
        if (folderName.startsWith('.') || folderName.toLowerCase() === 'forms') {
          return false;
        }
        
        // Filter to only include folders at the current level
        const resultParentLink = result.ParentLink || '';
        const normalizedResultParentLink = UrlUtils.normalizeFolderUrl(resultParentLink);
        const normalizedCurrentFolderUrl = UrlUtils.normalizeFolderUrl(folderFullUrl);
        
        if (normalizedResultParentLink !== normalizedCurrentFolderUrl) {
          return false;
        }
        
        // Verify path depth
        const resultPath = result.Path || '';
        let resultServerRelativeUrl = resultPath;
        if (resultPath.startsWith('http://') || resultPath.startsWith('https://')) {
          const urlObj = new URL(resultPath);
          resultServerRelativeUrl = urlObj.pathname;
        }
        
        const getPathSegments = (path: string): string[] => {
          return path.split('/').filter((p: string) => {
            return p && 
              p.toLowerCase() !== 'sites' && 
              p.toLowerCase() !== 'teams' && 
              !p.includes('.');
          });
        };
        
        const currentFolderPathSegments = getPathSegments(folderServerRelativeUrl);
        const resultPathSegments = getPathSegments(resultServerRelativeUrl);
        
        return resultPathSegments.length === currentFolderPathSegments.length + 1;
      })
      .map((result: any) => {
        const folderName = result.Title || '';
        const serverRelativeUrl = UrlUtils.buildFolderPath(folderServerRelativeUrl, folderName);
        
        return {
          key: serverRelativeUrl,
          name: folderName,
          serverRelativeUrl: serverRelativeUrl,
          path: serverRelativeUrl
        };
      });
    
    return searchFolders;
  }

  /**
   * Copy a file from source to destination
   */
  async copyFile(sourceServerRelativeUrl: string, destinationServerRelativeUrl: string): Promise<{
    serverRelativeUrl: string;
    uniqueId: string;
  }> {
    const sp = spfi().using(SPFx(this.context));
    
    // Copy file using copyByPath
    const sourceFile = sp.web.getFileByServerRelativePath(sourceServerRelativeUrl);
    await sourceFile.copyByPath(destinationServerRelativeUrl, true, true);
    
    // Get the created file
    const createdFile = sp.web.getFileByServerRelativePath(destinationServerRelativeUrl);
    const fileData = await createdFile.select('ServerRelativeUrl')();
    const listItemData = await createdFile.listItemAllFields.select('UniqueId')();
    
    return {
      serverRelativeUrl: fileData.ServerRelativeUrl,
      uniqueId: listItemData.UniqueId
    };
  }
}

