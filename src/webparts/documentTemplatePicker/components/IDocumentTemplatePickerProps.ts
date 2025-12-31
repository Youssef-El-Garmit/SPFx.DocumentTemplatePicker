import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDocumentTemplatePickerProps {
  context: WebPartContext;
  templatesLibraryId: string;
  templatesLibraryTitle: string;
  destinationLibraryId: string;
  destinationLibraryTitle: string;
  allowCreateAtRoot: boolean;
  onConfigure: () => void;
}

export interface ITemplateItem {
  key: string;
  name: string;
  fileType: string;
  fileIcon: string;
  fileUrl: string;
  serverRelativeUrl: string;
  uniqueId?: string;
  modified: Date;
  modifiedBy: string;
  size: number;
  isFolder: boolean;
  folderPath: string;
  thumbnailUrl?: string;
  previewUrl?: string;
}

export interface IFolderItem {
  key: string;
  name: string;
  serverRelativeUrl: string;
  path: string;
}
