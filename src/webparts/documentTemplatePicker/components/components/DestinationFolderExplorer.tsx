import * as React from 'react';
import {
  Stack,
  SearchBox,
  Breadcrumb,
  DetailsList,
  IColumn,
  SelectionMode,
  Selection,
  Spinner,
  SpinnerSize,
  Link,
  Icon
} from '@fluentui/react';
import styles from '../DocumentTemplatePicker.module.scss';
import { EmptyState } from './EmptyState';
import type { IFolderItem } from '../IDocumentTemplatePickerProps';
import { IBreadcrumbItem } from '@fluentui/react';
import { DocumentTemplatePicker as strings } from 'ComponentStrings';

export interface IDestinationFolderExplorerProps {
  folders: IFolderItem[];
  breadcrumbItems: IBreadcrumbItem[];
  searchQuery: string;
  loading: boolean;
  loadingMore: boolean;
  hasMore: boolean;
  selection: Selection;
  onSearchChange: (value: string) => void;
  onFolderClick: (folder: IFolderItem) => void;
  onScroll: (ev: React.UIEvent<HTMLDivElement>) => void;
}

export const DestinationFolderExplorer: React.FC<IDestinationFolderExplorerProps> = ({
  folders,
  breadcrumbItems,
  searchQuery,
  loading,
  loadingMore,
  hasMore,
  selection,
  onSearchChange,
  onFolderClick,
  onScroll
}) => {
  const columns: IColumn[] = [
    {
      key: 'icon',
      name: '',
      fieldName: 'icon',
      minWidth: 40,
      maxWidth: 40,
      onRender: (item: IFolderItem) => (
        <Icon iconName="Folder" className={styles.fileIcon} />
      )
    },
    {
      key: 'name',
      name: strings.Column_Name,
      fieldName: 'name',
      minWidth: 200,
      onRender: (item: IFolderItem) => (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Link onClick={() => onFolderClick(item)} className={styles.titleLink}>
            {item.name}
          </Link>
        </Stack>
      )
    }
  ];

  return (
    <Stack className={styles.destinationSection}>
      <SearchBox
        placeholder={strings.SearchPlaceholder_Folders}
        value={searchQuery}
        onChange={(_, newValue) => onSearchChange(newValue || '')}
        className={styles.destinationSearchBox}
      />

      {/* Breadcrumb - show when inside a folder (not at root) */}
      {breadcrumbItems.length > 0 && (
        <Breadcrumb
          items={breadcrumbItems}
          className={styles.breadcrumb}
        />
      )}

      <Stack className={styles.destinationListContainer}>
          {loading && folders.length === 0 ? (
            <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '200px' }}>
              <Spinner size={SpinnerSize.medium} label={strings.Loading_Folders} />
            </Stack>
          ) : folders.length === 0 ? (
            <EmptyState
              iconName={searchQuery ? "Search" : "Folder"}
              title={strings.EmptyState_NoFolders}
              description={searchQuery ? strings.EmptyState_NoFoldersDescription : strings.EmptyState_FolderEmpty}
              inline={true}
            />
          ) : (
            <Stack style={{ flex: 1, minHeight: 0, position: 'relative' }} tokens={{ childrenGap: 0 }}>
              <div 
                onScroll={onScroll}
                className={styles.destinationListScrollContainer}
                style={{ flex: 1, minHeight: 0 }}
              >
                {loading && folders.length > 0 && (
                  <Stack 
                    horizontalAlign="center" 
                    verticalAlign="center"
                    style={{ 
                      position: 'absolute',
                      top: 0,
                      left: 0,
                      right: 0,
                      bottom: 0,
                      backgroundColor: 'rgba(255, 255, 255, 0.8)',
                      zIndex: 10,
                      pointerEvents: 'none'
                    }}
                  >
                    <Spinner size={SpinnerSize.medium} label={strings.Loading_Folders} />
                  </Stack>
                )}
                <DetailsList
                  items={folders}
                  columns={columns}
                  selectionMode={SelectionMode.single}
                  selection={selection}
                  className={styles.detailsList}
                  getKey={(item: IFolderItem) => item.key || item.serverRelativeUrl}
                  onShouldVirtualize={() => false}
                />
              </div>
              {loadingMore && hasMore && (
                <Stack horizontalAlign="center" style={{ padding: '8px', flexShrink: 0 }}>
                  <Spinner size={SpinnerSize.small} label={strings.Loading_More} />
                </Stack>
              )}
            </Stack>
          )}
        </Stack>
    </Stack>
  );
};

