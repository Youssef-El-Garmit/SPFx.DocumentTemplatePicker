import * as React from 'react';
import styles from './DocumentTemplatePicker.module.scss';
import type { IDocumentTemplatePickerProps, ITemplateItem, IFolderItem } from './IDocumentTemplatePickerProps';
import { 
  Spinner, 
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon,
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  DetailsList,
  IColumn,
  SelectionMode,
  Selection,
  SearchBox,
  Dialog,
  DialogType,
  DialogFooter,
  IBreadcrumbItem,
  Breadcrumb,
  Link,
  Image,
  ImageFit
} from '@fluentui/react';
import { SharePointService } from './services/SharePointService';
import { FileUtils } from './utils/FileUtils';
import { UrlUtils } from './utils/UrlUtils';
import { BreadcrumbUtils } from './utils/BreadcrumbUtils';
import { SuccessStep } from './components/SuccessStep';
import { DialogTitle } from './components/DialogTitle';
import { PreviewSection } from './components/PreviewSection';
import { DestinationFolderExplorer } from './components/DestinationFolderExplorer';
import { format } from '@fluentui/react';
import { DocumentTemplatePicker as strings } from 'ComponentStrings';

interface IComponentState {
  items: ITemplateItem[];
  folders: IFolderItem[];
  loading: boolean;
  error: string | undefined;
  searchQuery: string;
  currentFolder: string;
  breadcrumbItems: IBreadcrumbItem[];
  libraryRootUrl: string;
  showDialog: boolean;
  selectedTemplate: ITemplateItem | undefined;
  destinationFolders: IFolderItem[];
  selectedDestinationFolder: string;
  copying: boolean;
  destinationSearchQuery: string;
  currentDestinationFolder: string;
  destinationLibraryRootUrl: string;
  destinationBreadcrumbItems: IBreadcrumbItem[];
  dialogMessage: string | undefined;
  dialogMessageType: MessageBarType | undefined;
  showSuccess: boolean;
  createdFileUrl: string | undefined;
  createdFileUniqueId: string | undefined;
  isOfficeDocument: boolean;
  destinationLoading: boolean;
  destinationLoadingMore: boolean;
  destinationHasMore: boolean;
  destinationTotalLoaded: number;
}

export default class DocumentTemplatePicker extends React.Component<IDocumentTemplatePickerProps, IComponentState> {
  private _destinationSelection: Selection;
  private _sharePointService: SharePointService;
  private _searchTimeout: NodeJS.Timeout | undefined;

  constructor(props: IDocumentTemplatePickerProps) {
    super(props);
    
    // Initialize SharePoint service
    this._sharePointService = new SharePointService(props.context);

    this.state = {
      items: [],
      folders: [],
      loading: false,
      error: undefined,
      searchQuery: '',
      currentFolder: '',
      breadcrumbItems: [],
      libraryRootUrl: '',
      showDialog: false,
      selectedTemplate: undefined,
      destinationFolders: [],
      selectedDestinationFolder: '',
      copying: false,
      destinationSearchQuery: '',
      currentDestinationFolder: '',
      destinationLibraryRootUrl: '',
      destinationBreadcrumbItems: [],
      dialogMessage: undefined,
      dialogMessageType: undefined,
      showSuccess: false,
      createdFileUrl: undefined,
      createdFileUniqueId: undefined,
      isOfficeDocument: false,
      destinationLoading: false,
      destinationLoadingMore: false,
      destinationHasMore: false,
      destinationTotalLoaded: 0
    };

    // Initialize selection for destination folders
    this._destinationSelection = new Selection({
      onSelectionChanged: () => {
        const selection = this._destinationSelection.getSelection();
        if (selection.length > 0) {
          const selectedFolder = selection[0] as IFolderItem;
          this.setState({ selectedDestinationFolder: selectedFolder.serverRelativeUrl });
        } else {
          this.setState({ selectedDestinationFolder: '' });
        }
      }
    });
  }

  public componentDidMount(): void {
    if (this.props.templatesLibraryId && this._isValidGuid(this.props.templatesLibraryId)) {
      void this._loadItems();
    }
  }

  public componentDidUpdate(prevProps: IDocumentTemplatePickerProps): void {
    if (prevProps.templatesLibraryId !== this.props.templatesLibraryId || 
        prevProps.destinationLibraryId !== this.props.destinationLibraryId) {
      // Only load if we have a valid GUID
      if (this.props.templatesLibraryId && this._isValidGuid(this.props.templatesLibraryId)) {
        void this._loadItems();
      } else {
        // Clear items if GUID is invalid
        this.setState({ items: [], folders: [], loading: false, error: undefined });
      }
    }
  }

  private _isValidGuid(guid: string): boolean {
    if (!guid || typeof guid !== 'string') {
      return false;
    }
    // GUID format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    return guidRegex.test(guid.trim());
  }

  private async _loadItems(folderPath: string = ''): Promise<void> {
    if (!this.props.templatesLibraryId || !this._isValidGuid(this.props.templatesLibraryId)) {
      this.setState({ items: [], folders: [], loading: false, error: undefined });
      return;
    }

    this.setState({ loading: true, error: undefined });

    try {
      const result = await this._sharePointService.loadTemplates(
        this.props.templatesLibraryId, 
        folderPath,
        this.props.templatesLibraryWebUrl
      );
      const folderServerRelativeUrl = folderPath || result.libraryRootUrl;
      
      const breadcrumbItems = BreadcrumbUtils.buildBreadcrumb({
        folderPath: folderServerRelativeUrl,
        libraryRootUrl: result.libraryRootUrl,
        libraryTitle: this.props.templatesLibraryTitle || 'Templates',
        onRootClick: () => this._navigateToRoot(),
        onFolderClick: (path: string) => this._navigateToFolderPath(path)
      });

      this.setState({ 
        items: result.items, 
        folders: result.folders,
        loading: false, 
        error: undefined,
        currentFolder: folderServerRelativeUrl,
        breadcrumbItems,
        libraryRootUrl: result.libraryRootUrl
      });
    } catch (error) {
      console.error('Error loading items:', error);
      this.setState({ 
        items: [], 
        folders: [],
        loading: false, 
        error: error instanceof Error ? error.message : strings.Error_LoadingTemplates 
      });
    }
  }


  private _getColumns(): IColumn[] {
    const columns: IColumn[] = [
      {
        key: 'icon',
        name: '',
        fieldName: 'fileIcon',
        minWidth: 50,
        maxWidth: 50,
        onRender: (item: ITemplateItem) => (
          <Icon iconName={FileUtils.getFileIcon(item.name)} className={styles.fileIcon} />
        )
      },
      {
        key: 'name',
        name: strings.Column_Title,
        fieldName: 'name',
        minWidth: 1,
        flexGrow: 1,
        onRender: (item: ITemplateItem) => (
          <Stack>
            <Link onClick={() => item.isFolder ? this._navigateToFolder(item) : this._showPreview(item)} className={styles.titleLink}>
              {item.name}
            </Link>
          </Stack>
        )
      }
    ];

    // Add Preview column only if enabled
    if (this.props.showPreviewColumn) {
      columns.push({
        key: 'preview',
        name: strings.Column_Preview,
        fieldName: 'preview',
        minWidth: 150,
        maxWidth: 200,
        onRender: (item: ITemplateItem) => {
          if (item.isFolder) return <></>;
          
          return (
            <div className={styles.previewThumbnail} onClick={() => this._showPreview(item)}>
              {item.thumbnailUrl ? (
                <Image 
                  src={item.thumbnailUrl} 
                  alt={item.name}
                  imageFit={ImageFit.cover}
                  className={styles.thumbnailImage}
                  onError={(e: any) => {
                    // Fallback to icon if thumbnail fails
                    e.target.style.display = 'none';
                    e.target.parentElement.querySelector(`.${styles.thumbnailPlaceholder}`).style.display = 'flex';
                  }}
                />
              ) : null}
              <div className={styles.thumbnailPlaceholder} style={{ display: item.thumbnailUrl ? 'none' : 'flex' }}>
                <Icon iconName={item.fileIcon} className={styles.thumbnailIcon} />
              </div>
            </div>
          );
        }
      });
    }

    return columns;
  }

  private _navigateToRoot = (): void => {
    void this._loadItems('');
  }

  private _navigateToFolderPath = (folderPath: string): void => {
    // Navigate to the clicked folder path
    // Ensure we use the exact path from breadcrumb
    void this._loadItems(folderPath);
  }

  private _navigateToDestinationRoot = (): void => {
    void this._loadDestinationFolders('');
  }

  private _navigateToDestinationFolderPath = (folderPath: string): void => {
    // Navigate to the clicked destination folder path
    void this._loadDestinationFolders(folderPath);
  }

  private _navigateToFolder = (item: ITemplateItem): void => {
    if (item.isFolder) {
      void this._loadItems(item.serverRelativeUrl);
    }
  }



  private _onSearchChange = (newValue: string): void => {
    this.setState({ searchQuery: newValue });
  }

  private _getFilteredItems(): ITemplateItem[] {
    const { items, searchQuery } = this.state;
    if (!searchQuery) return items;
    
    const query = searchQuery.toLowerCase();
    return items.filter(item => 
      item.name.toLowerCase().indexOf(query) !== -1
    );
  }

  private _showPreview = (item: ITemplateItem): void => {
    this.setState({ 
      selectedTemplate: item, 
      showDialog: true, 
      destinationSearchQuery: '', 
      currentDestinationFolder: '',
      selectedDestinationFolder: '',
      destinationFolders: [],
      destinationLoading: false,
      destinationLoadingMore: false,
      destinationHasMore: false,
      destinationTotalLoaded: 0,
      dialogMessage: undefined,
      dialogMessageType: undefined
    });
    this._destinationSelection.setAllSelected(false);
    void this._loadDestinationFolders('', false, '');
  }

  private _closeDialog = (): void => {
    this.setState({ 
      showDialog: false, 
      selectedTemplate: undefined,
      destinationFolders: [],
      selectedDestinationFolder: '',
      destinationSearchQuery: '',
      currentDestinationFolder: '',
      destinationLibraryRootUrl: '',
      destinationBreadcrumbItems: [],
      destinationTotalLoaded: 0,
      dialogMessage: undefined,
      dialogMessageType: undefined,
      showSuccess: false,
      createdFileUrl: undefined,
      createdFileUniqueId: undefined,
      isOfficeDocument: false
    });
  }

  private _openCreatedFile = (): void => {
    const { createdFileUrl, isOfficeDocument, selectedTemplate } = this.state;
    if (!createdFileUrl) {
      return;
    }

    // For Office documents, use Office protocol to open in desktop app
    if (isOfficeDocument && selectedTemplate) {
      const officeProtocol = FileUtils.getOfficeProtocol(selectedTemplate.name);
      if (officeProtocol) {
        try {
          window.location.href = `${officeProtocol}${createdFileUrl}`;
          return;
        } catch (error) {
          console.warn('Failed to open with Office protocol, falling back to direct URL:', error);
        }
      }
    }
    
    // For non-Office documents or if Office protocol fails, use direct URL
    window.open(createdFileUrl, '_blank');
  }


  private async _loadDestinationFolders(folderPath: string = '', append: boolean = false, searchQuery: string = ''): Promise<void> {
    if (!this.props.destinationLibraryId || !this._isValidGuid(this.props.destinationLibraryId)) {
      this.setState({ destinationFolders: [], currentDestinationFolder: '', destinationLoading: false, destinationLoadingMore: false, destinationTotalLoaded: 0 });
      return;
    }

    // Set loading state appropriately
    if (append) {
      this.setState({ destinationLoadingMore: true });
    } else {
      this.setState({ destinationLoading: true, destinationLoadingMore: false, destinationTotalLoaded: 0 });
    }

    try {
      const pageSize = 50;
      const skip = append ? this.state.destinationTotalLoaded : 0;
      
      const result = await this._sharePointService.loadDestinationFolders(
        this.props.destinationLibraryId,
        folderPath,
        pageSize,
        skip,
        searchQuery,
        this.props.destinationLibraryWebUrl
      );
      
      const folderServerRelativeUrl = folderPath || result.libraryRootUrl;

      // Build breadcrumb for destination (only if not appending)
      let destinationBreadcrumbItems = this.state.destinationBreadcrumbItems;
      if (!append) {
        destinationBreadcrumbItems = BreadcrumbUtils.buildBreadcrumb({
          folderPath: folderServerRelativeUrl,
          libraryRootUrl: result.libraryRootUrl,
          libraryTitle: this.props.destinationLibraryTitle || 'Destination',
          onRootClick: () => this._navigateToDestinationRoot(),
          onFolderClick: (path: string) => this._navigateToDestinationFolderPath(path)
        });
      }

      // Combine folders: append if loading more, replace if new load
      const updatedFolders = append ? [...this.state.destinationFolders, ...result.folders] : [...result.folders];
      const newTotalLoaded = append ? this.state.destinationTotalLoaded + result.folders.length : result.folders.length;

      // Update selection BEFORE state update to ensure it's ready
      if (!append) {
        this._destinationSelection.setAllSelected(false);
        this._destinationSelection.setItems(updatedFolders, true);
      } else {
        this._destinationSelection.setItems(updatedFolders, true);
        if (this.state.selectedDestinationFolder) {
          const selectedIndex = updatedFolders.findIndex(f => f.serverRelativeUrl === this.state.selectedDestinationFolder);
          if (selectedIndex >= 0) {
            this._destinationSelection.setIndexSelected(selectedIndex, true, false);
          }
        }
      }

      this.setState({ 
        destinationFolders: updatedFolders,
        currentDestinationFolder: folderServerRelativeUrl,
        destinationLibraryRootUrl: result.libraryRootUrl,
        destinationBreadcrumbItems,
        destinationLoading: false,
        destinationLoadingMore: false,
        destinationHasMore: result.hasMore,
        destinationTotalLoaded: newTotalLoaded,
        selectedDestinationFolder: append ? this.state.selectedDestinationFolder : ''
      });
    } catch (error) {
      console.error('Error loading destination folders:', error);
      this.setState({ 
        destinationFolders: append ? this.state.destinationFolders : [],
        currentDestinationFolder: append ? this.state.currentDestinationFolder : '',
        destinationLoading: false,
        destinationLoadingMore: false,
        destinationTotalLoaded: append ? this.state.destinationTotalLoaded : 0
      });
    }
  }

  private _onDestinationSearchChange = (newValue: string): void => {
    this.setState({ destinationSearchQuery: newValue });
    // Debounce search - reload folders with search query after 500ms
    if (this._searchTimeout) {
      clearTimeout(this._searchTimeout);
    }
    this._searchTimeout = setTimeout(() => {
      void this._loadDestinationFolders(this.state.currentDestinationFolder, false, newValue || '');
    }, 500);
  }

  private _onDestinationFolderClick = (folder: IFolderItem): void => {
    // Navigate into folder
    void this._loadDestinationFolders(folder.serverRelativeUrl, false, '');
    this.setState({ selectedDestinationFolder: '', destinationSearchQuery: '' });
  }

  private _onDestinationListScroll = (ev: React.UIEvent<HTMLDivElement>): void => {
    const { destinationLoading, destinationLoadingMore, destinationHasMore } = this.state;
    if (destinationLoading || destinationLoadingMore || !destinationHasMore) return;

    const element = ev.currentTarget;
    const scrollTop = element.scrollTop;
    const scrollHeight = element.scrollHeight;
    const clientHeight = element.clientHeight;

    // Load more when scrolled to 80% of the list
    if (scrollTop + clientHeight >= scrollHeight * 0.8) {
      void this._loadDestinationFolders(this.state.currentDestinationFolder, true, this.state.destinationSearchQuery);
    }
  }

  private async _onCreateFromTemplate(): Promise<void> {
    const { selectedTemplate, selectedDestinationFolder, currentDestinationFolder, destinationLibraryRootUrl } = this.state;
    if (!selectedTemplate || !this.props.destinationLibraryId) {
      return;
    }

    // Determine destination folder path
    const destinationFolderPath = selectedDestinationFolder || currentDestinationFolder || destinationLibraryRootUrl;
    const destinationUrl = UrlUtils.buildFolderPath(destinationFolderPath, selectedTemplate.name);

    this.setState({ copying: true, dialogMessage: undefined, dialogMessageType: undefined });

    try {
      const result = await this._sharePointService.copyFile(
        selectedTemplate.serverRelativeUrl, 
        destinationUrl,
        this.props.templatesLibraryWebUrl,
        this.props.destinationLibraryWebUrl
      );
      const webUrl = this.props.destinationLibraryWebUrl || this.props.context.pageContext.web.absoluteUrl;
      const fileUrl = UrlUtils.buildFullUrl(result.serverRelativeUrl, webUrl);
      const isOfficeDocument = FileUtils.isOfficeDocument(selectedTemplate.name);

      // Build success breadcrumb
      const successBreadcrumbItems = BreadcrumbUtils.buildSuccessBreadcrumb(
        destinationFolderPath,
        destinationLibraryRootUrl,
        this.props.destinationLibraryTitle || 'Destination',
        webUrl
      );

      this.setState({ 
        copying: false, 
        showSuccess: true,
        createdFileUrl: fileUrl,
        createdFileUniqueId: result.uniqueId,
        isOfficeDocument: isOfficeDocument,
        destinationBreadcrumbItems: successBreadcrumbItems
      });
      
      // Reload items to refresh the list
      void this._loadItems(this.state.currentFolder);
    } catch (error) {
      console.error('Error creating template:', error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.setState({ 
        copying: false,
        dialogMessage: format(strings.Error_CreatingTemplate, errorMessage),
        dialogMessageType: MessageBarType.error
      });
    }
  }

  private _getDialogContent(): React.ReactElement {
    const { selectedTemplate, destinationBreadcrumbItems, showSuccess, destinationFolders, destinationLoading, destinationLoadingMore, destinationHasMore, destinationSearchQuery } = this.state;
    if (!selectedTemplate) return <></>;

    const webUrl = this.props.context.pageContext.web.absoluteUrl;

    return (
      <Stack horizontal className={styles.dialogContent} tokens={{ childrenGap: 0 }}>
        {/* Left side - Preview */}
        <PreviewSection template={selectedTemplate} webUrl={webUrl} />

        {/* Right side - Success or Folder Explorer */}
        <Stack className={styles.destinationSection}>
          {showSuccess ? (
            <SuccessStep
              breadcrumbItems={destinationBreadcrumbItems}
              onClose={this._closeDialog}
              onOpenDocument={this._openCreatedFile}
            />
          ) : (
            <DestinationFolderExplorer
              folders={destinationFolders}
              breadcrumbItems={destinationBreadcrumbItems}
              searchQuery={destinationSearchQuery}
              loading={destinationLoading}
              loadingMore={destinationLoadingMore}
              hasMore={destinationHasMore}
              selection={this._destinationSelection}
              onSearchChange={this._onDestinationSearchChange}
              onFolderClick={this._onDestinationFolderClick}
              onScroll={this._onDestinationListScroll}
            />
          )}
        </Stack>
      </Stack>
    );
  }

  private _getDialogFooter(): React.ReactElement | null {
    const { showSuccess, copying, selectedDestinationFolder, currentDestinationFolder, destinationLibraryRootUrl } = this.state;
    
    if (showSuccess) {
      return null; // Success step has its own buttons
    }

    // Check if we're at root
    const isAtRoot = !currentDestinationFolder || currentDestinationFolder === destinationLibraryRootUrl;
    
    // Button is disabled if:
    // 1. Currently copying
    // 2. We're at root, no folder selected, and allowCreateAtRoot is false
    const isCreateDisabled = copying || (isAtRoot && !selectedDestinationFolder && !this.props.allowCreateAtRoot);

    return (
      <DialogFooter>
        {copying && <Spinner size={SpinnerSize.small} />}
        <DefaultButton
          text={strings.Button_Close}
          iconProps={{ iconName: 'Cancel' }}
          onClick={this._closeDialog}
          disabled={copying}
        />
        <PrimaryButton
          text={copying ? strings.Button_Creating : strings.Button_CreateDocument}
          iconProps={{ iconName: copying ? undefined : 'Add' }}
          onClick={() => { void this._onCreateFromTemplate(); }}
          disabled={isCreateDisabled}
        />
      </DialogFooter>
    );
  }

  public render(): React.ReactElement<IDocumentTemplatePickerProps> {
    const {
      templatesLibraryId, 
      templatesLibraryTitle, 
      destinationLibraryId
    } = this.props;
    
    const { 
      loading, 
      error, 
      searchQuery,
      breadcrumbItems,
      showDialog
    } = this.state;

    // Empty state - no configuration or invalid GUIDs
    if (!templatesLibraryId || !this._isValidGuid(templatesLibraryId) || 
        !destinationLibraryId || !this._isValidGuid(destinationLibraryId)) {
      return (
        <section className={styles.documentTemplatePicker}>
          <div className={styles.emptyState}>
            <div className={styles.emptyStateContent}>
              <Icon iconName="DocLibrary" className={styles.emptyStateIcon} />
              <Text variant="xxLarge" className={styles.emptyStateText}>
                {strings.EmptyState_ConfigureTitle}
              </Text>
              <Text variant="medium" className={styles.emptyStateDescription}>
                {!templatesLibraryId && !destinationLibraryId 
                  ? strings.EmptyState_ConfigureDescriptionAll
                  : !templatesLibraryId 
                    ? strings.EmptyState_ConfigureDescriptionTemplates
                    : strings.EmptyState_ConfigureDescriptionDestination}
              </Text>
              <PrimaryButton
                text={strings.Button_Configure}
                iconProps={{ iconName: 'Settings' }}
                onClick={this.props.onConfigure}
                className={styles.configureButton}
              />
            </div>
          </div>
        </section>
      );
    }

    const filteredItems = this._getFilteredItems();

    return (
      <section className={styles.documentTemplatePicker}>
        <Stack tokens={{ childrenGap: 12 }}>
          {/* Web Part Title */}
          {this.props.webPartTitle && (
            <Text variant="xxLarge" className={styles.webPartTitle}>
              {this.props.webPartTitle}
            </Text>
          )}
          
          {/* Search */}
          <SearchBox
            placeholder={strings.SearchPlaceholder_Templates}
            value={searchQuery}
            onChange={(_, newValue) => this._onSearchChange(newValue || '')}
            className={styles.searchBox}
          />

          {/* Breadcrumb - show when inside a folder (not at root) */}
          {breadcrumbItems.length > 0 && (
            <Breadcrumb
              items={breadcrumbItems}
              className={styles.breadcrumb}
            />
          )}

          {/* Error Message */}
          {error && (
            <MessageBar messageBarType={MessageBarType.error}>
              {error}
            </MessageBar>
          )}

          {/* Loading */}
          {loading ? (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.large} label={strings.Loading_Templates} />
            </div>
          ) : filteredItems.length === 0 ? (
            /* Empty state - replaces DetailsList only */
            <div className={styles.emptyStateInline}>
              <div className={styles.emptyStateContent}>
                {searchQuery ? (
                  <>
                    <Icon iconName="Search" className={styles.emptyStateIcon} />
                    <Text variant="large" className={styles.emptyStateText}>
                      {strings.EmptyState_NoResults}
                    </Text>
                    <Text variant="medium" className={styles.emptyStateDescription}>
                      {strings.EmptyState_NoResultsDescription}
                    </Text>
                  </>
                ) : (
                  <>
                    <Icon iconName="Document" className={styles.emptyStateIcon} />
                    <Text variant="large" className={styles.emptyStateText}>
                      {strings.EmptyState_NoTemplates}
                    </Text>
                    <Text variant="medium" className={styles.emptyStateDescription}>
                      {breadcrumbItems.length > 0 
                        ? strings.EmptyState_NoTemplatesInFolder
                        : templatesLibraryTitle 
                          ? format(strings.EmptyState_NoTemplatesInLibrary, templatesLibraryTitle)
                          : strings.EmptyState_NoTemplatesInFolder}
                    </Text>
                  </>
                )}
        </div>
        </div>
          ) : (
            /* DetailsList */
            <DetailsList
              items={filteredItems}
              columns={this._getColumns()}
              selectionMode={SelectionMode.none}
              className={styles.detailsList}
            />
          )}
        </Stack>

        {/* Preview & Destination Dialog */}
        <Dialog
          hidden={!showDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: this.state.selectedTemplate ? (
              <DialogTitle
                template={this.state.selectedTemplate}
                message={this.state.dialogMessage}
                messageType={this.state.dialogMessageType}
                onDismissMessage={() => this.setState({ dialogMessage: undefined, dialogMessageType: undefined })}
              />
            ) : undefined,
            closeButtonAriaLabel: strings.Dialog_CloseAriaLabel
          }}
          modalProps={{
            isBlocking: false,
            styles: {
              main: {
                maxWidth: '90vw',
                width: '90vw',
                minWidth: '1200px',
                maxHeight: '90vh'
              }
            }
          }}
          className="templateDialog"
        >
          {this._getDialogContent()}
          {this._getDialogFooter()}
        </Dialog>
      </section>
    );
  }
}
