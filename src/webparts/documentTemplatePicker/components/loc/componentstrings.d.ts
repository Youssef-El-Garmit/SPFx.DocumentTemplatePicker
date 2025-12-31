declare module 'ComponentStrings' {
    interface IDocumentTemplatePickerStrings {
        // Column names
        Column_Title: string;
        Column_Preview: string;
        Column_Name: string;

        // Buttons
        Button_Close: string;
        Button_CreateDocument: string;
        Button_Creating: string;
        Button_Configure: string;
        Button_OpenDocument: string;

        // Search placeholders
        SearchPlaceholder_Templates: string;
        SearchPlaceholder_Folders: string;

        // Loading messages
        Loading_Templates: string;
        Loading_Folders: string;
        Loading_More: string;

        // Success step
        Success_Title: string;
        Success_Message: string;

        // Dialog
        Dialog_CloseAriaLabel: string;
        Dialog_FileType: string; // "{0} Document"

        // Error messages
        Error_CreatingTemplate: string; // "Error creating template: {0}"
        Error_LoadingTemplates: string;

        // Empty states - Configuration
        EmptyState_ConfigureTitle: string;
        EmptyState_ConfigureDescriptionAll: string;
        EmptyState_ConfigureDescriptionTemplates: string;
        EmptyState_ConfigureDescriptionDestination: string;
        // Empty states - Search/Results
        EmptyState_NoResults: string;
        EmptyState_NoResultsDescription: string;
        EmptyState_NoTemplates: string;
        EmptyState_NoTemplatesInFolder: string;
        EmptyState_NoTemplatesInLibrary: string; // "The library \"{0}\" is empty"
        EmptyState_NoFolders: string;
        EmptyState_NoFoldersDescription: string;
        EmptyState_FolderEmpty: string;

        // Preview
        Preview_NotAvailable: string;
        Preview_DownloadFile: string;
    }

    export const DocumentTemplatePicker: IDocumentTemplatePickerStrings;
}

