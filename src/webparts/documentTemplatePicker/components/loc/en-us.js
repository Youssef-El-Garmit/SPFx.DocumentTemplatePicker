/* eslint-disable no-undef */
define([], function () {
    return {
        DocumentTemplatePicker: {
            // Column names
            Column_Title: "Title",
            Column_Preview: "Preview",
            Column_Name: "Name",

            // Buttons
            Button_Close: "Close",
            Button_CreateDocument: "Create Document",
            Button_Creating: "Creating...",
            Button_Configure: "Configure",
            Button_OpenDocument: "Open Document",

            // Search placeholders
            SearchPlaceholder_Templates: "Search templates...",
            SearchPlaceholder_Folders: "Search folders...",

            // Loading messages
            Loading_Templates: "Loading templates...",
            Loading_Folders: "Loading folders...",
            Loading_More: "Loading more...",

            // Success step
            Success_Title: "Document created successfully",
            Success_Message: "Your document has been created in the destination library.",

            // Dialog
            Dialog_CloseAriaLabel: "Close",
            Dialog_FileType: "{0} Document",

            // Error messages
            Error_CreatingTemplate: "Error creating template: {0}",
            Error_LoadingTemplates: "An error occurred while loading templates",

            // Empty states - Configuration
            EmptyState_ConfigureTitle: "Configure this web part",
            EmptyState_ConfigureDescriptionAll: "Select templates library and destination library to get started",
            EmptyState_ConfigureDescriptionTemplates: "Select a templates library where your templates are stored",
            EmptyState_ConfigureDescriptionDestination: "Select a destination library where new items will be created",
            // Empty states - Search/Results
            EmptyState_NoResults: "No results found",
            EmptyState_NoResultsDescription: "Try adjusting your search query",
            EmptyState_NoTemplates: "No templates found",
            EmptyState_NoTemplatesInFolder: "This folder is empty",
            EmptyState_NoTemplatesInLibrary: "The library \"{0}\" is empty",
            EmptyState_NoFolders: "No folders found",
            EmptyState_NoFoldersDescription: "Try adjusting your search query",
            EmptyState_FolderEmpty: "This folder is empty",

            // Preview
            Preview_NotAvailable: "Preview not available for this file type",
            Preview_DownloadFile: "Download file"
        }
    };
});

