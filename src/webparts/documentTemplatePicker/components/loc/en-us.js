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
            Button_Confirm: "Confirm",
            Button_Cancel: "Cancel",

            // Search placeholders
            SearchPlaceholder_Templates: "Search templates...",
            SearchPlaceholder_Folders: "Search folders...",
            
            // File name
            Label_FileName: "File name",
            Placeholder_FileName: "Enter file name...",

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
            Dialog_ConfirmTitle: "Confirm File Name",
            Dialog_ConfirmMessage: "Please confirm or change the file name before creating the document.",

            // Error messages
            Error_CreatingTemplate: "Error creating template: {0}",
            Error_LoadingTemplates: "An error occurred while loading templates",
            Error_FileName_Message1: "Please enter a name that doesn't include any of these",
            Error_FileName_Message2: "characters: \" * : < > ? / \\ |.",
            Error_FileName_EndsWithPeriod: "File or folder names can't end with: .",

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

