/* eslint-disable no-undef */
define([], function () {
    return {
        DocumentTemplatePicker: {
            // Column names
            Column_Title: "Titre",
            Column_Preview: "Aperçu",
            Column_Name: "Nom",

            // Buttons
            Button_Close: "Fermer",
            Button_CreateDocument: "Créer un document",
            Button_Creating: "Création en cours...",
            Button_Configure: "Configurer",
            Button_OpenDocument: "Ouvrir le document",

            // Search placeholders
            SearchPlaceholder_Templates: "Rechercher des modèles...",
            SearchPlaceholder_Folders: "Rechercher des dossiers...",

            // Loading messages
            Loading_Templates: "Chargement des modèles...",
            Loading_Folders: "Chargement des dossiers...",
            Loading_More: "Chargement supplémentaire...",

            // Success step
            Success_Title: "Document créé avec succès",
            Success_Message: "Votre document a été créé dans la bibliothèque de destination.",

            // Dialog
            Dialog_CloseAriaLabel: "Fermer",
            Dialog_FileType: "Document {0}",

            // Error messages
            Error_CreatingTemplate: "Erreur lors de la création du modèle : {0}",
            Error_LoadingTemplates: "Une erreur s'est produite lors du chargement des modèles",

            // Empty states - Configuration
            EmptyState_ConfigureTitle: "Configurer cette partie Web",
            EmptyState_ConfigureDescriptionAll: "Sélectionnez la bibliothèque de modèles et la bibliothèque de destination pour commencer",
            EmptyState_ConfigureDescriptionTemplates: "Sélectionnez une bibliothèque de modèles où vos modèles sont stockés",
            EmptyState_ConfigureDescriptionDestination: "Sélectionnez une bibliothèque de destination où les nouveaux éléments seront créés",
            // Empty states - Search/Results
            EmptyState_NoResults: "Aucun résultat trouvé",
            EmptyState_NoResultsDescription: "Essayez d'ajuster votre recherche",
            EmptyState_NoTemplates: "Aucun modèle trouvé",
            EmptyState_NoTemplatesInFolder: "Ce dossier est vide",
            EmptyState_NoTemplatesInLibrary: "La bibliothèque \"{0}\" est vide",
            EmptyState_NoFolders: "Aucun dossier trouvé",
            EmptyState_NoFoldersDescription: "Essayez d'ajuster votre recherche",
            EmptyState_FolderEmpty: "Ce dossier est vide",

            // Preview
            Preview_NotAvailable: "Aperçu non disponible pour ce type de fichier",
            Preview_DownloadFile: "Télécharger le fichier"
        }
    };
});

