import * as React from 'react';
import {
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  Stack,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  Text
} from '@fluentui/react';
import styles from '../DocumentTemplatePicker.module.scss';
import { FileUtils } from '../utils/FileUtils';
import { DocumentTemplatePicker as strings } from 'ComponentStrings';

export interface IFileNameConfirmationDialogProps {
  hidden: boolean;
  templateName: string;
  fileName: string;
  copying: boolean;
  onFileNameChange: (fileName: string) => void;
  onConfirm: () => void;
  onCancel: () => void;
}

export const FileNameConfirmationDialog: React.FC<IFileNameConfirmationDialogProps> = ({
  hidden,
  templateName,
  fileName,
  copying,
  onFileNameChange,
  onConfirm,
  onCancel
}) => {
  const fileExtension = FileUtils.getFileExtension(templateName);
  
  // Extract name without extension for display in input
  let fileNameWithoutExtension: string;
  if (fileExtension && fileName.toLowerCase().endsWith(`.${fileExtension.toLowerCase()}`)) {
    fileNameWithoutExtension = fileName.substring(0, fileName.length - fileExtension.length - 1);
  } else {
    // If no extension or different extension, extract name without any extension
    const lastDotIndex = fileName.lastIndexOf('.');
    fileNameWithoutExtension = lastDotIndex > 0 ? fileName.substring(0, lastDotIndex) : fileName;
  }

  // Disallowed characters in SharePoint filenames
  const disallowedChars = /["*:<>?/\\|]/;
  
  // Check if filename contains disallowed characters
  const hasDisallowedChars = disallowedChars.test(fileNameWithoutExtension);
  
  // Check if filename ends with a period (before extension)
  const endsWithPeriod = fileNameWithoutExtension.endsWith('.');

  const handleNameChange = (newValue: string | undefined): void => {
    // Allow free text including spaces anywhere in the filename
    // Preserve the exact input as the user types it (no trimming during input)
    const newName = newValue || '';
    
    // Rebuild full filename with the original extension from template
    // Only add extension if there's at least one non-whitespace character
    const finalFileName = fileExtension && newName.trim().length > 0 
      ? `${newName}.${fileExtension}` 
      : newName;
    onFileNameChange(finalFileName);
  };

  // Validate: filename should have at least one non-whitespace character, no disallowed characters, and not end with period
  const isFileNameValid = fileNameWithoutExtension.trim().length > 0 && !hasDisallowedChars && !endsWithPeriod;
  
  // Build error message - prioritize "ends with period" error over disallowed chars
  let errorMessage: React.ReactElement | undefined;
  if (endsWithPeriod) {
    errorMessage = (
      <Text variant="small" style={{ color: '#a4262c' }}>
        {strings.Error_FileName_EndsWithPeriod}
      </Text>
    );
  } else if (hasDisallowedChars) {
    errorMessage = (
      <Stack tokens={{ childrenGap: 0 }}>
        <Text variant="small" style={{ color: '#a4262c' }}>{strings.Error_FileName_Message1}</Text>
        <Text variant="small" style={{ color: '#a4262c' }}>{strings.Error_FileName_Message2}</Text>
      </Stack>
    );
  }

  return (
    <Dialog
      hidden={hidden}
      onDismiss={onCancel}
      dialogContentProps={{
        type: DialogType.normal,
        title: strings.Dialog_ConfirmTitle,
        subText: strings.Dialog_ConfirmMessage
      }}
      modalProps={{
        isBlocking: true
      }}
      className={styles.confirmationDialog}
    >
      <Stack tokens={{ childrenGap: 16 }}>
        <TextField
          label={strings.Label_FileName}
          value={fileNameWithoutExtension}
          onChange={(_, newValue) => handleNameChange(newValue)}
          placeholder={strings.Placeholder_FileName}
          disabled={copying}
          suffix={fileExtension ? `.${fileExtension}` : undefined}
          errorMessage={errorMessage}
        />
      </Stack>
      <DialogFooter>
        {copying && <Spinner size={SpinnerSize.small} />}
        <DefaultButton
          text={strings.Button_Cancel}
          onClick={onCancel}
          disabled={copying}
        />
        <PrimaryButton
          text={copying ? strings.Button_Creating : strings.Button_Confirm}
          onClick={onConfirm}
          disabled={!isFileNameValid || copying}
        />
      </DialogFooter>
    </Dialog>
  );
};
