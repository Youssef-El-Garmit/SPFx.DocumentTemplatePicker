import * as React from 'react';
import { Stack, Text, MessageBar, MessageBarType, Icon } from '@fluentui/react';
import { format } from '@fluentui/react';
import styles from '../DocumentTemplatePicker.module.scss';
import { FileUtils } from '../utils/FileUtils';
import type { ITemplateItem } from '../IDocumentTemplatePickerProps';
import { DocumentTemplatePicker as strings } from 'ComponentStrings';

export interface IDialogTitleProps {
  template: ITemplateItem;
  message?: string;
  messageType?: MessageBarType;
  onDismissMessage?: () => void;
}

export const DialogTitle: React.FC<IDialogTitleProps> = ({
  template,
  message,
  messageType,
  onDismissMessage
}) => {
  const fileExtension = template.name?.split('.').pop()?.toUpperCase() || '';

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ width: '100%' }}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} style={{ maxWidth: '100%' }}>
        <div 
          className={styles.fileIconContainer}
          style={{ 
            backgroundColor: '#fff7e9',
            color: '#ff7f00'
          }}
        >
          <Icon iconName={FileUtils.getFileIcon(template.name)} className={styles.fileIcon} />
        </div>
        <Stack grow style={{ minWidth: 0 }}>
          <Text variant="mediumPlus" className={styles.fileName} style={{ 
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap'
          }}>
            {template.name}
          </Text>
          <Text variant="small" className={styles.fileType}>
            {format(strings.Dialog_FileType, fileExtension)}
          </Text>
        </Stack>
      </Stack>
      {message && messageType && (
        <MessageBar messageBarType={messageType} onDismiss={onDismissMessage}>
          {message}
        </MessageBar>
      )}
    </Stack>
  );
};

