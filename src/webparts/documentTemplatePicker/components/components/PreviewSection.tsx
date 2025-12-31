import * as React from 'react';
import { Stack, Text, Link, Icon } from '@fluentui/react';
import styles from '../DocumentTemplatePicker.module.scss';
import { FileUtils } from '../utils/FileUtils';
import type { ITemplateItem } from '../IDocumentTemplatePickerProps';
import { DocumentTemplatePicker as strings } from 'ComponentStrings';

export interface IPreviewSectionProps {
  template: ITemplateItem;
  webUrl: string;
}

export const PreviewSection: React.FC<IPreviewSectionProps> = ({
  template,
  webUrl
}) => {
  const previewUrl = template.previewUrl;
  const fileUrl = `${webUrl}${template.serverRelativeUrl}`;

  return (
    <Stack className={styles.previewSection}>
      <div className={styles.previewContainer}>
        {previewUrl ? (
          <iframe
            src={previewUrl}
            className={styles.previewIframe}
            title={template.name}
            allow="autoplay"
          />
        ) : (
          <Stack horizontalAlign="center" verticalAlign="center" className={styles.previewPlaceholder} style={{ width: '100%' }}>
            <Icon iconName={FileUtils.getFileIcon(template.name)} className={styles.previewPlaceholderIcon} />
            <Text variant="large" className={styles.previewPlaceholderText}>
              {template.name}
            </Text>
            <Text variant="small">
              {strings.Preview_NotAvailable}
            </Text>
            <Link href={fileUrl} target="_blank" style={{ marginTop: 16 }}>
              {strings.Preview_DownloadFile}
            </Link>
          </Stack>
        )}
      </div>
    </Stack>
  );
};

