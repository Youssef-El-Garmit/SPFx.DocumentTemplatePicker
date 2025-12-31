import * as React from 'react';
import { Stack, Text, PrimaryButton, DefaultButton, Breadcrumb, IBreadcrumbItem } from '@fluentui/react';
import styles from '../DocumentTemplatePicker.module.scss';
import { DocumentTemplatePicker as strings } from 'ComponentStrings';

export interface ISuccessStepProps {
  breadcrumbItems: IBreadcrumbItem[];
  onClose: () => void;
  onOpenDocument: () => void;
}

export const SuccessStep: React.FC<ISuccessStepProps> = ({
  breadcrumbItems,
  onClose,
  onOpenDocument
}) => {
  return (
    <Stack 
      tokens={{ childrenGap: 24 }} 
      style={{ 
        padding: '40px',
        alignItems: 'center',
        justifyContent: 'center',
        minHeight: '400px',
        textAlign: 'center',
        height: '100%',
        display: 'flex',
        flexDirection: 'column'
      }}
    >
      {/* Big Green Checkmark */}
      <Stack
        style={{
          width: '80px',
          height: '80px',
          borderRadius: '50%',
          backgroundColor: '#107c10',
          alignItems: 'center',
          justifyContent: 'center',
          marginBottom: '16px'
        }}
      >
        <Text
          style={{
            fontSize: '48px',
            color: '#ffffff',
            fontWeight: 'bold'
          }}
        >
          ✓
        </Text>
      </Stack>

      {/* Success Message */}
      <Text
        variant="xxLarge"
        style={{
          fontWeight: 600,
          color: '#323130',
          marginBottom: '8px'
        }}
      >
        {strings.Success_Title}
      </Text>

      {/* Information Message */}
      <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: '16px' }}>
        <Text
          variant="large"
          style={{
            color: '#605e5c',
            fontWeight: 500
          }}
        >
            {strings.Success_Message}
        </Text>
      </Stack>

      {/* Breadcrumb - Destination path */}
      {breadcrumbItems.length > 0 && (
        <Breadcrumb
          items={breadcrumbItems}
          className={styles.breadcrumb}
          styles={{ root: { width: '100%', maxWidth: '100%', overflow: 'hidden' } }}
        />
      )}

      {/* Action Buttons */}
      <Stack 
        horizontal 
        tokens={{ childrenGap: 12 }} 
        style={{ marginTop: '32px' }}
      >
        <DefaultButton
          text={strings.Button_Close}
          iconProps={{ iconName: 'Cancel' }}
          onClick={onClose}
          styles={{
            root: {
              minWidth: '140px'
            }
          }}
        />
        <PrimaryButton
          text={strings.Button_OpenDocument}
          iconProps={{ iconName: 'OpenInNewWindow' }}
          onClick={onOpenDocument}
          styles={{
            root: {
              minWidth: '180px'
            }
          }}
        />
      </Stack>
    </Stack>
  );
};

