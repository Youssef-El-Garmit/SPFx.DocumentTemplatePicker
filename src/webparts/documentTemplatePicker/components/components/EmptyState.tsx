import * as React from 'react';
import { Text, Icon } from '@fluentui/react';
import styles from '../DocumentTemplatePicker.module.scss';

export interface IEmptyStateProps {
  iconName: string;
  title: string;
  description: string;
  inline?: boolean;
}

export const EmptyState: React.FC<IEmptyStateProps> = ({
  iconName,
  title,
  description,
  inline = false
}) => {
  const containerClass = inline ? styles.emptyStateInline : styles.emptyState;
  const contentClass = styles.emptyStateContent;

  return (
    <div className={containerClass}>
      <div className={contentClass}>
        <Icon iconName={iconName} className={styles.emptyStateIcon} />
        <Text variant={inline ? "large" : "xxLarge"} className={styles.emptyStateText}>
          {title}
        </Text>
        <Text variant={inline ? "medium" : "medium"} className={styles.emptyStateDescription}>
          {description}
        </Text>
      </div>
    </div>
  );
};

