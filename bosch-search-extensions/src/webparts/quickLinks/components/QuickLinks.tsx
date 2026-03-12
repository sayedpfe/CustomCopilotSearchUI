import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IQuickLink } from '../../../models';
import styles from './QuickLinks.module.scss';

export interface IQuickLinksProps {
  links: IQuickLink[];
  columns: number;
  showDescriptions: boolean;
}

export const QuickLinks: React.FC<IQuickLinksProps> = ({ links, columns, showDescriptions }) => {
  if (!links || links.length === 0) {
    return (
      <div className={styles.emptyState}>
        No quick links configured. Edit the web part to add links.
      </div>
    );
  }

  const gridStyle: React.CSSProperties = {
    gridTemplateColumns: `repeat(${columns}, 1fr)`,
  };

  return (
    <div className={styles.quickLinksGrid} style={gridStyle}>
      {links.map((link, index) => (
        <a
          key={index}
          className={styles.quickLinkTile}
          href={link.url}
          target={link.openInNewTab ? '_blank' : '_self'}
          rel={link.openInNewTab ? 'noopener noreferrer' : undefined}
          title={link.title}
        >
          {link.iconName && (
            <Icon iconName={link.iconName} className={styles.tileIcon} />
          )}
          <span className={styles.tileTitle}>{link.title}</span>
          {showDescriptions && link.description && (
            <span className={styles.tileDescription}>{link.description}</span>
          )}
        </a>
      ))}
    </div>
  );
};
