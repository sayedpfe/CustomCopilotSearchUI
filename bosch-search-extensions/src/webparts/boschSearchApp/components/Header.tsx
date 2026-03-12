import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import styles from './BoschSearchApp.module.scss';

type SearchVertical = 'all' | 'copilot' | 'images' | 'videos' | 'more';

export interface IHeaderProps {
  appName: string;
  userName: string;
  userPhotoUrl: string;
  activeVertical: SearchVertical;
  onVerticalChange: (vertical: SearchVertical) => void;
  onLogoClick: () => void;
  showVerticals: boolean;
  hasCopilot: boolean | null;
}

const VERTICALS: Array<{ key: SearchVertical; label: string; icon: string }> = [
  { key: 'copilot', label: 'Copilot', icon: 'Robot' },
  { key: 'images', label: 'Images', icon: 'Photo2' },
  { key: 'videos', label: 'Videos', icon: 'Video' },
  { key: 'more', label: 'More', icon: 'More' },
];

export const Header: React.FC<IHeaderProps> = ({
  appName,
  userName,
  userPhotoUrl,
  activeVertical,
  onVerticalChange,
  onLogoClick,
  showVerticals,
  hasCopilot,
}) => {
  return (
    <header className={styles.header}>
      <div className={styles.headerLeft}>
        <button className={styles.appTitle} onClick={onLogoClick}>
          {appName}
        </button>

        {showVerticals && (
          <nav className={styles.verticalNav}>
            {VERTICALS
              .filter((v) => v.key !== 'copilot' || hasCopilot)
              .map((v) => (
                <button
                  key={v.key}
                  className={`${styles.verticalTab} ${activeVertical === v.key ? styles.verticalTabActive : ''}`}
                  onClick={() => onVerticalChange(v.key)}
                >
                  {v.label}
                </button>
              ))}
          </nav>
        )}
      </div>

      <div className={styles.headerRight}>
        <button className={styles.headerIconButton} title="Settings">
          <Icon iconName="Settings" />
        </button>
        <Persona
          imageUrl={userPhotoUrl}
          text={userName}
          size={PersonaSize.size32}
          hidePersonaDetails
          className={styles.userAvatar}
        />
      </div>
    </header>
  );
};
