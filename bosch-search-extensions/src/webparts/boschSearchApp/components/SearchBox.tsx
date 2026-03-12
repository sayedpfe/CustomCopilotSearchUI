import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './BoschSearchApp.module.scss';

type SearchScope = 'work' | 'web';

export interface ISearchBoxProps {
  initialQuery?: string;
  searchScope: SearchScope;
  onSearchScopeChange: (scope: SearchScope) => void;
  onSearch: (query: string) => void;
  variant: 'landing' | 'results';
  hasCopilot: boolean | null;
}

export const SearchBox: React.FC<ISearchBoxProps> = ({
  initialQuery,
  searchScope,
  onSearchScopeChange,
  onSearch,
  variant,
  hasCopilot,
}) => {
  const [value, setValue] = useState(initialQuery || '');
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (initialQuery !== undefined) {
      setValue(initialQuery);
    }
  }, [initialQuery]);

  const handleSubmit = (e: React.FormEvent): void => {
    e.preventDefault();
    if (value.trim()) {
      onSearch(value.trim());
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'Enter') {
      handleSubmit(e);
    }
  };

  const isLanding = variant === 'landing';

  return (
    <div className={isLanding ? styles.searchBoxLanding : styles.searchBoxResults}>
      {isLanding && (
        <div className={styles.scopeToggle}>
          <button
            className={`${styles.scopeButton} ${searchScope === 'work' ? styles.scopeButtonActive : ''}`}
            onClick={() => onSearchScopeChange('work')}
          >
            Work
          </button>
          <button
            className={`${styles.scopeButton} ${searchScope === 'web' ? styles.scopeButtonActive : ''}`}
            onClick={() => onSearchScopeChange('web')}
          >
            Web
          </button>
        </div>
      )}

      <div className={isLanding ? styles.searchInputWrapperLanding : styles.searchInputWrapperResults}>
        <Icon iconName="Search" className={styles.searchIcon} />
        <input
          ref={inputRef}
          type="text"
          className={styles.searchInput}
          placeholder={`Search with Bosch AI...`}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={handleKeyDown}
          autoFocus={isLanding}
        />
        <div className={styles.searchActions}>
          <button className={styles.searchActionButton} title="Voice search">
            <Icon iconName="Microphone" />
          </button>
          {hasCopilot && (
            <button
              className={styles.copilotButton}
              title="Search with Copilot"
              onClick={() => onSearch(value.trim())}
            >
              <Icon iconName="Robot" />
            </button>
          )}
        </div>
      </div>
    </div>
  );
};
