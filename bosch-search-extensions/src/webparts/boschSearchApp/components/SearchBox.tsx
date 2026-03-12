import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SearchSuggestService } from '../../../services/SearchSuggestService';
import styles from './BoschSearchApp.module.scss';

type SearchScope = 'work' | 'web';

export interface ISearchBoxProps {
  initialQuery?: string;
  searchScope: SearchScope;
  onSearchScopeChange: (scope: SearchScope) => void;
  onSearch: (query: string) => void;
  variant: 'landing' | 'results';
  hasCopilot: boolean | null;
  /** SharePoint context — needed for the autocomplete suggest API. Optional: box works without it. */
  context?: WebPartContext;
}

export const SearchBox: React.FC<ISearchBoxProps> = ({
  initialQuery,
  searchScope,
  onSearchScopeChange,
  onSearch,
  variant,
  hasCopilot,
  context,
}) => {
  const [value, setValue] = useState(initialQuery || '');
  const [suggestions, setSuggestions] = useState<string[]>([]);
  const [activeSuggestionIndex, setActiveSuggestionIndex] = useState(-1);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const wrapperRef = useRef<HTMLDivElement>(null);
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const suggestSvcRef = useRef<SearchSuggestService | null>(null);

  // Instantiate suggest service once when context becomes available
  useEffect(() => {
    if (context) suggestSvcRef.current = new SearchSuggestService(context);
  }, [context]);

  // Keep internal value in sync when the parent changes the initial query
  useEffect(() => {
    if (initialQuery !== undefined) setValue(initialQuery);
  }, [initialQuery]);

  // Close the dropdown when the user clicks anywhere outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent): void => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) {
        setShowSuggestions(false);
        setActiveSuggestionIndex(-1);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const closeSuggestions = (): void => {
    setSuggestions([]);
    setShowSuggestions(false);
    setActiveSuggestionIndex(-1);
  };

  const triggerSearch = (term: string): void => {
    const trimmed = term.trim();
    if (!trimmed) return;
    closeSuggestions();
    onSearch(trimmed);
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const newValue = e.target.value;
    setValue(newValue);
    setActiveSuggestionIndex(-1);

    if (debounceRef.current) clearTimeout(debounceRef.current);

    if (!suggestSvcRef.current || newValue.trim().length < 2) {
      closeSuggestions();
      return;
    }

    // Debounce: wait 300 ms after the user stops typing before calling the API
    debounceRef.current = setTimeout(async () => {
      if (!suggestSvcRef.current) return;
      const results = await suggestSvcRef.current.getSuggestions(newValue.trim());
      setSuggestions(results);
      setShowSuggestions(results.length > 0);
    }, 300);
  };

  const handleSubmit = (e: React.FormEvent): void => {
    e.preventDefault();
    triggerSearch(value);
  };

  const handleKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'Escape') {
      closeSuggestions();
      return;
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setActiveSuggestionIndex((i) => Math.min(i + 1, suggestions.length - 1));
      return;
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActiveSuggestionIndex((i) => Math.max(i - 1, -1));
      return;
    }
    if (e.key === 'Enter') {
      // If a suggestion is highlighted, search that, otherwise search the typed text
      if (activeSuggestionIndex >= 0 && suggestions[activeSuggestionIndex]) {
        const chosen = suggestions[activeSuggestionIndex];
        setValue(chosen);
        triggerSearch(chosen);
      } else {
        triggerSearch(value);
      }
    }
  };

  const handleSuggestionMouseDown = (suggestion: string): void => {
    // Use mousedown (not click) so it fires before the input's blur event
    setValue(suggestion);
    triggerSearch(suggestion);
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

      {/* Wrapper gives the dropdown a relative-positioning parent */}
      <div ref={wrapperRef} className={styles.searchBoxWrapper}>
        <div className={isLanding ? styles.searchInputWrapperLanding : styles.searchInputWrapperResults}>
          <Icon iconName="Search" className={styles.searchIcon} />
          <input
            ref={inputRef}
            type="text"
            className={styles.searchInput}
            placeholder="Search with Bosch AI..."
            value={value}
            onChange={handleChange}
            onKeyDown={handleKeyDown}
            autoFocus={isLanding}
            autoComplete="off"
            aria-autocomplete="list"
            aria-haspopup="listbox"
            aria-expanded={showSuggestions}
          />
          <div className={styles.searchActions}>
            <button
              type="button"
              className={styles.searchActionButton}
              title="Search"
              onClick={() => triggerSearch(value)}
            >
              <Icon iconName="Search" />
            </button>
            {hasCopilot && (
              <button
                type="button"
                className={styles.copilotButton}
                title="Search with Copilot"
                onClick={() => triggerSearch(value)}
              >
                <Icon iconName="Robot" />
              </button>
            )}
          </div>
        </div>

        {/* Autocomplete suggestions dropdown */}
        {showSuggestions && suggestions.length > 0 && (
          <ul
            className={styles.suggestionsDropdown}
            role="listbox"
            aria-label="Search suggestions"
          >
            {suggestions.map((suggestion, idx) => (
              <li
                key={idx}
                role="option"
                aria-selected={activeSuggestionIndex === idx}
                className={`${styles.suggestionItem}${activeSuggestionIndex === idx ? ` ${styles.suggestionItemActive}` : ''}`}
                onMouseDown={() => handleSuggestionMouseDown(suggestion)}
                onMouseEnter={() => setActiveSuggestionIndex(idx)}
              >
                <Icon iconName="Search" className={styles.suggestionIcon} />
                {suggestion}
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
};

