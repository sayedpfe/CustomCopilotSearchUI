import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useGraphClient } from '../../../hooks/useGraphClient';
import { CopilotDetectionService } from '../../../services/CopilotDetectionService';
import { BackgroundService } from '../../../services/BackgroundService';
import { Header } from './Header';
import { SearchLanding } from './SearchLanding';
import { SearchResults } from './SearchResults';
import { ChatPanel } from './ChatPanel';
import styles from './BoschSearchApp.module.scss';

export interface IBoschSearchAppProps {
  context: WebPartContext;
  groundingMode: 'work' | 'web' | 'both';
  maxRetrievalResults: number;
  showCopilotLink: boolean;
  newsSourceSiteUrl: string;
  promotedResultsListName: string;
  announcementsListName: string;
  analyticsListName: string;
  backgroundEnabled: boolean;
  backgroundLibraryUrl: string;
}

type AppView = 'landing' | 'results';
type SearchScope = 'work' | 'web';
type SearchVertical = 'all' | 'copilot' | 'images' | 'videos' | 'more';

export const BoschSearchApp: React.FC<IBoschSearchAppProps> = (props) => {
  const { context, groundingMode, newsSourceSiteUrl, backgroundEnabled, backgroundLibraryUrl } = props;
  const { graphClient } = useGraphClient(context);
  const [view, setView] = useState<AppView>('landing');
  const [query, setQuery] = useState('');
  const [searchScope, setSearchScope] = useState<SearchScope>(
    groundingMode === 'web' ? 'web' : 'work'
  );
  const [activeVertical, setActiveVertical] = useState<SearchVertical>('all');
  const [hasCopilot, setHasCopilot] = useState<boolean | null>(null);
  const [isChatOpen, setIsChatOpen] = useState(false);
  const [userName, setUserName] = useState('');
  const [userPhoto, setUserPhoto] = useState('');
  const [backgroundImageUrl, setBackgroundImageUrl] = useState<string | undefined>(undefined);

  // Detect Copilot license
  useEffect(() => {
    if (!graphClient) return;
    const detector = new CopilotDetectionService(graphClient);
    detector.hasCopilotLicense().then(setHasCopilot);
  }, [graphClient]);

  // Load user profile
  useEffect(() => {
    if (!graphClient) return;
    graphClient
      .api('/me')
      .select('displayName')
      .get()
      .then((user: { displayName: string }) => {
        setUserName(user.displayName || '');
      })
      .catch(() => { /* ignore */ });

    graphClient
      .api('/me/photo/$value')
      .responseType('blob' as never)
      .get()
      .then((blob: Blob) => {
        setUserPhoto(URL.createObjectURL(blob));
      })
      .catch(() => { /* no photo */ });
  }, [graphClient]);

  // Load daily background image from SharePoint library when enabled
  useEffect(() => {
    if (!graphClient || !backgroundEnabled || !backgroundLibraryUrl) {
      setBackgroundImageUrl(undefined);
      return;
    }
    const svc = new BackgroundService(graphClient);
    svc.getDailyBackgroundImageUrl(backgroundLibraryUrl)
      .then((url) => setBackgroundImageUrl(url || undefined))
      .catch(() => setBackgroundImageUrl(undefined));
  }, [graphClient, backgroundEnabled, backgroundLibraryUrl]);

  const handleSearch = useCallback((searchText: string) => {
    if (!searchText.trim()) return;
    setQuery(searchText.trim());
    setView('results');
  }, []);

  const handleLogoClick = useCallback(() => {
    setView('landing');
    setQuery('');
    setActiveVertical('all');
  }, []);

  const handleVerticalChange = useCallback((vertical: SearchVertical) => {
    setActiveVertical(vertical);
    if (vertical === 'copilot') {
      setIsChatOpen(true);
    }
  }, []);

  return (
    <div className={styles.boschSearchApp}>
      <Header
        appName="Bosch AI Search"
        userName={userName}
        userPhotoUrl={userPhoto}
        activeVertical={activeVertical}
        onVerticalChange={handleVerticalChange}
        onLogoClick={handleLogoClick}
        showVerticals={view === 'results'}
        hasCopilot={hasCopilot}
      />

      <main className={styles.mainContent}>
        {view === 'landing' && (
          <SearchLanding
            searchScope={searchScope}
            onSearchScopeChange={setSearchScope}
            onSearch={handleSearch}
            graphClient={graphClient}
            newsSourceSiteUrl={newsSourceSiteUrl}
            hasCopilot={hasCopilot}
            backgroundImageUrl={backgroundImageUrl}
            context={context}
          />
        )}

        {view === 'results' && graphClient && (
          <SearchResults
            query={query}
            searchScope={searchScope}
            onSearchScopeChange={setSearchScope}
            onSearch={handleSearch}
            graphClient={graphClient}
            hasCopilot={hasCopilot}
            activeVertical={activeVertical}
            {...props}
          />
        )}
      </main>

      {hasCopilot && (
        <ChatPanel
          isOpen={isChatOpen}
          onDismiss={() => setIsChatOpen(false)}
          onOpen={() => setIsChatOpen(true)}
          graphClient={graphClient}
          groundingMode={props.groundingMode}
          hasCopilot={hasCopilot}
        />
      )}
    </div>
  );
};
