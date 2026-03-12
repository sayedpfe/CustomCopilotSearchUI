import * as React from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { SearchBox } from './SearchBox';
import { NewsCarousel } from './NewsCarousel';
import styles from './BoschSearchApp.module.scss';

type SearchScope = 'work' | 'web';

export interface ISearchLandingProps {
  searchScope: SearchScope;
  onSearchScopeChange: (scope: SearchScope) => void;
  onSearch: (query: string) => void;
  graphClient: MSGraphClientV3 | undefined;
  newsSourceSiteUrl: string;
  hasCopilot: boolean | null;
  backgroundImageUrl?: string;
}

export const SearchLanding: React.FC<ISearchLandingProps> = ({
  searchScope,
  onSearchScopeChange,
  onSearch,
  graphClient,
  newsSourceSiteUrl,
  hasCopilot,
  backgroundImageUrl,
}) => {
  const containerClass = backgroundImageUrl
    ? `${styles.landingContainer} ${styles.landingContainerWithBackground}`
    : styles.landingContainer;

  return (
    <div
      className={containerClass}
      style={backgroundImageUrl ? { backgroundImage: `url('${backgroundImageUrl}')` } : undefined}
    >
      <div className={styles.landingHero}>
        <div className={styles.boschLogo}>
          <span className={`${styles.boschLogoText}${backgroundImageUrl ? ` ${styles.boschLogoTextLight}` : ''}`}>BOSCH</span>
        </div>

        <SearchBox
          searchScope={searchScope}
          onSearchScopeChange={onSearchScopeChange}
          onSearch={onSearch}
          variant="landing"
          hasCopilot={hasCopilot}
        />
      </div>

      {graphClient && (
        <div className={`${styles.newsSection}${backgroundImageUrl ? ` ${styles.newsSectionOnBackground}` : ''}`}>
          <h2 className={styles.newsSectionTitle}>Bosch News</h2>
          <NewsCarousel
            graphClient={graphClient}
            siteUrl={newsSourceSiteUrl}
          />
        </div>
      )}
    </div>
  );
};
