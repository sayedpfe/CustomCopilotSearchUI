import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './BoschSearchApp.module.scss';

interface INewsItem {
  id: string;
  title: string;
  description: string;
  url: string;
  imageUrl: string;
  publishedDate: string;
  views: number;
  likes: number;
}

export interface INewsCarouselProps {
  graphClient: MSGraphClientV3;
  siteUrl: string;
}

export const NewsCarousel: React.FC<INewsCarouselProps> = ({
  graphClient,
  siteUrl,
}) => {
  const [newsItems, setNewsItems] = useState<INewsItem[]>([]);
  const [loading, setLoading] = useState(true);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    fetchNews();
  }, [graphClient, siteUrl]);

  const fetchNews = async (): Promise<void> => {
    try {
      // Use Graph Search to find SharePoint news pages
      const response = await graphClient
        .api('/search/query')
        .version('v1.0')
        .post({
          requests: [
            {
              entityTypes: ['listItem'],
              query: {
                queryString: 'PromotedState:2 contentclass:STS_ListItem_WebPageLibrary',
              },
              from: 0,
              size: 12,
              fields: [
                'title',
                'description',
                'path',
                'ViewsLifeTime',
                'LikesCount',
                'PictureThumbnailURL',
                'LastModifiedTime',
              ],
            },
          ],
        });

      const hits = response?.value?.[0]?.hitsContainers?.[0]?.hits || [];
      const items: INewsItem[] = hits.map(
        (hit: {
          resource: {
            id: string;
            fields?: Record<string, string>;
          };
          summary?: string;
        }) => {
          const fields = hit.resource?.fields || {};
          return {
            id: hit.resource?.id || '',
            title: fields.title || 'Untitled',
            description: fields.description || hit.summary || '',
            url: fields.path || '#',
            imageUrl: fields.pictureThumbnailURL || fields.PictureThumbnailURL || '',
            publishedDate: fields.lastModifiedTime || fields.LastModifiedTime || '',
            views: parseInt(fields.viewsLifeTime || fields.ViewsLifeTime || '0', 10),
            likes: parseInt(fields.likesCount || fields.LikesCount || '0', 10),
          };
        }
      );

      setNewsItems(items);
    } catch (err) {
      console.error('[NewsCarousel] Error fetching news:', err);
      // Show placeholder items on error
      setNewsItems(getPlaceholderNews());
    } finally {
      setLoading(false);
    }
  };

  const scroll = (direction: 'left' | 'right'): void => {
    if (!scrollRef.current) return;
    const amount = 320;
    scrollRef.current.scrollBy({
      left: direction === 'left' ? -amount : amount,
      behavior: 'smooth',
    });
  };

  if (loading) {
    return (
      <div className={styles.newsCarouselLoading}>
        {Array.from({ length: 6 }).map((_, i) => (
          <div key={i} className={styles.newsCardSkeleton} />
        ))}
      </div>
    );
  }

  if (newsItems.length === 0) return null;

  return (
    <div className={styles.newsCarousel}>
      <button
        className={`${styles.carouselArrow} ${styles.carouselArrowLeft}`}
        onClick={() => scroll('left')}
        aria-label="Scroll left"
      >
        <Icon iconName="ChevronLeft" />
      </button>

      <div ref={scrollRef} className={styles.newsCarouselTrack}>
        {newsItems.map((item) => (
          <a
            key={item.id}
            href={item.url}
            target="_blank"
            rel="noopener noreferrer"
            className={styles.newsCard}
          >
            <div
              className={styles.newsCardImage}
              style={{
                backgroundImage: item.imageUrl
                  ? `url(${item.imageUrl})`
                  : 'linear-gradient(135deg, #0078d4, #00bcf2)',
              }}
            >
              <div className={styles.newsCardImageOverlay}>
                <span className={styles.newsCardCategory}>
                  {item.title.split(' ').slice(0, 3).join(' ')}
                </span>
              </div>
            </div>
            <div className={styles.newsCardBody}>
              <h3 className={styles.newsCardTitle}>{item.title}</h3>
              <p className={styles.newsCardDescription}>
                {item.description.substring(0, 100)}
                {item.description.length > 100 ? '...' : ''}
              </p>
              <div className={styles.newsCardMeta}>
                <span>
                  <Icon iconName="View" /> {item.views.toLocaleString()}
                </span>
                <span>
                  <Icon iconName="Like" /> {item.likes}
                </span>
              </div>
            </div>
          </a>
        ))}
      </div>

      <button
        className={`${styles.carouselArrow} ${styles.carouselArrowRight}`}
        onClick={() => scroll('right')}
        aria-label="Scroll right"
      >
        <Icon iconName="ChevronRight" />
      </button>
    </div>
  );
};

function getPlaceholderNews(): INewsItem[] {
  const topics = [
    'AI transforms automotive industry',
    'Smart home innovation breakthroughs',
    'IoT solutions for manufacturing',
    'Sustainable technology advances',
    'Next-gen mobility solutions',
    'Industry 4.0 revolutionizes production',
    'Connected devices ecosystem grows',
    'Energy efficiency breakthroughs',
    'Digital workplace transformation',
    'Bosch is introducing M365 Copilot',
  ];

  return topics.map((title, i) => ({
    id: `placeholder-${i}`,
    title,
    description: `Discover the latest developments in ${title.toLowerCase()}.`,
    url: '#',
    imageUrl: '',
    publishedDate: new Date().toISOString(),
    views: Math.floor(Math.random() * 5000) + 500,
    likes: Math.floor(Math.random() * 50) + 1,
  }));
}
