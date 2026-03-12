import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointListService } from '../../../services/SharePointListService';
import { IAnnouncement, AnnouncementSeverity } from '../../../models';
import styles from './AnnouncementsBanner.module.scss';

export interface IAnnouncementsBannerProps {
  context: WebPartContext;
  listName: string;
  maxItems: number;
  allowDismiss: boolean;
}

const DISMISSED_KEY = 'bosch_search_dismissed_announcements';

function getDismissedIds(): number[] {
  try {
    const stored = localStorage.getItem(DISMISSED_KEY);
    return stored ? JSON.parse(stored) : [];
  } catch {
    return [];
  }
}

function addDismissedId(id: number): void {
  const dismissed = getDismissedIds();
  if (!dismissed.includes(id)) {
    dismissed.push(id);
    localStorage.setItem(DISMISSED_KEY, JSON.stringify(dismissed));
  }
}

function severityToMessageBarType(severity: AnnouncementSeverity): MessageBarType {
  switch (severity) {
    case 'Warning': return MessageBarType.warning;
    case 'Error': return MessageBarType.error;
    case 'Success': return MessageBarType.success;
    case 'Info':
    default: return MessageBarType.info;
  }
}

export const AnnouncementsBanner: React.FC<IAnnouncementsBannerProps> = ({
  context,
  listName,
  maxItems,
  allowDismiss,
}) => {
  const [announcements, setAnnouncements] = useState<IAnnouncement[]>([]);
  const [dismissedIds, setDismissedIds] = useState<number[]>(getDismissedIds());

  useEffect(() => {
    const service = new SharePointListService(context);
    const now = new Date().toISOString();
    const filter = `IsActive eq 1 and StartDate le datetime'${now}' and EndDate ge datetime'${now}'`;

    service
      .getListItems<IAnnouncement>(
        listName,
        filter,
        ['Id', 'Title', 'Message', 'Severity', 'StartDate', 'EndDate', 'IsActive', 'TargetAudience', 'SortOrder'],
        'SortOrder asc',
        maxItems
      )
      .then((items) => setAnnouncements(items))
      .catch((err) => console.error('[AnnouncementsBanner] Error:', err));
  }, [context, listName, maxItems]);

  const handleDismiss = useCallback((id: number) => {
    addDismissedId(id);
    setDismissedIds((prev) => [...prev, id]);
  }, []);

  const visibleAnnouncements = announcements.filter((a) => !dismissedIds.includes(a.id));

  if (visibleAnnouncements.length === 0) {
    return null;
  }

  return (
    <div className={styles.announcementsContainer}>
      {visibleAnnouncements.map((announcement) => (
        <MessageBar
          key={announcement.id}
          messageBarType={severityToMessageBarType(announcement.severity)}
          isMultiline={false}
          onDismiss={allowDismiss ? () => handleDismiss(announcement.id) : undefined}
          dismissButtonAriaLabel="Dismiss"
        >
          <strong>{announcement.title}</strong> — {announcement.message}
        </MessageBar>
      ))}
    </div>
  );
};
