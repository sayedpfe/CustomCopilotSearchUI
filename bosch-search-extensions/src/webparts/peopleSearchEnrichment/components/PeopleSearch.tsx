import * as React from 'react';
import { useState, useEffect } from 'react';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useSearchQuery } from '../../../hooks/useSearchQuery';
import { useGraphClient } from '../../../hooks/useGraphClient';
import { PeopleGraphService } from '../../../services/PeopleGraphService';
import { IEnrichedPerson } from '../../../models';
import { formatDate } from '../../../common/Utils';
import styles from './PeopleSearch.module.scss';

export interface IPeopleSearchProps {
  context: WebPartContext;
  showOrgChart: boolean;
  showRecentDocs: boolean;
  maxRecentDocs: number;
}

export const PeopleSearch: React.FC<IPeopleSearchProps> = ({
  context,
  showOrgChart,
  showRecentDocs,
  maxRecentDocs,
}) => {
  const searchQuery = useSearchQuery();
  const { graphClient } = useGraphClient(context);
  const [people, setPeople] = useState<IEnrichedPerson[]>([]);
  const [expandedPersonId, setExpandedPersonId] = useState<string | null>(null);
  const [detailsMap, setDetailsMap] = useState<Map<string, IEnrichedPerson>>(new Map());
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (!searchQuery || !graphClient) {
      setPeople([]);
      return;
    }

    setLoading(true);
    const service = new PeopleGraphService(graphClient);
    service
      .searchPeople(searchQuery)
      .then((results) => {
        setPeople(results);
        setLoading(false);
      })
      .catch((err) => {
        console.error('[PeopleSearch] Error:', err);
        setPeople([]);
        setLoading(false);
      });
  }, [searchQuery, graphClient]);

  const handleExpand = async (personId: string): Promise<void> => {
    if (expandedPersonId === personId) {
      setExpandedPersonId(null);
      return;
    }

    setExpandedPersonId(personId);

    if (detailsMap.has(personId)) return;

    if (!graphClient) return;
    const service = new PeopleGraphService(graphClient);
    try {
      const details = await service.getPersonDetails(personId);
      setDetailsMap((prev) => new Map(prev).set(personId, details));
    } catch (err) {
      console.error('[PeopleSearch] Error fetching details:', err);
    }
  };

  if (!searchQuery) return null;
  if (loading) return <Spinner label="Searching people..." />;
  if (people.length === 0) return null;

  return (
    <div className={styles.peopleContainer}>
      {people.map((person) => {
        const isExpanded = expandedPersonId === person.id;
        const details = detailsMap.get(person.id);

        return (
          <div key={person.id} className={styles.personCard}>
            <div className={styles.personHeader}>
              <Persona
                text={person.displayName}
                secondaryText={person.jobTitle}
                size={PersonaSize.size48}
                imageUrl={person.photoUrl || undefined}
              />
              <div className={styles.personInfo}>
                <p className={styles.personName}>{person.displayName}</p>
                <p className={styles.personTitle}>{person.jobTitle}</p>
                <p className={styles.personDepartment}>{person.department} {person.officeLocation ? `| ${person.officeLocation}` : ''}</p>
              </div>
            </div>

            <button
              className={styles.expandButton}
              onClick={() => handleExpand(person.id)}
            >
              {isExpanded ? 'Show less' : 'Show more details'}
            </button>

            {isExpanded && details && (
              <>
                {showOrgChart && (
                  <>
                    <div className={styles.sectionHeader}>Organization</div>
                    {details.manager && (
                      <div className={styles.orgChartRow}>
                        <span className={styles.orgLabel}>Manager:</span>
                        <span>{details.manager.displayName} - {details.manager.jobTitle}</span>
                      </div>
                    )}
                    {details.directReports && details.directReports.length > 0 && (
                      <>
                        <div className={styles.orgChartRow}>
                          <span className={styles.orgLabel}>Reports:</span>
                          <span>{details.directReports.length} direct report{details.directReports.length !== 1 ? 's' : ''}</span>
                        </div>
                        {details.directReports.slice(0, 5).map((report, i) => (
                          <div key={i} className={styles.orgChartRow} style={{ paddingLeft: 68 }}>
                            <span>{report.displayName} - {report.jobTitle}</span>
                          </div>
                        ))}
                      </>
                    )}
                  </>
                )}

                {showRecentDocs && details.recentDocuments && details.recentDocuments.length > 0 && (
                  <>
                    <div className={styles.sectionHeader}>Recent Documents</div>
                    <ul className={styles.docList}>
                      {details.recentDocuments.slice(0, maxRecentDocs).map((doc, i) => (
                        <li key={i} className={styles.docItem}>
                          <a
                            href={doc.webUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            className={styles.docLink}
                          >
                            {doc.name}
                          </a>
                          <span className={styles.docDate}>{formatDate(doc.lastModified)}</span>
                        </li>
                      ))}
                    </ul>
                  </>
                )}
              </>
            )}
          </div>
        );
      })}
    </div>
  );
};
