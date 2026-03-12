export interface IPersonSummary {
  displayName: string;
  jobTitle: string;
  mail?: string;
}

export interface IRecentDocument {
  name: string;
  webUrl: string;
  lastModified: string;
}

export interface IEnrichedPerson {
  id: string;
  displayName: string;
  mail: string;
  jobTitle: string;
  department: string;
  officeLocation: string;
  photoUrl: string;
  manager?: IPersonSummary;
  directReports?: IPersonSummary[];
  recentDocuments?: IRecentDocument[];
  expertiseTags?: string[];
  costCenter?: string;
  ntId?: string;
}
