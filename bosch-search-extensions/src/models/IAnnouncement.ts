export type AnnouncementSeverity = 'Info' | 'Warning' | 'Error' | 'Success';
export type TargetAudience = 'All' | 'HR' | 'Sales' | 'IT' | 'Engineering';

export interface IAnnouncement {
  id: number;
  title: string;
  message: string;
  severity: AnnouncementSeverity;
  startDate: string;
  endDate: string;
  isActive: boolean;
  targetAudience: TargetAudience;
  sortOrder: number;
}
