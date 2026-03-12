export interface IPromotedResult {
  id: number;
  title: string;
  description: string;
  url: string;
  keywords: string;
  iconUrl?: string;
  isActive: boolean;
  startDate?: string;
  endDate?: string;
  sortOrder: number;
}
