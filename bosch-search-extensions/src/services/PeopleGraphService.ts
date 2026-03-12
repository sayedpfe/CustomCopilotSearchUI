import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IEnrichedPerson, IPersonSummary, IRecentDocument } from '../models';

export class PeopleGraphService {
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  public async searchPeople(query: string): Promise<IEnrichedPerson[]> {
    console.log(`[PeopleGraphService] GET /me/people — query="${query}"`);
    const response = await this.graphClient
      .api('/me/people')
      .filter(`startswith(displayName,'${query}') or startswith(givenName,'${query}') or startswith(surname,'${query}')`)
      .top(10)
      .select('id,displayName,jobTitle,department,officeLocation,userPrincipalName,scoredEmailAddresses')
      .get();

    const people = response.value || [];
    return people.map((person: Record<string, unknown>) => ({
      id: person.id as string,
      displayName: person.displayName as string,
      mail: ((person.scoredEmailAddresses as Array<Record<string, string>>)?.[0]?.address) || '',
      jobTitle: (person.jobTitle as string) || '',
      department: (person.department as string) || '',
      officeLocation: (person.officeLocation as string) || '',
      photoUrl: '',
    }));
  }

  public async getPersonDetails(userId: string): Promise<IEnrichedPerson> {
    const [user, manager, directReports, recentDocs] = await Promise.all([
      this.getUser(userId),
      this.getManager(userId),
      this.getDirectReports(userId),
      this.getRecentDocuments(userId),
    ]);

    return {
      ...user,
      manager,
      directReports,
      recentDocuments: recentDocs,
    };
  }

  private async getUser(userId: string): Promise<IEnrichedPerson> {
    console.log(`[PeopleGraphService] GET /users/${userId}`);
    const user = await this.graphClient
      .api(`/users/${userId}`)
      .select('id,displayName,mail,jobTitle,department,officeLocation')
      .get();

    return {
      id: user.id,
      displayName: user.displayName || '',
      mail: user.mail || '',
      jobTitle: user.jobTitle || '',
      department: user.department || '',
      officeLocation: user.officeLocation || '',
      photoUrl: '',
    };
  }

  private async getManager(userId: string): Promise<IPersonSummary | undefined> {
    try {
      console.log(`[PeopleGraphService] GET /users/${userId}/manager`);
      const manager = await this.graphClient
        .api(`/users/${userId}/manager`)
        .select('displayName,jobTitle,mail')
        .get();

      return {
        displayName: manager.displayName || '',
        jobTitle: manager.jobTitle || '',
        mail: manager.mail || '',
      };
    } catch {
      return undefined;
    }
  }

  private async getDirectReports(userId: string): Promise<IPersonSummary[]> {
    try {
      console.log(`[PeopleGraphService] GET /users/${userId}/directReports`);
      const response = await this.graphClient
        .api(`/users/${userId}/directReports`)
        .select('displayName,jobTitle')
        .top(10)
        .get();

      return (response.value || []).map((report: Record<string, string>) => ({
        displayName: report.displayName || '',
        jobTitle: report.jobTitle || '',
      }));
    } catch {
      return [];
    }
  }

  private async getRecentDocuments(userId: string): Promise<IRecentDocument[]> {
    try {
      console.log(`[PeopleGraphService] GET /users/${userId}/drive/recent`);
      const response = await this.graphClient
        .api(`/users/${userId}/drive/recent`)
        .top(5)
        .get();

      return (response.value || []).map((doc: Record<string, unknown>) => ({
        name: (doc.name as string) || '',
        webUrl: (doc.webUrl as string) || '',
        lastModified: (doc.lastModifiedDateTime as string) || '',
      }));
    } catch {
      return [];
    }
  }

  public async getProfilePhoto(userId: string): Promise<string> {
    try {
      console.log(`[PeopleGraphService] GET /users/${userId}/photo/$value`);
      const photoBlob = await this.graphClient
        .api(`/users/${userId}/photo/$value`)
        .responseType('blob' as never)
        .get();

      return URL.createObjectURL(photoBlob);
    } catch {
      return '';
    }
  }
}
