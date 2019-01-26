
export interface IJobApplication {
  FileSystemObjectType: number;
  Id: number;
  ServerRedirectedEmbedUri?: any;
  ServerRedirectedEmbedUrl: string;
  ContentTypeId: string;
  Title: string;
  ComplianceAssetId?: any;
  Cover_x0020_Note: string;
  Current_x0020_Role : string;
  Current_x0020_Manager : Manager;
  JobId: number;
  ID: number;
  Modified: string;
  Created: string;
  Author : Author;
  AuthorId: number;
  EditorId: number;
  OData__UIVersionString: string;
  Attachments: boolean;
  GUID: string;
}

export interface Manager {
  JobTitle?: string;
  Name?: string;
  EMail?: string;
  FirstName?: string;
  LastName?: string;
}

export interface Author {
  JobTitle?: string;
  Name?: string;
  EMail?: string;
  FirstName?: string;
  LastName?: string;
  Id? : number;
}
