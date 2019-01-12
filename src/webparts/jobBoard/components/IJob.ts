
export interface IJob {
  FileSystemObjectType: number;
  Id: number;
  ServerRedirectedEmbedUri?: any;
  ServerRedirectedEmbedUrl: string;
  ContentTypeId: string;
  Title: string;
  ComplianceAssetId?: any;
  Job_x0020_Tags: IJobTag[];
  Job_x0020_Level: string;
  Manager: string;
  Location: string;
  Description: string;
  Deadline: string;
  ID: number;
  Modified: string;
  Created: string;
  AuthorId: number;
  EditorId: number;
  OData__UIVersionString: string;
  Attachments: boolean;
  GUID: string;
}

export interface IJobTag {
  Label: string;
  TermGuid: string;
  WssId: number;
}
