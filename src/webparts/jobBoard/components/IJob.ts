
export interface IJob {
  AttachmentFiles: AttachmentFile[];
  Manager: Manager;
  Id: number;
  Title: string;
  Job_x0020_Level?: string;
  Manager_x0020_Name?: string;
  JobTags : JobTags[];
  Location: string;
  Description?: string;
  Deadline: string;
  View_x0020_Count: number;
  Area_x0020_of_x0020_Expertise?: any;
  Team?: any;
  Area?: any;
  ID: number;
  Created: string;
}

export interface Manager {
  JobTitle?: string;
  Name?: string;
  EMail?: string;
  FirstName?: string;
  LastName?: string;
  Id? : number;
}

export interface Author {
  JobTitle?: string;
  Name?: string;
  EMail?: string;
  FirstName?: string;
  LastName?: string;
  Id? : number;
}

export interface AttachmentFile {
  FileName: string;
  FileNameAsPath: FileNameAsPath;
  ServerRelativePath: FileNameAsPath;
  ServerRelativeUrl: string;
}

export interface JobTags {
  Label : string;
  TermGuid : string;
  WssId : number;
}

export interface FileNameAsPath {
  DecodedUrl: string;
}
