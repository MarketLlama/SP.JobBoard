
export interface IJob {
  AttachmentFiles: AttachmentFile[];
  Manager: Manager;
  Id: number;
  Title: string;
  Job_x0020_Level?: string;
  JobTags : JobTags[];
  Location: string;
  Description?: string;
  Deadline: string;
  ID: number;
  Created: string;
}

export interface Manager {
  JobTitle?: string;
  Name?: string;
  EMail?: string;
  FirstName?: string;
  LastName?: string;
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
