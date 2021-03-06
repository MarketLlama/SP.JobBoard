export interface IJobApplicationGraph {
  createdDateTime: string;
  eTag: string;
  id: string;
  lastModifiedDateTime: string;
  webUrl: string;
  createdBy: CreatedBy;
  lastModifiedBy: CreatedBy;
  contentType: ContentType;
  fields: Fields;
}

export interface Fields {
  Title: string;
  Cover_x0020_Note: string;
  JobLookupId: string;
  Job_x003a_TitleLookupId: string;
  Job_x003a_LocationLookupId: string;
  Job_x003a_DeadlineLookupId: string;
  Job_x003a_CreatedLookupId: string;
  id: string;
  Current_x0020_Role : string;
  Current_x0020_ManagerLookupId : number;
  ContentType: string;
  Modified: string;
  Created: string;
  AuthorLookupId: string;
  EditorLookupId: string;
  _UIVersionString: string;
  Attachments: boolean;
  Edit: string;
  LinkTitleNoMenu: string;
  LinkTitle: string;
  ItemChildCount: string;
  FolderChildCount: string;
  _ComplianceFlags: string;
  _ComplianceTag: string;
  _ComplianceTagWrittenTime: string;
  _ComplianceTagUserId: string;
  AppAuthorLookupId: string;
  AppEditorLookupId: string;
}

export interface ContentType {
  id: string;
}

export interface CreatedBy {
  user: User;
}

export interface User {
  email: string;
  id: string;
  displayName: string;
}
export interface hrManager{
  Department: string;
  DisplayName: string;
  Email: string;
  JobTitle: string;
  LoginName: string;
  Mobile: string;
  PrincipalId: number;
  PrincipalType: number;
  SIPAddress: string;
}
