import { Client } from '@microsoft/microsoft-graph-client';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IGraphSite {
  createdDateTime: string;
  description: string;
  id: string;
  lastModifiedDateTime: string;
  name: string;
  webUrl: string;
  displayName: string;
}
export interface IGraphSiteLists {
  value: IGraphSiteList[];
}

export interface IGraphSiteList {
  createdDateTime: string;
  description: string;
  eTag: string;
  id: string;
  lastModifiedDateTime: string;
  name: string;
  webUrl: string;
  displayName: string;
  createdBy: CreatedBy;
  list: ListDetails;
}

export interface ListDetails {
  contentTypesEnabled: boolean;
  hidden: boolean;
  template: string;
}

export interface CreatedBy {
  user: User;
}

export interface User {
  email: string;
  id: string;
  displayName: string;
}

export interface IListObj {
  value: IListItem[];
}

export interface IListItem {

  id: string;
  createdBy: CreatedBy;
  fields: IListItemFields;
}

export interface IListItemFields {

  Title: string;
  Cover_x0020_Note: string;
  JobLookupId: string;
  id: string;
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
}

export interface IGraphIds {
    siteId : string;
    listId : string;
}

export interface IGraphServiceProps {
    context : WebPartContext;
}

export class GraphService {
    private _context : WebPartContext;

    constructor(props : IGraphServiceProps) {
        this._context = props.context;
    }

    private _getAuthenticatedClient(accessToken : string) {
        // Initialize Graph client
        const client : Client = Client.init({
            // Use the provided access token to authenticate
            // requests
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        return client;
    }

    public async getSite (accessToken  : string){
        const host = location.host;
        const serverRelativePath = this._context.pageContext.web.serverRelativeUrl;
        const client = this._getAuthenticatedClient(accessToken);

        const site = await client.api(`/sites/${host}:${serverRelativePath}`)
            .get();

        return site;
    }

    public async getSiteLists (accessToken : string , siteId : string){
        const client = this._getAuthenticatedClient(accessToken);
        const siteLists = await client.api(`/sites/${siteId}/lists`)
            .get();
        return siteLists;
    }

    public async getListItems(accessToken : string , siteId : string, listId : string) {
        const client = this._getAuthenticatedClient(accessToken);

        const listItems = await client.api(`/sites/${siteId}/lists/${listId}/items`)
            .expand('fields')
            .select('Id,createdBy')
            .get();

        return listItems;
    }

    public async setListItem(accessToken : string,  siteId : string, listId : string ,item : any) {
        const client = this._getAuthenticatedClient(accessToken);
        const listItem = await client.api(`/sites/${siteId}/lists/${listId}/items`)
            .post(
                {
                    "fields": item
                });
        return listItem;
    }
}