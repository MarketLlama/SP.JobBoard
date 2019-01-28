import { Client } from '@microsoft/microsoft-graph-client';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

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
  siteId: string;
  listId: string;
}

export interface IGraphServiceProps {
  context: WebPartContext;
}

export class GraphService {
  private _context: WebPartContext;

  constructor(props: IGraphServiceProps) {
    this._context = props.context;
  }

  public async getSite(client: MSGraphClient) {
    try {
      const host = location.host;
      const serverRelativePath = this._context.pageContext.web.serverRelativeUrl;
      const site = await client.api(`/sites/${host}:${serverRelativePath}`)
        .get();

      return site;
    } catch (error) {
      console.log(error);
    }
  }

  public async getSiteLists(client: MSGraphClient, siteId: string) {
    try {
      const siteLists = await client.api(`/sites/${siteId}/lists`)
        .get();
      return siteLists;
    } catch (error) {
      console.log(error);
    }
  }

  public async getListItems(client: MSGraphClient, siteId: string, listId: string) {
    try {
      const listItems = await client.api(`/sites/${siteId}/lists/${listId}/items`)
        .expand('fields')
        .select('Id,createdBy')
        .get();
      return listItems;
    } catch (error) {
      console.log(error);
    }
  }

  public async setListItem(client: MSGraphClient, siteId: string, listId: string, item: any) {
    try {
      const listItem = await client.api(`/sites/${siteId}/lists/${listId}/items`)
      .post({
        "fields": item
      });
    return listItem;
    } catch (error) {
      console.log(error);
    }

  }
}
