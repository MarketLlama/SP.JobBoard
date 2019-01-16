import { Client } from '@microsoft/microsoft-graph-client';

export class GraphService {

    private _getAuthenticatedClient(accessToken) {
        // Initialize Graph client
        const client = Client.init({
            // Use the provided access token to authenticate
            // requests
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        return client;
    }

    public async getUserDetails(accessToken) {
        const client = this._getAuthenticatedClient(accessToken);

        const user = await client.api('/me').get();
        return user;
    }

    public async  getEvents(accessToken) {
        const client = this._getAuthenticatedClient(accessToken);

        const events = await client
            .api('/me/events')
            .select('subject,organizer,start,end')
            .orderby('createdDateTime DESC')
            .get();

        return events;
    }

    public async getListItems(accessToken) {
        const client = this._getAuthenticatedClient(accessToken);

        const listItems = await client.api('/sites/troposphere.sharepoint.com,469ce93f-2c9a-4b5f-8aeb-f5d0202e5c99,15465711-d00b-4514-849c-38062b5ee76e/lists/b1961348-c4d6-4827-95b3-7bc3780296c7/items')
            .expand('fields')
            .select('Id,createdBy')
            .get();

        return listItems;
    }

    public async setListItem(accessToken, item) {
        const client = this._getAuthenticatedClient(accessToken);

        const listItem = await client.api('/sites/troposphere.sharepoint.com,469ce93f-2c9a-4b5f-8aeb-f5d0202e5c99,15465711-d00b-4514-849c-38062b5ee76e/lists/b1961348-c4d6-4827-95b3-7bc3780296c7/items')
            .post(
                {
                    "fields": {
                        "Title": "Title A",
                        "Location": "Manchester"
                    }
                }
            )
        return listItem;
    }
}