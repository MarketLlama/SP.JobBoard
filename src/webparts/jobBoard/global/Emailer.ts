import { sp, EmailProperties } from "@pnp/sp";
import { Client } from '@microsoft/microsoft-graph-client';

export default class Emailer {

  private _getAuthenticatedClient(accessToken: string) {
    // Initialize Graph client
    const client: Client = Client.init({
      // Use the provided access token to authenticate
      // requests
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    return client;
  }

  private _getUsersEmail = async () => {
    let userEmail: string;
    try {
      userEmail = await sp.utility.getCurrentUserEmailAddresses();
    } catch (error) {
      console.log(error);
    }

    return userEmail;
  }

  public postMail = async (accessToken) => {
    const client = this._getAuthenticatedClient(accessToken);
    const email : string = await this._getUsersEmail();
    const mail = {
      subject: "Testing",
      toRecipients: [{
        emailAddress: {
          address: email
        }
      }],
      body: {
        content: "<h1>MicrosoftGraph JavaScript Sample</h1>Check out https://github.com/microsoftgraph/msgraph-sdk-javascript",
        contentType: "html"
      }
    };

    client.api('/users/me/sendMail')
      .post({ message: mail }, (err, res) => {
        console.log(res);
      });
  }
}
