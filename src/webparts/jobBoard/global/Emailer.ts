import { sp, EmailProperties } from "@pnp/sp";
import { Client } from '@microsoft/microsoft-graph-client';
import '../emailContent/standardEmailTemplate.html';

export default class Emailer {
  private _emailTemplate = require("../emailContent/standardEmailTemplate.html");

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

  private _getEmailContent = () => {
    let emailTemplate = this._emailTemplate.toString();
    /*emailTemplate = emailTemplate.replace(/{{emailContent}}/gi, this.state.emailText)
      .replace(/{{pageURL}}/gi, tenantUri + this._currentPage.FileRef)
      .replace(/{{userName}}/gi, user.UserName)
      .replace(/{{pageTitle}}/gi, this._currentPage.Title);*/
    return emailTemplate;
  }

  private _getBase64(file : File) {
    return new Promise((resolve, reject) => {
      const reader : FileReader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        let encoded = reader.result.replace(/^data:(.*;base64,)?/, '');
        resolve(encoded);
      };
      reader.onerror = error => reject(error);
    });
  }


  public postMail = async (accessToken, file : File) => {
    const client = this._getAuthenticatedClient(accessToken);

    const email : string = await this._getUsersEmail();
    const emailTemplate = this._getEmailContent();

    let fileString = await this._getBase64(file);

    const mail = {
      subject: "Job Application",
      toRecipients: [{
        emailAddress: {
          address: email
        }
      }],
      body: {
        content: emailTemplate,
        contentType: "html"
      },
      attachments: [
        {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: file.name,
          contentType : file.type,
          contentBytes : fileString
        }
      ]
    };

    client.api('/users/me/sendMail')
      .post({ message: mail }, (err, res) => {
        console.log(res);
        if(err){ console.log(err); }
      });
  }
}
