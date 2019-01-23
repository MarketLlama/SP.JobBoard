import { sp, EmailProperties } from "@pnp/sp";
import { Client } from '@microsoft/microsoft-graph-client';
import * as moment from 'moment';
import '../emailContent/standardEmailTemplate.html';
import { IJob } from "../components/IJob";
import { IJobApplicationGraph } from "./IJobApplicationGraph";

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


  private _getEmailContent = (job : IJob, application : IJobApplicationGraph) => {
    let emailTemplate = this._emailTemplate.toString();
    emailTemplate = emailTemplate.replace(/{{userName}}/gi, `${job.Manager.FirstName}`)
          .replace(/{{jobName}}/gi, job.Title)
          .replace(/{{jobLocation}}/gi, job.Location)
          .replace(/{{jobLevel}}/gi, job.Job_x0020_Level)
          .replace(/{{deadline}}/gi, moment(job.Deadline).format('YYYY-MM-DD'))
          .replace(/{{jobDescription}}/gi, job.Description)
          .replace(/{{appName}}/gi, application.createdBy.user.displayName)
          .replace(/{{appDate}}/gi, moment(application.createdDateTime).format('YYYY-MM-DD'))
          .replace(/{{areaOfExpertise}}/gi, job.Area_x0020_of_x0020_Expertise)
          .replace(/{{team}}/gi, job.Team)
          .replace(/{{coverNote}}/gi, application.fields.Cover_x0020_Note);
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


  public postMail = async (accessToken, file : File, job :IJob , application : IJobApplicationGraph) => {
    const client = this._getAuthenticatedClient(accessToken);

    const userEmail : string = await this._getUsersEmail();
    const emailTemplate = this._getEmailContent(job, application);

    let fileString = await this._getBase64(file);

    const mail = {
      subject: "Job Application",
      toRecipients: [{
        emailAddress: {
          address: job.Manager.EMail
        }
      }],
      ccRecipients:[{
        emailAddress: {
          address: userEmail
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
