import { sp, EmailProperties, PrincipalType, PrincipalSource } from "@pnp/sp";
import { Client } from '@microsoft/microsoft-graph-client';
import * as moment from 'moment';
import '../emailContent/standardEmailTemplate.html';
import { IJob } from "../components/IJob";
import { IJobApplicationGraph , hrManager } from "./IJobApplicationGraph";
import { SiteUser } from "@pnp/sp/src/siteusers";

export default class Emailer {
  private _emailTemplate = require("../emailContent/standardEmailTemplate.html");
  private _newJobEmailTemplate = require("../emailContent/jobCreatedEmailTemplate.html");


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

  private _getManager = async(userId : number) => {
    let managerDetails : any;
    try {
      managerDetails = await sp.web.getUserById(userId).get();
    } catch (error) {
      console.log(error);
    }

    return managerDetails;
  }

  private _getHRManagerDetails = async(hrEmail : string) =>{
    let user : hrManager;
    try {
      user = await sp.utility.resolvePrincipal(hrEmail,
          PrincipalType.User,
          PrincipalSource.All,
          true,
          false);

    } catch (error) {
      console.log(error);
    }
    return user;
  }

  private _getNewJobEmailContent = async (hr : hrManager, job : IJob) =>{
    let emailTemplate = this._newJobEmailTemplate.toString();
    let creator = await this._getManager(job.AuthorId);

    console.log(job);
    emailTemplate = emailTemplate.replace(/{{userName}}/gi, `${hr.DisplayName}`)
    .replace(/{{jobName}}/gi, job.Title)
    .replace(/{{jobLocation}}/gi, job.Location)
    .replace(/{{jobLevel}}/gi, job.Job_x0020_Level)
    .replace(/{{deadline}}/gi, moment(job.Deadline).format('YYYY-MM-DD'))
    .replace(/{{managerName}}/gi, creator.Title)
    .replace(/{{areaOfExpertise}}/gi, job.Area_x0020_of_x0020_Expertise)
    .replace(/{{team}}/gi, job.Team)
    .replace(/{{roleContact}}/gi, job.Manager_x0020_Name)
    .replace(/{{jobDescription}}/gi, job.Description);
    return emailTemplate;
  }

  public sendNewJobEmail = async(accessToken, hrEmail : string, job : IJob) =>{
    const client = this._getAuthenticatedClient(accessToken);

    const hrManagerDetails : hrManager = await this._getHRManagerDetails(hrEmail);
    const userEmail : string = await this._getUsersEmail();

    const emailTemplate : string = await this._getNewJobEmailContent(hrManagerDetails, job);

    const mail = {
      subject: `New Job Created : ${job.Title}`,
      toRecipients: [{
        emailAddress: {
          address: hrManagerDetails.Email
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
      }
    };

    client.api('/users/me/sendMail')
      .post({ message: mail }, (err, res) => {
        console.log(res);
        if(err){ console.log(err); }
      });
  }

  private _getEmailContent = async (job : IJob, application : IJobApplicationGraph) => {
    let manager = await this._getManager(application.fields.Current_x0020_ManagerLookupId);
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
          .replace(/{{currentRole}}/gi, application.fields.Current_x0020_Role)
          .replace(/{{currentManager}}/gi, manager.Title)
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
    const emailTemplate : string = await this._getEmailContent(job, application);

    const mail = {
      subject: `Job Application : ${job.Title}`,
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
      }
    };

    if(file){
      let fileString = await this._getBase64(file);
      mail["attachments"] = [
        {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: file.name,
          contentType : file.type,
          contentBytes : fileString
        }
      ];
    }

    client.api('/users/me/sendMail')
      .post({ message: mail }, (err, res) => {
        console.log(res);
        if(err){ console.log(err); }
      });
  }
}
