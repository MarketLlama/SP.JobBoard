import { IWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";
import { UserAgentApplication, User } from 'msal';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IJobBoardProps {
  description: string;
  graphClient : MSGraphClient;
  context : WebPartContext;
  hrEmail : string;
  isIE : boolean;
}
