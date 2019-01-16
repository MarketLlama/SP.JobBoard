import { IWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";
import { UserAgentApplication, User } from 'msal';

export interface IJobBoardProps {
  description: string;
  context : WebPartContext;
  userAgentApplication : UserAgentApplication;
  user : User;
}
