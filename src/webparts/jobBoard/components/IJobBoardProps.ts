import { IWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClient } from '@microsoft/sp-http';

export interface IJobBoardProps {
  description: string;
  graphClient : MSGraphClient;
  context : WebPartContext;
  hrEmail : string;
  isIE : boolean;
}
