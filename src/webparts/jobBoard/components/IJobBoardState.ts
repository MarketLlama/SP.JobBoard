import { IJob } from "./IJob";
import { IUser } from './IUser';
import { IError } from './IError';
export interface IJobBoardState {
  jobs: any;
  showSubmissionForm : boolean;
  showApplicationForm : boolean;
  selectedJob? : IJob | null;
  isAuthenticated?: boolean;
  user?: IUser | {};
  error?: IError;
}
