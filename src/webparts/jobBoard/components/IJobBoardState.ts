import { IJob } from "./IJob";

export interface IJobBoardState {
  jobs: any;
  showSubmissionForm : boolean;
  showApplicationForm : boolean;
  selectedJob? : IJob | null;
}
