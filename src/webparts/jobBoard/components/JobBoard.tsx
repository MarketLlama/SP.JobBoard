import * as React from 'react';
import styles from './JobBoard.module.scss';
import { IJobBoardProps } from './IJobBoardProps';
import { IJobBoardState } from './IJobBoardState';
import { IJob } from './IJob';
import Moment from 'react-moment';
import { PivotLinkSize, PivotLinkFormat, PivotItem, Pivot } from 'office-ui-fabric-react/lib/Pivot';
import { DefaultButton, IButtonProps, PrimaryButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import ErrorMessage from '../global/ErrorMessage';
import {
  IDocumentCardLogoProps,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardLogo,
  DocumentCardTitle,
  DocumentCardPreview,
  DocumentCardActions,
  IDocumentCardPreviewProps,
  IDocumentCardPreviewImage
} from 'office-ui-fabric-react/lib/DocumentCard';
import pnp, { Web, Site } from "@pnp/pnpjs";
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { SecurityTrimmedControl, PermissionLevel } from "@pnp/spfx-controls-react/lib/SecurityTrimmedControl";
import { SPPermission } from '@microsoft/sp-page-context';
import JobSubmissionFrom from './JobSubmissionForm';
import JobApplicationForm from './JobApplicationForm';
import JobApplicationView from './JobApplicationsView';
import JobFilterPanel from './JobFilterPanel';
import JobSubmissionFormEdit from './JobSubmissionFormEdit';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { sp } from '@pnp/sp-addinhelpers';
import { SiteGroup } from '@pnp/sp/src/sitegroups';


export default class JobBoard extends React.Component<IJobBoardProps, IJobBoardState> {
  private _jobs: IJob[] = [];
  constructor(props: IJobBoardProps) {
    super(props);
    this.state = {
      jobs: [],
      showApplicationForm: false,
      showSubmissionForm: false,
      selectedJob: null,
      showFilter: false,
      showEditForm: false,
      selectedId: 0,
      isHR: false,
      isManager: false
    };
    this.getJobs();
    this._checkUserInHRGroup();
    this._checkUserInManagerGroup();
  }

  public render(): React.ReactElement<IJobBoardProps> {

    const remoteSite: string = `${this.props.context.pageContext.web.absoluteUrl}`;
    const listUrl: string = `${this.props.context.pageContext.web.absoluteUrl}/Lists/Jobs`;
    return (
      <div className={styles.jobBoard}>
        <div className={styles.container}>
          {this.state.error ?
            <ErrorMessage debug={this.state.error.debug} message={this.state.error.message} /> : null}
          <Pivot linkFormat={PivotLinkFormat.links} linkSize={PivotLinkSize.normal}>
            <PivotItem linkText="Opportunities">
              {this.state.isManager ?
                <PrimaryButton
                  style={{ margin: '10px', marginLeft: '20px' }}
                  disabled={false}
                  iconProps={{ iconName: 'Add' }}
                  text="New Opportunity"
                  onClick={this._newJob}
                /> : true}
              <br />
              {this.props.isIE ? null :
                <TextField label="Search for Career Opportunity" iconProps={{ iconName: 'Search' }} onKeyUp={this._filterJobs} />}
              <br />
              <div className={styles.masonry}>
                {this.state.jobs}
              </div>
            </PivotItem>
            {this.state.isManager ?
              <PivotItem hidden={!this.state.isManager} linkText="Applications">
                <JobApplicationView context={this.props.context} />
              </PivotItem> : ''}
          </Pivot>
        </div>
        <JobApplicationForm context={this.props.context} parent={this}
          job={this.state.selectedJob} />
        <JobSubmissionFrom context={this.props.context} parent={this} />
        <JobFilterPanel showPanel={this.state.showFilter} parent={this} context={this.props.context} />
        <JobSubmissionFormEdit context={this.props.context} parent={this} job={this.state.selectedJob} />
      </div>
    );
  }

  private _newJob = () => {
    this.setState({
      showSubmissionForm: true
    });
  }

  public getJobs = async () => {
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    let _jobs = [];
    let jobItems: IJob[] = await web.lists.getByTitle('Jobs').items
      .expand('Manager', 'AttachmentFiles').select('Id', 'Title', 'Location', 'Deadline', 'Description', 'Created', 'Job_x0020_Level',
        'Manager/JobTitle', 'Manager/Name', 'Manager/EMail', 'Manager/Id', 'AttachmentFiles', 'JobTags', 'Area', 'Team', 'Area_x0020_of_x0020_Expertise',
        'Manager/FirstName', 'Manager/LastName').get();
    for (let i = 0; i < jobItems.length; i++) {
      _jobs.push(this._onRenderJobCard(jobItems[i]));
    }
    this._jobs = jobItems;
    this.setState({
      jobs: _jobs
    });
  }

  private _onRenderJobCard = (job: IJob): JSX.Element => {

    let jobTags = [];
    if (job.JobTags.length > 0) {
      job.JobTags.forEach(tag => {
        jobTags.push(<li><a href="#" className={styles.tag}>{tag.Label}</a></li>);
      });
    }
    return (
      <div className={styles.brick}>
        <DocumentCard className="ms-fadeIn400">
          <div className="ms-DocumentCard-details">
            <div className={styles.jobTitle}>
              <Icon iconName="Pinned" className={styles.pin} />
              <DocumentCardTitle title={job.Title} shouldTruncate={true} />
            </div>
            <div>
              <ul className={styles.jobDetails}>
                <li><b>Location</b> : {job.Location}</li>
                <li><b>Level</b> : {job.Job_x0020_Level}</li>
                <li><b>Team</b> : {job.Team}</li>
                <li><b>Area of Expertise</b> : {job.Area_x0020_of_x0020_Expertise}</li>
                <li><b>Deadline</b> : <Moment format="DD/MM/YYYY">{job.Deadline}</Moment></li>
              </ul>
            </div>
            {job.AttachmentFiles.length > 0 ?
              <div className={styles.documentLink}>
                <a href={job.AttachmentFiles[0].ServerRelativeUrl}>
                  <FileTypeIcon type={IconType.image} path={job.AttachmentFiles[0] ? job.AttachmentFiles[0].ServerRelativeUrl : ''} />
                  {job.AttachmentFiles[0].FileName}
                </a>
              </div> : null}
            <DocumentCardActivity
              activity="is the contact for the opportunity"
              people={[{
                name: `${job.Manager.FirstName} ${job.Manager.LastName}`,
                profileImageSrc: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${job.Manager.EMail}&UA=0&size=HR64x64`
              }]}
            />
            <ul className={styles.tags}>
              {jobTags}
            </ul>
            <DocumentCardActions
              actions={[
                {
                  iconProps: { iconName: 'OpenInNewWindow' },
                  onClick: (ev: any) => {
                    this._showJobApplication(job);
                    ev.preventDefault();
                    ev.stopPropagation();
                  },
                }, {
                  iconProps: { iconName: 'Delete' },
                  disabled: !this.state.isHR,
                  onClick: (ev: any) => {
                    this._deleteJob(job);
                    ev.preventDefault();
                    ev.stopPropagation();
                  },
                }, {
                  iconProps: { iconName: 'Edit' },
                  disabled: (!this.state.isHR),
                  onClick: (ev: any) => {
                    this._editJob(job);
                    ev.preventDefault();
                    ev.stopPropagation();
                  },
                }
              ]}
            />

          </div>
        </DocumentCard>
      </div>
    );
  }

  private _deleteJob = async (job: IJob) => {
    if (confirm('It it ok to delete this job? \nThis will delete all applications for this job')) {
      const web = new Web(this.props.context.pageContext.web.absoluteUrl);
      await web.lists.getByTitle('Jobs').items.getById(job.ID).delete();
      this.getJobs();
    }
  }

  private _editJob = (job: IJob) => {
    this.setState({
      selectedJob: job,
      showEditForm: true
    });
  }

  private _showJobApplication = (job: IJob) => {
    this.setState({
      selectedJob: job,
      showApplicationForm: true
    });
  }

  private _filterJobs = (event: any): void => {
    let text: string = event.target.value;
    let _jobJSX = [];
    let jobs = text ? this._jobs.filter(i => i.Title.toLowerCase().indexOf(text.toLowerCase()) > -1) : this._jobs;

    for (let i = 0; i < jobs.length; i++) {
      _jobJSX.push(this._onRenderJobCard(jobs[i]));
    }
    this.setState({
      jobs: _jobJSX
    });
  }

  private _checkUserInHRGroup = async () => {
    try {
      const hrGroup: SiteGroup = await sp.web.currentUser.groups.getByName('HR Users').get();
      if (hrGroup) {
        this.setState({
          isHR: true,
          isManager: true
        });
      }
    } catch (error) {
      console.log(error);
    }
  }

  private _checkUserInManagerGroup = async () => {
    try {
      const managerGroup: SiteGroup = await sp.web.currentUser.groups.getByName('Managers').get();
      if (managerGroup) {
        this.setState({
          isManager: true
        });
      }
    } catch (error) {
      console.log(error);
    }
  }
}
