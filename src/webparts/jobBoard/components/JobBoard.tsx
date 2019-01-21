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
import pnp, { Web } from "@pnp/pnpjs";
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { SecurityTrimmedControl , PermissionLevel} from "@pnp/spfx-controls-react/lib/SecurityTrimmedControl";
import { SPPermission } from '@microsoft/sp-page-context';
import MSALConfig from '../global/MSAL-Config';
import JobSubmissionFrom from './JobSubmissionForm';
import JobApplicationForm from './JobApplicationForm';
import JobApplicationView from './JobApplicationsView';
import JobFilterPanel from './JobFilterPanel';
import { Icon } from 'office-ui-fabric-react/lib/Icon';


export default class JobBoard extends React.Component<IJobBoardProps, IJobBoardState> {
  private _accessToken : string;
  constructor(props) {
    super(props);
    this.state = {
        jobs : [],
        showApplicationForm : false,
        showSubmissionForm : false,
        selectedJob : null,
        showFilter : false
    };
  }

  public render(): React.ReactElement<IJobBoardProps> {
    const remoteSite : string = `${this.props.context.pageContext.web.absoluteUrl}`;
    const listUrl : string = `${this.props.context.pageContext.web.absoluteUrl }/Lists/Jobs`;
    return (
      <div className={styles.jobBoard}>
        <div className={styles.container}>
          {this.state.error ?
            <ErrorMessage debug={this.state.error.debug} message={this.state.error.message} /> : null}
          <Pivot linkFormat={PivotLinkFormat.links} linkSize={PivotLinkSize.normal}>
            <PivotItem linkText="Jobs">
              <IconButton iconProps={{ iconName: 'filter' }} title="filter" ariaLabel="filter" style={{right: 0, position:'fixed'}} onClick={this._showFilter}/>
              <br/>
              <SecurityTrimmedControl context={this.props.context}
                level={PermissionLevel.remoteListOrLib}
                remoteSiteUrl={remoteSite}
                relativeLibOrListUrl={listUrl}
                permissions={[SPPermission.addListItems]}>
                <PrimaryButton
                  disabled={false}
                  iconProps={{ iconName: 'Add' }}
                  text="New Job"
                  onClick={this._newJob}
                />
              </SecurityTrimmedControl>
              <br />
              <div className={styles.masonry}>
                {this.state.jobs}
              </div>
            </PivotItem>
            <PivotItem linkText="Applications">
              <br/>
              <JobApplicationView context={this.props.context} />
            </PivotItem>
          </Pivot>
        </div>
        <JobApplicationForm context={this.props.context} parent={this}
          job={this.state.selectedJob} accessToken={this._accessToken} />
        <JobSubmissionFrom  context={this.props.context} parent={this} />
        <JobFilterPanel showPanel={this.state.showFilter} parent={this} />
      </div>
    );
  }

  private _showFilter = () =>{
    this.setState({
      showFilter : true
    });
  }

  public componentDidMount() {
    this.login();
    this.getJobs();
  }


  protected login = async () => {
    try {
      if (!this.props.user) {
        await this.props.userAgentApplication.loginPopup(MSALConfig.scopes);
        this._accessToken = await this.props.userAgentApplication.acquireTokenSilent(MSALConfig.scopes);
      } else {
        this._accessToken = await this.props.userAgentApplication.acquireTokenSilent(MSALConfig.scopes);
      }

    }
    catch (err) {
      var errParts = err.split('|');
      this.setState({
        isAuthenticated: false,
        user: {},
        error: { message: errParts[1], debug: errParts[0] }
      });
    }
  }

  protected logout = () => {
    this.props.userAgentApplication.logout();
  }

  private _newJob = () => {
    this.setState({
      showSubmissionForm : true
    });
  }

  public getJobs = async () =>{
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    let _jobs = [];
    let jobItems : IJob[] = await web.lists.getByTitle('Jobs').items
      .expand('Manager', 'AttachmentFiles').select('Id','Title','Location','Deadline','Description', 'Created', 'Job_x0020_Level',
        'Manager/JobTitle','Manager/Name', 'Manager/EMail', 'AttachmentFiles', 'JobTags', 'View_x0020_Count',
        'Manager/FirstName', 'Manager/LastName').get();
    for (let i = 0; i < jobItems.length ; i++) {
      _jobs.push(this._onRenderJobCard(jobItems[i]));
    }
    this.setState({
      jobs : _jobs
    });
  }

  private _onRenderJobCard = (job : IJob) : JSX.Element =>{

    let jobTags = [];
    if(job.JobTags.length > 0 ){
      job.JobTags.forEach(tag =>{
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
                <li><b>Job Location</b> : {job.Location}</li>
                <li><b>Job Level</b> : {job.Job_x0020_Level}</li>
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
              activity="is the hiring manager"
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
                    this._showJob(job);
                    ev.preventDefault();
                    ev.stopPropagation();
                  },
                  ariaLabel: 'share action'
                }
              ]}
              views={job.View_x0020_Count}
            />

          </div>
        </DocumentCard>
      </div>
    );
  }

  private _showJob = (job : IJob) =>{
    this.setState({
      selectedJob : job,
      showApplicationForm : true
    });
  }
}
