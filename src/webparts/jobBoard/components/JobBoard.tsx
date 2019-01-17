import * as React from 'react';
import styles from './JobBoard.module.scss';
import { IJobBoardProps } from './IJobBoardProps';
import { IJobBoardState } from './IJobBoardState';
import { IJob } from './IJob';
import Moment from 'react-moment';
import { PivotLinkSize, PivotLinkFormat, PivotItem, Pivot } from 'office-ui-fabric-react/lib/Pivot';
import { DefaultButton, IButtonProps, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
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
import pnp from "@pnp/pnpjs";
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { SecurityTrimmedControl , PermissionLevel} from "@pnp/spfx-controls-react/lib/SecurityTrimmedControl";
import { SPPermission } from '@microsoft/sp-page-context';
import MSALConfig from '../global/MSAL-Config';
import JobSubmissionFrom from './JobSubmissionForm';
import JobApplicationForm from './JobApplicationForm';
import JobApplicationView from './JobApplicationsView';


export default class JobBoard extends React.Component<IJobBoardProps, IJobBoardState> {
  private _accessToken : string;
  constructor(props) {
    super(props);
    this.state = {
        jobs : [],
        showApplicationForm : false,
        showSubmissionForm : false,
        selectedJob : null
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
              <JobApplicationView />
            </PivotItem>
          </Pivot>
        </div>
        <JobSubmissionFrom showForm={this.state.showSubmissionForm} context={this.props.context} parent={this} />
        <JobApplicationForm showForm={this.state.showApplicationForm} context={this.props.context} parent={this}
          job={this.state.selectedJob} accessToken={this._accessToken} />
      </div>
    );
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
    let _jobs = [];
    let jobItems : IJob[] = await pnp.sp.web.lists.getByTitle('Jobs').items
      .expand('Manager', 'AttachmentFiles').select('Id','Title','Location','Deadline','Description', 'Created', 'Job_x0020_Level',
        'Manager/JobTitle','Manager/Name', 'Manager/EMail', 'AttachmentFiles',
        'Manager/FirstName', 'Manager/LastName').get();
    for (let i = 0; i < jobItems.length ; i++) {
      _jobs.push(this._onRenderJobCard(jobItems[i]));
    }
    this.setState({
      jobs : _jobs
    });
  }

  private _onRenderJobCard = (job : IJob) : JSX.Element =>{
    const previewPropsUsingIcon: IDocumentCardPreviewProps = {
      previewImages: [
        {
          previewIconProps: { iconName: 'OpenFile', styles: { root: { fontSize: 42, color: '#ffffff' } } },
          height: 150
        }
      ]
    };
    return (
      <div className={styles.brick}>
        <DocumentCard>
          <DocumentCardPreview {...previewPropsUsingIcon} />
          <div className="ms-DocumentCard-details">
            <DocumentCardTitle title={job.Title} shouldTruncate={true} />
            <div>
              <ul className={styles.jobDetails}>
                <li><b>Location</b> : {job.Location}</li>
                <li><b>Level</b> : {job.Job_x0020_Level}</li>
                <li><b>Deadline</b> : <Moment format="DD/MM/YYYY">{job.Deadline}</Moment></li>
              </ul>
            </div>
            {job.AttachmentFiles.length > 0 ?
            <div className={styles.documentLink}>
              <a href={job.AttachmentFiles[0].ServerRelativeUrl}>
                <FileTypeIcon type={IconType.image} path={job.AttachmentFiles[0]? job.AttachmentFiles[0].ServerRelativeUrl : ''} />
                {job.AttachmentFiles[0].FileName}
              </a>
            </div> : null}
            <DocumentCardActivity
              activity="Created By"
              people={[{ name: `${job.Manager.FirstName} ${job.Manager.LastName}`,
              profileImageSrc: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${job.Manager.EMail}&UA=0&size=HR64x64` }]}
            />
          <DocumentCardActions
            actions={[
              {
                iconProps: { iconName: 'OpenInNewWindow' },
                onClick: (ev: any) => {
                  this._showJob(job);
                  //ev.preventDefault();
                  //ev.stopPropagation();
                },
                ariaLabel: 'share action'
              }
            ]}
            views={432}
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
