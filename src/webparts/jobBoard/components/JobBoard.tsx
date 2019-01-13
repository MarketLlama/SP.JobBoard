import * as React from 'react';
import styles from './JobBoard.module.scss';
import { IJobBoardProps } from './IJobBoardProps';
import { IJobBoardState } from './IJobBoardState';
import { IJob } from './IJob';
import Moment from 'react-moment';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
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
import { personaPresenceSize } from 'office-ui-fabric-react/lib/Persona';
import pnp from "@pnp/pnpjs";
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { SecurityTrimmedControl , PermissionLevel} from "@pnp/spfx-controls-react/lib/SecurityTrimmedControl";
import { SPPermission } from '@microsoft/sp-page-context';
import JobSubmissionFrom from './JobSubmissionForm';

export default class JobBoard extends React.Component<IJobBoardProps, IJobBoardState> {
  constructor(props) {
    super(props);
    this.state = {
        jobs : [],
        showApplicationForm : false,
        showSubmissionForm : false
    };
  }

  public render(): React.ReactElement<IJobBoardProps> {
    const remoteSite : string = `${this.props.context.pageContext.web.absoluteUrl}`;
    const listUrl : string = `${this.props.context.pageContext.web.absoluteUrl }/Lists/Jobs`;
    return (
      <div className={ styles.jobBoard }>
        <div className={ styles.container }>
          <SecurityTrimmedControl context={this.props.context}
                          level={PermissionLevel.remoteListOrLib}
                          remoteSiteUrl={remoteSite}
                          relativeLibOrListUrl={listUrl}
                          permissions={[SPPermission.addListItems]}>

            <DefaultButton
                disabled={false}
                iconProps={{ iconName: 'Add' }}
                text="New Job"
                onClick = {this._newJob}
              />
           </SecurityTrimmedControl>
           <br/>
           <div className={styles.masonry}>
              {this.state.jobs}
           </div>
        </div>
        <JobSubmissionFrom showForm={this.state.showSubmissionForm} context={this.props.context} parent={this}/>
      </div>
    );
  }

  public componentDidMount() {
    this.getJobs();
    this._getJobApplication();
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
    console.log(jobItems);
    for (let i = 0; i < jobItems.length ; i++) {
      _jobs.push(this._onRenderJobCard(jobItems[i]));
    }
    this.setState({
      jobs : _jobs
    });
  }

  private _getJobApplication = () =>{
    pnp.sp.web.lists.getByTitle('Job Applications').items.get().then(items =>{
      console.log(items);
    }, error =>{
      console.log(error);
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
              <span>Location : {job.Location}</span>
              <span>Level : {job.Job_x0020_Level}</span>
              <span>Deadline : <Moment format="DD/MM/YYYY">{job.Deadline}</Moment></span>
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
                  console.log('Open id ' + job.Id);
                  ev.preventDefault();
                  ev.stopPropagation();
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


}
