import * as React from 'react';
import styles from './JobBoard.module.scss';
import { IJobBoardProps } from './IJobBoardProps';
import { IJobBoardState } from './IJobBoardState';
import { IJob, IJobTag } from './IJob';
import { IJobApplication } from './IJobApplication';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import {
  IDocumentCardLogoProps,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardLogo,
  DocumentCardTitle
} from 'office-ui-fabric-react/lib/DocumentCard';
import { personaPresenceSize } from 'office-ui-fabric-react/lib/Persona';
import pnp from "@pnp/pnpjs";
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
    let jobItems : IJob[] = await pnp.sp.web.lists.getByTitle('Jobs').items.get();
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

    const logoProps: IDocumentCardLogoProps = {
      logoIcon: 'OutlookLogo'
    };

    return (
      <div className={styles.brick}>
        <DocumentCard onClickHref="http://bing.com">
          <DocumentCardLogo {...logoProps} />
          <DocumentCardTitle
            title={job.Title}
            shouldTruncate={true}
          />
          <DocumentCardActivity
            activity="Created a few minutes ago"
            people={[{ name: 'Annie Lindqvist', profileImageSrc: ''}]}
          />
        </DocumentCard>
      </div>
    );
  }


}
