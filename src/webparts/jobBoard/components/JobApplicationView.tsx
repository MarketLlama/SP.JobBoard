import * as React from 'react';
import styles from './JobBoard.module.scss';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Web } from '@pnp/pnpjs';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { FileTypeIcon, IconType } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { Facepile, IFacepilePersona, IFacepileProps } from 'office-ui-fabric-react/lib/Facepile';
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IJob} from './IJob';
import Moment from 'react-moment';
import { Panel, PanelType, TextField } from 'office-ui-fabric-react';
import  {IJobApplication}  from './IJobApplication';
import JobApplicationsView from './JobApplicationsView';

export interface JobApplicationViewProps {
  context: WebPartContext;
  jobId: number;
  parent: JobApplicationsView;
}

export interface JobApplicationViewState {
  job?: IJob;
  application?: IJobApplication;
}

export default class JobApplicationView extends React.Component<JobApplicationViewProps, JobApplicationViewState> {
  private _web = new Web(this.props.context.pageContext.web.absoluteUrl);

  constructor(props: JobApplicationViewProps) {
    super(props);
    this.state = {
      job: null,
      application: null
    };
  }

  public render() {

    if (this.state.job == null) {
      return (<div />);
    }

    let persona: IFacepilePersona[] = [{
      imageUrl: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${this.state.job.Manager.EMail}&UA=0&size=HR64x64`
    }];

    const facepileProps: IFacepileProps = {
      personaSize: PersonaSize.size32,
      personas: persona,
      className: styles.managerPicture,
      getPersonaProps: () => {
        return {
          imageShouldFadeIn: true
        };
      },
      ariaDescription: 'To move through the items use left and right arrow keys.'
    };

    let currentManagerPersona: IFacepilePersona[] = [{
      imageUrl: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${this.state.application.Current_x0020_Manager.EMail}&UA=0&size=HR64x64`
    }];

    const currentManagerfacepileProps: IFacepileProps = {
      personaSize: PersonaSize.size32,
      personas: currentManagerPersona,
      className: styles.managerPicture,
      getPersonaProps: () => {
        return {
          imageShouldFadeIn: true
        };
      },
      ariaDescription: 'To move through the items use left and right arrow keys.'
    };

    let applicantPersona: IFacepilePersona[] = [{
      imageUrl: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${this.state.application.Author.EMail}&UA=0&size=HR64x64`
    }];

    const applicantfacepileProps: IFacepileProps = {
      personaSize: PersonaSize.size32,
      personas: applicantPersona,
      className: styles.managerPicture,
      getPersonaProps: () => {
        return {
          imageShouldFadeIn: true
        };
      },
      ariaDescription: 'To move through the items use left and right arrow keys.'
    };

    return (
      <Panel
        isOpen={this.props.parent.state.showApplicationPanel}
        // tslint:disable-next-line:jsx-no-lambda
        onDismiss={() => this.props.parent.setState({ showApplicationPanel: false })}
        type={PanelType.large}
        isFooterAtBottom={true}
        onRenderFooterContent={this._onRenderFooterContent}
        headerText="Application Details"
        className={styles.modalContainer}
      >
        <div className={styles.modalBody}>
          <div className={[styles.content, "ms-Grid"].join(' ')} dir="ltr">
            <h4>Opportunity Details</h4>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Job Title : </b>{this.state.job.Title}
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Deadline : </b><Moment format="DD/MM/YYYY">{this.state.job.Deadline}</Moment>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Job Location : </b>{this.state.job.Location}
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Job Level : </b>{this.state.job.Job_x0020_Level}
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Team : </b>{this.state.job.Team}
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Area of Expertise : </b>{this.state.job.Area_x0020_of_x0020_Expertise}
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <span style={{ display: 'inline-flex' }}><b>Leader (Contact for the Role) :
                  </b><Facepile {...facepileProps} /> {`${this.state.job.Manager.FirstName} ${this.state.job.Manager.LastName}`} </span>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                <p><b>Job Description :</b></p>
                <div dangerouslySetInnerHTML={{ __html: this.state.job.Description }}></div>
              </div>
            </div>
            {this.state.job.AttachmentFiles.length > 0 ?
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <a href={this.state.job.AttachmentFiles[0].ServerRelativeUrl}>
                    <FileTypeIcon type={IconType.image} path={this.state.job.AttachmentFiles[0] ?
                      this.state.job.AttachmentFiles[0].ServerRelativeUrl : ''} />
                    {this.state.job.AttachmentFiles[0].FileName}
                  </a>
                </div>
              </div> : null}
            <hr/>
            <h4>Role Application Details</h4>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                {this.state.application.Current_x0020_Manager ?
                <span style={{ display: 'inline-flex' }}><b>Applicant :
                    </b><Facepile {...applicantfacepileProps} /> {`${this.state.application.Author.FirstName} ${this.state.application.Author.LastName}`} </span> :
                    null}
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Application Date : </b> <Moment format="DD/MM/YYYY">{this.state.application.Created}</Moment>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
              {this.state.application.Current_x0020_Manager ?
              <span style={{ display: 'inline-flex' }}><b>Current Manager :
                  </b><Facepile {...currentManagerfacepileProps} /> {`${this.state.application.Current_x0020_Manager.FirstName} ${this.state.application.Current_x0020_Manager.LastName}`} </span> :
                  null}
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Current Role : </b>{this.state.application.Current_x0020_Role}
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                <p><b>Applicant Cover Note :</b></p>
                <div dangerouslySetInnerHTML={{ __html: this.state.application.Cover_x0020_Note }}></div>
              </div>
            </div>
          </div>
        </div>
      </Panel>
    );
  }

  public componentWillReceiveProps(nextProps: JobApplicationViewProps) {
    if (nextProps.parent.state.showApplicationPanel === true) {
      this._getJobApplcationDetails(nextProps);
    }
  }

  private _getJobApplcationDetails = async (nextProps: JobApplicationViewProps) => {
    let Id: number = nextProps.jobId;
    let jobApplication: IJobApplication = await this._web.lists
      .getByTitle('Job Applications').items.getById(Id).expand('Current_x0020_Manager', 'Author').select('Id', 'Title', 'JobId',
      'Cover_x0020_Note', 'Current_x0020_Role' ,'Current_x0020_Manager/Title' , 'Current_x0020_Manager/JobTitle', 'Current_x0020_Manager/Name',
      'Current_x0020_Manager/EMail', 'Current_x0020_Manager/Id', 'Current_x0020_Manager/FirstName' , 'Current_x0020_Manager/LastName',
      'Author/Title' , 'Author/JobTitle', 'Author/Name', 'Author/EMail', 'Author/Id' , 'Author/FirstName' , 'Author/LastName').get();

    let job: IJob = await this._web.lists.getByTitle('Jobs').items.getById(jobApplication.JobId).expand('Manager', 'AttachmentFiles').select('Id', 'Title', 'Location', 'Deadline',
      'Description', 'Created', 'Job_x0020_Level', 'Manager/JobTitle', 'Manager/Name', 'Manager/EMail',
      'Manager/Id', 'AttachmentFiles', 'JobTags', 'View_x0020_Count', 'Area', 'Team', 'Area_x0020_of_x0020_Expertise',
      'Manager/FirstName', 'Manager/LastName').get();


    this.setState({
      job: job,
      application: jobApplication
    });

  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <DefaultButton onClick={this._closePanel}>Close</DefaultButton>
      </div>
    );
  }


  private _closePanel = () => {
    this.props.parent.setState({
      showApplicationPanel: false
    });
  }

}
