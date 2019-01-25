import * as React from 'react';
import styles from './JobBoard.module.scss';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, View, ItemAddResult, Web } from '@pnp/pnpjs';
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Tinymce } from '../global/Tinymce';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { Facepile, IFacepilePersona, IFacepileProps } from 'office-ui-fabric-react/lib/Facepile';
import { PersonaSize, IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import JobBoard from './JobBoard';
import { IJob, Manager, AttachmentFile } from './IJob';
import * as moment from 'moment';
import Moment from 'react-moment';
import { GraphService, IGraphSite, IGraphSiteLists, IGraphIds } from '../global/GraphService';
import Emailer from '../global/Emailer';
import { Panel, PanelType, TextField } from 'office-ui-fabric-react';
import { IJobApplicationGraph } from '../global/IJobApplicationGraph';
import { DirectionalHint } from 'office-ui-fabric-react/lib/Tooltip';
import { IJobApplication } from './JobApplicationsView';

export interface JobApplicationViewProps {
  job? : IJob;
  application : IJobApplication;
}

export interface JobApplicationViewState {

}

export default class JobApplicationView extends React.Component<JobApplicationViewProps, JobApplicationViewState> {

  constructor(props: JobApplicationViewProps) {
    super(props);
    this.state = { : };
  }

  public render() {
    return (
      <div>
        <Panel
          isOpen={this.props.parent.state.showApplicationForm}
          // tslint:disable-next-line:jsx-no-lambda
          onDismiss={() => this.props.parent.setState({ showApplicationForm: false })}
          type={PanelType.large}
          isFooterAtBottom={true}
          onRenderFooterContent={this._onRenderFooterContent}
          headerText="Apply for Job"
          className={styles.modalContainer}
        >
          <div className={styles.modalBody}>
            <div className={[styles.content, "ms-Grid"].join(' ')} dir="ltr">
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <b>Job Title : </b>{job.Title}
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <b>Deadline : </b><Moment format="DD/MM/YYYY">{job.Deadline}</Moment>
                </div>
              </div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <b>Job Location : </b>{job.Location}
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <b>Job Level : </b>{job.Job_x0020_Level}
                </div>
              </div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <b>Team : </b>{job.Team}
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <b>Area of Expertise : </b>{job.Area_x0020_of_x0020_Expertise}
                </div>
              </div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <span style={{ display: 'inline-flex' }}><b>Leader (Contact for the Role) : </b><Facepile {...facepileProps} /> {`${manager.FirstName} ${manager.LastName}`} </span>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <b>IT or Digital : </b>{job.Area}
                </div>
              </div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <p><b>Job Description :</b></p>
                  <div dangerouslySetInnerHTML={{ __html: job.Description }}></div>
                </div>
              </div>
              {attachment.length > 0 ?
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <a href={attachment[0].ServerRelativeUrl}>
                      <FileTypeIcon type={IconType.image} path={attachment[0] ? attachment[0].ServerRelativeUrl : ''} />
                      {attachment[0].FileName}
                    </a>
                  </div>
                </div> : null}
              <h4>Role Application</h4>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">

                </div>
              </div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <p>Cover Note</p>
                  <div dangerouslySetInnerHTML={{ __html: this.props.application.Cover_x0020_Note }}></div>
                </div>
              </div>
            </div>
            {this.state.isLoading ? <Spinner className={styles.loading} size={SpinnerSize.large} label="loading..." ariaLive="assertive" /> : null}
        </Panel>
      </div>
        );

    }
  }
