import * as React from 'react';
import styles from './JobBoard.module.scss';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/pnpjs';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { PrimaryButton, ActionButton } from 'office-ui-fabric-react/lib/Button';
import { Tinymce } from '../global/Tinymce';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { Facepile, IFacepilePersona, IFacepileProps } from 'office-ui-fabric-react/lib/Facepile';
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import JobBoard from './JobBoard';
import { IJob, Manager, AttachmentFile} from './IJob';
import Moment from 'react-moment';
import { GraphService , IGraphSite, IGraphSiteLists, IGraphIds} from '../global/GraphService';
import Emailer from '../global/Emailer';

export interface JobApplicationFormProps {
  job: IJob;
  context: WebPartContext;
  accessToken : string;
  parent: JobBoard;
  showForm: boolean;
}

export interface JobApplicationFormState {
  showModal: boolean;
  isLoading: boolean;
  file: File;
  jobDetails?: any;
  jobTagLabels: string;
  applicationText: string;
}


class JobApplicationForm extends React.Component<JobApplicationFormProps, JobApplicationFormState> {
  private _graphService : GraphService;
  private _graphServiceDetails : IGraphIds;

  constructor(props: JobApplicationFormProps) {
    super(props);
    this.state = {
      showModal: false,
      isLoading: false,
      jobTagLabels: '',
      applicationText: '',
      file: null
    };
    this._graphService = new GraphService({context : this.props.context});
  }

  public render() {

    const job = (this.props.job ? this.props.job : {} as IJob);
    const manager = (job.Manager ? job.Manager : {} as Manager);
    const attachment = (job.AttachmentFiles ? job.AttachmentFiles : [{}] as AttachmentFile[]);

    let persona: IFacepilePersona[] = [{
      imageUrl: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${manager.EMail}&UA=0&size=HR64x64`
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

    return (
      <Modal
        titleAriaId="titleId"
        subtitleAriaId="subtitleId"
        isOpen={this.state.showModal}
        onDismiss={this._closeModal}
        isBlocking={false}
        className={styles.modalContainer}
        onLayerDidMount={this._onLayerMount}
      >
        <div className={styles.modalHeader}>
          <span style={{ padding: "20px" }} id="titleId">Application : {job ? job.Title : ''}</span>
          <ActionButton className={styles.closeButton} iconProps={{ iconName: 'Cancel' }} onClick={this._closeModal} />
        </div>
        <div id="subtitleId" className={styles.modalBody}>
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
                <span style={{ display: 'inline-flex' }}><b>Manager : </b><Facepile {...facepileProps} /> {`${manager.FirstName} ${manager.LastName}`} </span>
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Job Tags : </b>{this.state.jobTagLabels}
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
            <br />
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                <p>Cover Note</p>
                <Tinymce onChange={this._setJobApplicationText} />
              </div>
            </div>
          </div>
          <br />
          <div className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg3">
                <input type="File"
                  id="file"
                  onChange={(e) => this._handleFile(e.target.files)}
                  style={{ display: "none" }} />
                <PrimaryButton iconProps={{ iconName: 'Upload' }}
                  id="button"
                  value="Upload"
                  onClick={() => { document.getElementById("file").click(); }}>
                    Upload Supporting Document
                  </PrimaryButton>
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg9">
                <span className={styles.fileName}>{this.state.file ? this.state.file.name : ''}</span>
              </div>
            </div>
            <br />
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg11" />
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg1">
                <PrimaryButton
                  value="submit"
                  onClick={this._submitForm}>
                  Apply
                  </PrimaryButton>
              </div>
            </div>
          </div>
          {this.state.isLoading ? <Spinner className={styles.loading} size={SpinnerSize.large} label="loading..." ariaLive="assertive" /> : null}
        </div>
      </Modal>
    );
  }

  private _onLayerMount = async () => {
    this._getJobDetails();
    await this._getListDetails();
  }

  //TODO : Make function less chatty, but this will have to do for now.
  private _getListDetails = async () =>{
    try{
      let site : IGraphSite = await this._graphService.getSite(this.props.accessToken);

      let siteLists : IGraphSiteLists = await this._graphService.getSiteLists(this.props.accessToken, site.id);

      let listArray = siteLists.value;
      let jobApplicationList = listArray.filter(list =>{
        return list.name == "Job Applications";
      });

      if(!jobApplicationList[0]){
        console.log('No list called Job Applications in site');
        this._closeModal();
      } else {
        //set Id needed for Graph API calls...
        this._graphServiceDetails = {
          siteId : site.id,
          listId : jobApplicationList[0].id
        };
      }
    } catch (error){
      console.log(error);
    }
  }

  public componentWillReceiveProps(nextProps: JobApplicationFormProps) {
    this.setState({
      showModal: nextProps.showForm
    });
  }

  private _closeModal = () => {
    this.setState({
      showModal: false
    });
  }

  public _setJobApplicationText = (e) => {
    this.setState({
      applicationText: e.target.getContent()
    });
  }

  private _handleFile = (files: FileList) => {
    this.setState({
      file: files[0]
    });
  }

  private _submitForm = async () => {
    try {
      let results = await this._graphService.setListItem(this.props.accessToken, this._graphServiceDetails.siteId, this._graphServiceDetails.listId, {
        Cover_x0020_Note: this.state.applicationText,
        JobLookupId: this.props.job.Id,
        Title : 'Something'
      });
      let emailer : Emailer = new Emailer();
      await emailer.postMail(this.props.accessToken);
      this._closeModal();
    } catch (error) {
      console.log(error);
    }
  }

  private _getJobDetails = async () => {

    if (this.props.job) {
      let job: IJob = await sp.web.lists.getByTitle('Jobs').items.getById(this.props.job.Id).get();
      let tagLabels: string = '';
      if (job.JobTags) {
        job.JobTags.forEach(tag => {
          tagLabels += ` ${tag.Label};`;
        });
      }
      this.setState({
        jobDetails: job,
        jobTagLabels: tagLabels
      });

    }
  }
}

export default JobApplicationForm;
