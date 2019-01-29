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
import { IJob, Manager, AttachmentFile} from './IJob';
import * as moment from 'moment';
import Moment from 'react-moment';
import { GraphService , IGraphSite, IGraphSiteLists, IGraphIds} from '../global/GraphService';
import Emailer from '../global/Emailer';
import { Panel , PanelType, TextField} from 'office-ui-fabric-react';
import { IJobApplicationGraph } from '../global/IJobApplicationGraph';
import { DirectionalHint } from 'office-ui-fabric-react/lib/Tooltip';
import { Icon } from 'office-ui-fabric-react/lib/components/Icon';

export interface JobApplicationFormProps {
  job: IJob;
  context: WebPartContext;
  parent: JobBoard;
}

export interface JobApplicationFormState {
  isLoading: boolean;
  file: File;
  jobDetails?: any;
  jobTagLabels: string;
  applicationText: string;
  currentRole : string;
  currentManagerId? : number;
  currentManagerName? : string;
  hideError : boolean;
}

class JobApplicationForm extends React.Component<JobApplicationFormProps, JobApplicationFormState> {
  private _graphService : GraphService;
  private _graphServiceDetails : IGraphIds;
  private _web = new Web(this.props.context.pageContext.web.absoluteUrl);
  private _defaultText = require('../global/applicationDefaultText.html');

  constructor(props: JobApplicationFormProps) {
    super(props);
    this.state = {
      isLoading: false,
      jobTagLabels: '',
      applicationText: '',
      currentRole : '',
      file: null,
      hideError : true
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
      <Panel
        isOpen={this.props.parent.state.showApplicationForm}
        // tslint:disable-next-line:jsx-no-lambda
        onDismiss={this._closePanel}
        type={PanelType.large}
        isFooterAtBottom={true}
        onRenderFooterContent={this._onRenderFooterContent}
        headerText="Apply for Opportunity"
        className={styles.modalContainer}
      >
        <div className={styles.modalBody}>
          <div className={[styles.content, "ms-Grid"].join(' ')} dir="ltr">
            <h4>Opportunity Details</h4>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Opportunity : </b>{job.Title}
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Deadline : </b><Moment format="DD/MM/YYYY">{job.Deadline}</Moment>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Location : </b>{job.Location}
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <b>Level : </b>{job.Job_x0020_Level}
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
                <span style={{ display: 'inline-flex' }}><b>Leader (Contact for the Opportunity) : </b><Facepile {...facepileProps} /> {`${manager.FirstName} ${manager.LastName}`} </span>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                <p><b>Opportunity Description :</b></p>
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
            <hr/>
            <h4>Opportunity Application</h4>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <TextField label="Current Opportunity " required={true}
                  onChanged={(value) => this.setState({ currentRole: value })} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <PeoplePicker
                    context={this.props.context}
                    titleText="Current Manager *"
                    showtooltip={true}
                    tooltipDirectional={DirectionalHint.topCenter}
                    tooltipMessage="Surname first to search"
                    personSelectionLimit={1}
                    groupName={""} // IT Leadership
                    isRequired={true}
                    ensureUser={true}
                    selectedItems={this._setCurrentManager}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={500} />
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                <p>Cover Note (No more than 500 words)</p>
                <Tinymce onChange={this._setJobApplicationText} defaultValue={this._defaultText}/>
              </div>
            </div>
          </div>
          <br />
          <div className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg4">
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
          </div>
        </div>
        {this.state.isLoading ? <Spinner className={styles.loading} size={SpinnerSize.large} label="loading..." ariaLive="assertive" /> : null}
      </Panel>
    );
  }
  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton value="submit" style={{ marginRight: '8px' }} onClick={this._submitForm}>Apply</PrimaryButton>
        <DefaultButton onClick={this._closePanel}>Cancel</DefaultButton>
        <span className={styles.errorMessage} hidden={this.state.hideError}> <Icon iconName="StatusErrorFull" />
          Please complete all required fields</span>
      </div>
    );
  }

  private _validation = () : boolean =>{
    const s = this.state;
    if(
      s.currentRole == '' || s.currentRole == null ||
      s.currentManagerId == 0 || s.currentManagerId == null
    ) {
      this.setState({
        hideError : false
      });
      return false;
    } else {
      return true;
    }
  }

  public componentWillReceiveProps(newProps : JobApplicationFormProps){
    if(newProps.parent.state.showApplicationForm === true){
        this._onLayerMount(newProps);
    }
  }


  private _onLayerMount = async (newProps : JobApplicationFormProps) => {
    await this._getJobDetails(newProps.job);
    await this._getListDetails();
  }

  //TODO : Make function less chatty, but this will have to do for now.
  private _getListDetails = async () =>{
    try{
      let site : IGraphSite = await this._graphService.getSite(this.props.parent.props.graphClient);

      let siteLists : IGraphSiteLists = await this._graphService.getSiteLists(this.props.parent.props.graphClient, site.id);

      let listArray = siteLists.value;
      let jobApplicationList = listArray.filter(list =>{
        return list.name == "Job Applications";
      });

      if(!jobApplicationList[0]){
        console.log('No list called Job Applications in site');
        this._closePanel();
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

  private _closePanel = () => {
    this.props.parent.setState({
      showApplicationForm: false
    });
    this.setState({
      file : null,
      applicationText : '',
      hideError : true
    });
  }



  public _setJobApplicationText = (e) => {
    this.setState({
      applicationText: e.target.getContent()
    });
  }

  private _setCurrentManager = (items: IPersonaProps[]) => {
    const id = parseInt(items[0].id);
    this._web.getUserById(id).get().then((profile: any) => {
      this.setState({
        currentManagerId: profile.Id,
        currentManagerName: profile.Title
      });
    });
  }

  private _handleFile = (files: FileList) => {
    if(files.length > 0){
      this.setState({
        file: files[0]
      });
    } else {
      this.setState({
        file : null
      });
    }
  }

  private _submitForm = async () => {
    if(!this._validation()){
      return;
    }
    this._setLoading(true);
    let now = moment();
    try {
      let result : IJobApplicationGraph = await this._graphService.setListItem(this.props.parent.props.graphClient, this._graphServiceDetails.siteId, this._graphServiceDetails.listId, {
        Cover_x0020_Note: this.state.applicationText,
        Current_x0020_Role : this.state.currentRole,
        Current_x0020_ManagerLookupId : this.state.currentManagerId,
        JobLookupId: this.props.job.Id,
        Title : `${now.format('YYYY-MM-DD')} - ${this.props.context.pageContext.user.displayName}`
      });
      let emailer : Emailer = new Emailer();
      let application  : IJobApplicationGraph = result;
      await emailer.postMail(this.props.parent.props.graphClient, this.state.file, this.state.jobDetails ,application);
      this._closePanel();
      this._setLoading(false);
    } catch (error) {
      console.log(error);
      this._setLoading(false);
    }
  }

  private _getJobDetails = async (_job : IJob) => {
    let jobId : number = this.props.job?  this.props.job.Id : _job.Id;
    if (jobId) {
      let job: IJob = await this._web.lists.getByTitle('Jobs').items.getById(jobId).expand('Manager', 'AttachmentFiles').select('Id','Title','Location','Deadline','Description', 'Created', 'Job_x0020_Level',
      'Manager/JobTitle','Manager/Name', 'Manager/EMail', 'AttachmentFiles', 'JobTags', 'Area', 'Team', 'Area_x0020_of_x0020_Expertise',
      'Manager/FirstName', 'Manager/LastName').get();
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

      console.log(job);
    }
  }

  private _setLoading = (loadingStatus: boolean) => {
    this.setState({
      isLoading: loadingStatus
    });
  }
}

export default JobApplicationForm;
