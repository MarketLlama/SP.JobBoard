import * as React from 'react';
import styles from './JobBoard.module.scss';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, ItemAddResult, ItemUpdateResult, Item, Web } from '@pnp/pnpjs';
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Panel, PanelType, Checkbox } from 'office-ui-fabric-react';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DayPickerStrings } from '../global/IDatePickerStrings';
import { Tinymce } from '../global/Tinymce';
import { DirectionalHint } from 'office-ui-fabric-react/lib/Callout';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import JobBoard from './JobBoard';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import Emailer from '../global/Emailer';
import { IJob } from './IJob';

export interface JobSubmissionFromProps {
  context: WebPartContext;
  parent: JobBoard;
}

export interface JobSubmissionFromState {
  showSubmissionPanel: boolean;
  isLoading: boolean;
  jobTitle: string;
  jobLocation: string;
  jobDescription: string;
  team: string;
  areaOfExpertise: string;
  managerId: number;
  managerName: string;
  file: File;
  deadline?: Date | null;
  firstDayOfWeek?: DayOfWeek;
  jobLevel: string;
  jobLevels: any[];
  submitDisabled: boolean;
  hideError : boolean;
}

class JobSubmissionFrom extends React.Component<JobSubmissionFromProps, JobSubmissionFromState> {
  private _web = new Web(this.props.context.pageContext.web.absoluteUrl);

  constructor(props: JobSubmissionFromProps) {
    super(props);
    this.state = {
      showSubmissionPanel: false,
      isLoading: false,
      jobTitle: '',
      jobLocation: '',
      jobDescription: '',
      managerName: '',
      managerId: 0,
      areaOfExpertise: '',
      team: '',
      file: null,
      deadline: null,
      jobLevel: '',
      jobLevels: [],
      firstDayOfWeek: DayOfWeek.Sunday,
      submitDisabled: true,
      hideError : true
    };
  }

  public componentDidMount() {
    this._getJobLevels();
  }

  public render() {
    const { firstDayOfWeek, deadline } = this.state;
    const minDate = new Date();
    const fortnightAway = new Date(Date.now() + 12096e5); //thats 14 day in milliseconds.
    return (
      <Panel
        isOpen={this.props.parent.state.showSubmissionForm}
        // tslint:disable-next-line:jsx-no-lambda
        onDismiss={this._closePanel}
        type={PanelType.large}
        headerText="Create an Opportunity"
        isFooterAtBottom={true}
        onRenderFooterContent={this._onRenderFooterContent}
        className={styles.modalContainer}
      >
        <div id="subtitleId" className={styles.modalBody}>
          <div className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <TextField label="Role Title " required={true}
                  onChanged={(value) => this.setState({ jobTitle: value })} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <DatePicker
                  label="Deadline Date"
                  isRequired={true}
                  firstDayOfWeek={firstDayOfWeek}
                  strings={DayPickerStrings}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  onSelectDate={this._setDeadline}
                  value={deadline!}
                  minDate={minDate}
                  maxDate={fortnightAway}
                />
              </div>
            </div>
            <br />
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <TextField label="Location " required={true}
                  onChanged={(value) => this.setState({ jobLocation: value })} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <Dropdown
                  placeholder="Select a Job Level"
                  label="Level"
                  options={this.state.jobLevels}
                  onChanged={(selected) => this.setState({ jobLevel: selected.text })}
                  required={true} />
              </div>
            </div>
            <br />
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <TextField label="Team " required={true}
                  onChanged={(value) => this.setState({ team: value })} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <TextField label="Area of Expertise" required={true}
                  onChanged={(value) => this.setState({ areaOfExpertise: value })} />
              </div>
            </div>
            <br />
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <PeoplePicker
                  context={this.props.context}
                  titleText="Leader (Contact for the Role) *"
                  showtooltip={true}
                  tooltipDirectional={DirectionalHint.topCenter}
                  tooltipMessage="Surname first to search"
                  personSelectionLimit={1}
                  groupName={""} // IT Leadership
                  isRequired={true}
                  selectedItems={this._setManager}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={500} />
              </div>
            </div>
            <br />
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                <p>Role Description</p>
                <Tinymce onChange={this._setJobDesciption} />
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
          <br />
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              <Checkbox label="Has the opportunity been approved and funded?" onChange={this._enableSubmit} />
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
        <PrimaryButton value="submit" iconProps={{ iconName: 'Add' }} disabled={this.state.submitDisabled}
          style={{ marginRight: '8px' }} onClick={this._submitForm}>Create</PrimaryButton>
        <DefaultButton onClick={this._closePanel}>Cancel</DefaultButton>
        <span className={styles.errorMessage} hidden={this.state.hideError}> <Icon iconName="StatusErrorFull" />
          Please complete all required fields</span>
      </div>
    );
  }

  private _validation = () : boolean =>{
    const s = this.state;
    if(
      s.deadline == null ||
      s.jobTitle == '' || s.jobTitle == null ||
      s.jobLocation == '' || s.jobLocation == null ||
      s.team == '' || s.team == null ||
      s.areaOfExpertise == '' || s.areaOfExpertise == null ||
      s.managerId == 0
    ) {
      this.setState({
        hideError : false
      });
      return false;
    } else {
      return true;
    }
  }

  private _closePanel = () => {
    this.props.parent.setState({
      showSubmissionForm: false
    });
    this.setState({
      file: null,
      jobDescription: '',
      hideError : true,
      submitDisabled: true
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

  private _setDeadline = (date: Date | null | undefined) => {
    this.setState({
      deadline: date
    });
  }

  private _setManager = async (items: IPersonaProps[]) => {
    await this._web.ensureUser(items[0].id);
    this._web.siteUsers.getByLoginName(items[0].id).get().then((profile: any) => {
      console.log(profile);
      this.setState({
        managerId: profile.Id,
        managerName: profile.Title
      });
    });
  }

  private _enableSubmit = (ev : React.FormEvent<HTMLElement | HTMLInputElement>, value: boolean) => {
    if (value) {
      this.setState({
        submitDisabled: false
      });
    } else {
      this.setState({
        submitDisabled: true
      });
    }
  }

  private _submitForm = async () => {
    if(!this._validation()){
      return;
    }
    try {
      this._setLoading(true);
      const s = this.state;
      let itemResult: ItemAddResult = await this._web.lists.getByTitle('jobs').items.add({
        Title: s.jobTitle,
        Description: s.jobDescription,
        Deadline: s.deadline,
        Location: s.jobLocation,
        ManagerId: s.managerId,
        Job_x0020_Level: s.jobLevel,
        Manager_x0020_Name: s.managerName,
        Area_x0020_of_x0020_Expertise: s.areaOfExpertise,
        Team: s.team
      });
      /*
      if (this.state.jobTags.length > 0) {
        let metaDataUpdateResults: ItemUpdateResult = await setItemMetaDataMultiField(itemResult.item, "JobTags", ...this.state.jobTags);
      }*/
      if (this.state.file != null) {
        let item: Item = itemResult.item;
        await item.attachmentFiles.add(this.state.file.name, this.state.file);
      }

      let newJob : IJob= itemResult.data;

      let emailer : Emailer = new Emailer();
      await emailer.sendNewJobEmail(this.props.parent.props.graphClient, this.props.parent.props.hrEmail, newJob);

      this.props.parent.getJobs();
      this._setLoading(false);
      this._closePanel();
    } catch (error) {
      this._setLoading(false);
      console.log(error);
    }
  }

  public _setJobDesciption = (e) => {
    this.setState({
      jobDescription: e.target.getContent()
    });
  }

  private _getJobLevels = () => {
    this._web.lists.getByTitle('Jobs').fields.getByTitle('Job Level').get().then(field => {
      let choices: any[] = field.Choices;
      let jobLevels: IDropdownOption[] = [];
      choices.forEach(choice => {
        jobLevels.push({
          key: choice,
          text: choice
        });
      });
      this.setState({
        jobLevels: jobLevels
      });
    }, error => {
      console.log(error);
    });
  }

  private _setLoading = (loadingStatus: boolean) => {
    this.setState({
      isLoading: loadingStatus
    });
  }
}

export default JobSubmissionFrom;
