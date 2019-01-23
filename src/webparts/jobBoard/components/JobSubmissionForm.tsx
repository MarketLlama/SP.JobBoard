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
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { taxonomy, setItemMetaDataMultiField, ITerm, ITermData } from "@pnp/sp-taxonomy";
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import JobBoard from './JobBoard';

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
  jobTags?: ITermData[] | null;
  jobTagsString: string;
  team: string;
  areaOfExpertise: string;
  managerId: number;
  managerName: string;
  file: File;
  deadline?: Date | null;
  firstDayOfWeek?: DayOfWeek;
  jobLevel: string;
  jobLevels: any[];
  digitalOrITArr: any[];
  digitalOrIT : string;
  submitDisabled: boolean;
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
      jobTags: null,
      jobTagsString: '',
      managerName: '',
      managerId: 0,
      areaOfExpertise: '',
      team: '',
      file: null,
      deadline: null,
      jobLevel: '',
      jobLevels: [],
      digitalOrITArr: [],
      digitalOrIT : '',
      firstDayOfWeek: DayOfWeek.Sunday,
      submitDisabled: true
    };
  }

  public componentDidMount() {
    this._getJobLevels();
    this._getITorDigital();
  }

  public render() {
    const { firstDayOfWeek, deadline } = this.state;
    const minDate = new Date();
    return (
      <Panel
        isOpen={this.props.parent.state.showSubmissionForm}
        // tslint:disable-next-line:jsx-no-lambda
        onDismiss={() => this.props.parent.setState({ showSubmissionForm: false })}
        type={PanelType.large}
        headerText="Create Job"
        isFooterAtBottom={true}
        onRenderFooterContent={this._onRenderFooterContent}
        className={styles.modalContainer}
      >
        <div id="subtitleId" className={styles.modalBody}>
          <div className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <TextField label="Job Title " required={true}
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
                />
              </div>
            </div>
            <br />
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <TextField label="Job Location " required={true}
                  onChanged={(value) => this.setState({ jobLocation: value })} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <Dropdown
                  placeholder="Select a Job Level"
                  label="Job Level"
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
                  titleText="Manager"
                  showtooltip={true}
                  tooltipDirectional={DirectionalHint.topCenter}
                  tooltipMessage="Surname first to search"
                  personSelectionLimit={1}
                  groupName={""} // IT Leadership
                  isRequired={true}
                  selectedItems={this._setManager}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
              <Dropdown
                  placeholder="Select Digital or IT"
                  label="Digital or IT?"
                  options={this.state.digitalOrITArr}
                  onChanged={(selected) => this.setState({ digitalOrIT: selected.text })}
                  required={true} />
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
          </div>
          <br />
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              <Checkbox label="Has the details of this role been approved by senior management and HR?" onChange={this._enableSubmit} />
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
      </div>
    );
  }


  private _closePanel = () => {
    this.props.parent.setState({
      showSubmissionForm: false
    });
    this.setState({
      file: null,
      jobDescription: ''
    });
  }

  private _handleFile = (files: FileList) => {
    this.setState({
      file: files[0]
    });
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

  private _setJobTags = (items: IPickerTerms) => {
    this._getTerms(items);
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

  private _getTerms = async (items: IPickerTerms) => {
    let promises = [];
    items.forEach(item => {
      promises.push(taxonomy.getDefaultSiteCollectionTermStore().
        getTermById(item.key).get());
    });
    let terms = await Promise.all(promises);
    this.setState({
      jobTags: terms
    });
  }

  private _submitForm = async () => {
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
        Team: s.team,
        Area: s.digitalOrIT,
        View_x0020_Count: 0
      });
      /*
      if (this.state.jobTags.length > 0) {
        let metaDataUpdateResults: ItemUpdateResult = await setItemMetaDataMultiField(itemResult.item, "JobTags", ...this.state.jobTags);
      }*/
      if (this.state.file != null) {
        let item: Item = itemResult.item;
        await item.attachmentFiles.add(this.state.file.name, this.state.file);
      }
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

  private _getITorDigital = async() =>{
    try {
      let field = await this._web.lists.getByTitle('Jobs').fields.getByTitle('Area').get();
      let choices = field.Choices;
      let options : IDropdownOption[] = [];
      choices.forEach(choice => {
        options.push({
          key : choice,
          text : choice
        });
      });
      this.setState({
        digitalOrITArr : options
      });
    } catch (error) {
      console.log(error);
    }
  }

  private _setLoading = (loadingStatus: boolean) => {
    this.setState({
      isLoading: loadingStatus
    });
  }
}

export default JobSubmissionFrom;
