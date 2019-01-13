import * as React from 'react';
import styles from './JobBoard.module.scss';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, ItemAddResult, ItemUpdateResult, Item } from '@pnp/pnpjs';
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { PrimaryButton, ActionButton } from 'office-ui-fabric-react/lib/Button';
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
  parent : JobBoard;
  showForm: boolean;
}

export interface JobSubmissionFromState {
  showModal: boolean;
  isLoading : boolean;
  jobTitle: string;
  jobLocation: string;
  jobDescription: string;
  jobTags?: ITermData[] | null;
  jobTagsString: string;
  manager?: IPersonaProps | null;
  managerId: number;
  file: File;
  deadline?: Date | null;
  firstDayOfWeek?: DayOfWeek;
  jobLevel: string;
  jobLevels: any[];
}

class JobSubmissionFrom extends React.Component<JobSubmissionFromProps, JobSubmissionFromState> {
  constructor(props: JobSubmissionFromProps) {
    super(props);
    this.state = {
      showModal: false,
      isLoading : false,
      jobTitle: '',
      jobLocation: '',
      jobDescription: '',
      jobTags: null,
      jobTagsString: '',
      manager: null,
      managerId: 0,
      file: null,
      deadline: null,
      jobLevel: '',
      jobLevels: [],
      firstDayOfWeek: DayOfWeek.Sunday
    };
  }

  public componentDidMount() {
    this._getJobLevels();
  }

  //WARNING! To be deprecated in React v17. Use new lifecycle static getDerivedStateFromProps instead.
  public componentWillReceiveProps(nextProps: JobSubmissionFromProps) {
    this.setState({
      showModal: nextProps.showForm
    });
  }

  public render() {
    const { firstDayOfWeek, deadline } = this.state;
    return (
      <Modal
        titleAriaId="titleId"
        subtitleAriaId="subtitleId"
        isOpen={this.state.showModal}
        onDismiss={this._closeModal}
        isBlocking={false}
        className={styles.modalContainer}
      >
        <div className={styles.modalHeader}>
          <span style={{ padding: "20px" }} id="titleId">Create Job</span>
          <ActionButton className={styles.closeButton} iconProps={{ iconName: 'Cancel' }} onClick={this._closeModal} />
        </div>
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
                <TaxonomyPicker
                  allowMultipleSelections={true}
                  termsetNameOrID="IT Job tags"
                  panelTitle="Select Tag"
                  label="Job Tags"
                  context={this.props.context}
                  onChange={this._setJobTags}
                  isTermSetSelectable={false}
                />
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
            <br />
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg11" />
              <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg1">
                <PrimaryButton
                  value="submit"
                  onClick={this._submitForm}>
                  Submit
                  </PrimaryButton>
              </div>
            </div>
            <br />
          </div>
          { this.state.isLoading ? <Spinner className={styles.loading} size={SpinnerSize.large} label="loading..." ariaLive="assertive" /> : null }
        </div>
      </Modal>
    );
  }

  private _closeModal = () => {
    this.setState({
      showModal: false
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

  private _setManager = (items: IPersonaProps[]) => {
    sp.web.siteUsers.getByLoginName(items[0].id).get().then((profile: any) => {
      this.setState({
        managerId: profile.Id
      });
    });
  }

  private _setJobTags = (items: IPickerTerms) => {
    this._getTerms(items);
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
    this._setLoading(true);
    const s = this.state;
    let itemResult: ItemAddResult = await sp.web.lists.getByTitle('jobs').items.add({
      Title: s.jobTitle,
      Description: s.jobDescription,
      Deadline: s.deadline,
      Location: s.jobLocation,
      ManagerId: s.managerId,
      Job_x0020_Level: s.jobLevel
    });
    if (this.state.jobTags.length > 0) {
      let metaDataUpdateResults: ItemUpdateResult = await setItemMetaDataMultiField(itemResult.item, "JobTags", ...this.state.jobTags);
    }
    if (this.state.file != null) {
      let item: Item = itemResult.item;
      await item.attachmentFiles.add(this.state.file.name, this.state.file);
    }
    this.props.parent.getJobs();
    this._setLoading(false);

    this._closeModal();
  }

  public _setJobDesciption = (e) => {
    this.setState({
      jobDescription: e.target.getContent()
    });
  }

  private _getJobLevels = () => {
    sp.web.lists.getByTitle('Jobs').fields.getByTitle('Job Level').get().then(field => {
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

  private _setLoading = (loadingStatus : boolean) =>{
    this.setState({
      isLoading : loadingStatus
    });
  }
}

export default JobSubmissionFrom;
