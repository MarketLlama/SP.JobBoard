import * as React from 'react';
import styles from './JobBoard.module.scss';
import { IWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";
import pnp, {sp , Web, Site}  from '@pnp/pnpjs';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { PrimaryButton, ActionButton } from 'office-ui-fabric-react/lib/Button';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DayPickerStrings} from '../global/IDatePickerStrings';
import {Tinymce} from '../global/Tinymce';
import { DirectionalHint } from 'office-ui-fabric-react/lib/Callout';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export interface JobSubmissionFromProps {
  context : WebPartContext;
  showForm : boolean;
}

export interface JobSubmissionFromState {
  showModal : boolean;
  file : File;
  deadline? : Date | null;
  firstDayOfWeek?: DayOfWeek;
  jobLevels : any[];
}

class JobSubmissionFrom extends React.Component<JobSubmissionFromProps, JobSubmissionFromState> {
  constructor(props: JobSubmissionFromProps) {
    super(props);
    this.state = {
     showModal : false,
     file : null,
     deadline : null,
     jobLevels : [],
     firstDayOfWeek : DayOfWeek.Sunday
    };
  }

  public componentDidMount() {
    this._getJobLevels();
  }

  //WARNING! To be deprecated in React v17. Use new lifecycle static getDerivedStateFromProps instead.
  public componentWillReceiveProps(nextProps : JobSubmissionFromProps) {
    this.setState({
      showModal : nextProps.showForm
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
                  <TextField label="Job Title " required={true}/>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <DatePicker
                    label = "Deadline Date"
                    isRequired={true}
                    firstDayOfWeek={firstDayOfWeek}
                    strings={DayPickerStrings}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    onSelectDate={this._onDateSelected}
                    value={deadline!}
                  />
                </div>
              </div>
              <br/>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <PeoplePicker
                    context={this.props.context}
                    titleText="Manager"
                    showtooltip={true}
                    tooltipDirectional = {DirectionalHint.topCenter}
                    tooltipMessage = "Surname first to search"
                    personSelectionLimit={1}
                    groupName={""} // IT Leadership
                    isRequired={true}
                    selectedItems={this._getPeoplePickerItems}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <Dropdown
                  placeholder="Select a Job Level"
                  label="Job Level"
                  options={this.state.jobLevels}
                  required={true}/>
                </div>
              </div>
              <br/>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <TaxonomyPicker
                      allowMultipleSelections={true}
                      termsetNameOrID="IT Job tags"
                      panelTitle="Select Tag"
                      label="Job Tags"
                      context={this.props.context}
                      onChange={this._onTaxPickerChange}
                      isTermSetSelectable={false}
                    />
                </div>
              </div>
              <br/>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <p>Role Description</p>
                  <Tinymce onChange={this._handleEditorChange} />
                </div>
              </div>
            </div>
            <br/>
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
                  <span className={styles.fileName}>{this.state.file? this.state.file.name : ''}</span>
                </div>
              </div>
            </div>
          </div>
        </Modal>
    );
  }

  private _closeModal = () =>{
    this.setState({
      showModal : false
    });
  }

  private _handleFile = (files : FileList)=>{
    this.setState({
      file : files[0]
    });
  }

  private _onDateSelected = () =>{

  }

  private _getPeoplePickerItems = () => {

  }

  private _onTaxPickerChange = () => {

  }

  public _handleEditorChange = (e) => {
    console.log('Content was updated:', e.target.getContent());
  }

  private _getJobLevels = () => {
    sp.web.lists.getByTitle('Jobs').fields.getByTitle('Job Level').get().then(field => {
      let choices : any[] = field.Choices;
      let jobLevels : IDropdownOption[] = [];
      choices.forEach(choice => {
        jobLevels.push({
          key: choice,
          text : choice
        });
      });
      this.setState({
        jobLevels : jobLevels
      });
    }, error =>{
      console.log(error);
    });
  }
}

export default JobSubmissionFrom;
