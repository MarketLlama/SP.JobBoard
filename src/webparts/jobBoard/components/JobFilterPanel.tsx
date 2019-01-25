import * as React from 'react';
import styles from './JobBoard.module.scss';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { ChoiceGroup } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import {
  Accordion,
  AccordionItem,
  AccordionItemTitle,
  AccordionItemBody,
} from 'react-accessible-accordion';
import {
  taxonomy,
  ITermStore,
  ITerm,
  ITermData,
  ITermStoreData,
  ITermSet
} from "@pnp/sp-taxonomy";
import 'react-accessible-accordion/dist/fancy-example.css';
import JobBoard from './JobBoard';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Web , CamlQuery} from '@pnp/sp';
import { Checkbox } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DirectionalHint } from 'office-ui-fabric-react/lib/Tooltip';
import * as CamlBuilder from 'camljs';

export interface JobFilterPanelProps {
  showPanel: boolean;
  parent: JobBoard;
  context: WebPartContext;
}

export interface JobFilterPanelState {
  showPanel: boolean;
  jobLevels?: any[];
  jobTags?: any[];
}

export interface IFilterItem {
  label: string;
  section: string;
}

class JobFilterPanel extends React.Component<JobFilterPanelProps, JobFilterPanelState> {
  private _web = new Web(this.props.context.pageContext.web.absoluteUrl);
  private _filterArray: Array<IFilterItem> = new Array<IFilterItem>();
  private _jobIds : Array<number> = new Array<number>();

  constructor(props: JobFilterPanelProps) {
    super(props);
    this.state = {
      showPanel: false,
      jobTags: [],
      jobLevels: []
    };
  }

  public render(): JSX.Element {
    return (
      <div>
        <Panel
          isOpen={this.state.showPanel}
          type={PanelType.medium}
          onDismiss={this._onClosePanel}
          headerText="Filters"
          closeButtonAriaLabel="Close"
          onRenderFooterContent={this._onRenderFooterContent}
          className={styles.filterPanel}
        >
          <Accordion>
            <AccordionItem>
              <AccordionItemTitle className={styles.accordionTitle}>
                <h4>Job Level</h4>
              </AccordionItemTitle>
              <AccordionItemBody>
                {this.state.jobLevels}
              </AccordionItemBody>
            </AccordionItem>
            <AccordionItem>
              <AccordionItemTitle className={styles.accordionTitle}>
                <h4>Manager </h4>
              </AccordionItemTitle>
              <AccordionItemBody>
                <PeoplePicker
                  context={this.props.context}
                  titleText="Manager"
                  showtooltip={true}
                  tooltipDirectional={DirectionalHint.topCenter}
                  tooltipMessage="Surname first to search"
                  personSelectionLimit={3}
                  groupName={""} // IT Leadership
                  isRequired={true}
                  selectedItems={this._setManagers}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
              </AccordionItemBody>
            </AccordionItem>
          </Accordion>
        </Panel>
      </div>
    );
  }

  private _onClosePanel = (): void => {
    this.setState({ showPanel: false });
    this.props.parent.setState({
      showFilter: false
    });
  }

  public componentDidMount() {
    this._getJobLevels();
    //this._getJobTags();
  }

  //WARNING! To be deprecated in React v17. Use new lifecycle static getDerivedStateFromProps instead.
  public componentWillReceiveProps(nextProps: JobFilterPanelProps) {
    this.setState({
      showPanel: nextProps.showPanel
    });
  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={this._onClosePanel} style={{ marginRight: '8px' }}>
          Save
        </PrimaryButton>
        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  }

  private _getJobLevels = () => {
    let jobLevels = [];
    this._web.lists.getByTitle('Jobs').fields.getByTitle('Job Level').get().then(field => {
      let choices: any[] = field.Choices;
      choices.forEach(choice => {
        jobLevels.push(
          <Checkbox
            label={choice}
            onChange={this._onLevelCheckboxChange}
            className={styles.checkboxes}
            value={choice}
          />
        );
      });
      console.log(jobLevels);
      this.setState({
        jobLevels: jobLevels
      });
    }, error => {
      console.log(error);
    });
  }

  private _getJobTags = async () => {
    var termSetName = 'IT Job tags';
    var options = [];
    let store: (ITermStoreData & ITermStore)[] = await taxonomy.termStores.get();
    let termSet: ITermSet = await store[0].getTermSetsByName(termSetName, 1033).getByName(termSetName).get();
    let terms: (ITermData & ITerm)[] = await termSet.terms.get();
    let regExp = /\(([^)]+)\)/;
    for (let term of terms) {
      options.push(
        <Checkbox
          label={term.Name}
          onChange={this._onTagCheckboxChange}
          value={term.Name}
          className={styles.checkboxes}
        />
      );
    }
    this.setState({
      jobTags: options
    });
    return options;
  }

  private _setManagers = () => {
    this.setState({

    });
  }

  private _addToFilterArray = (item: IFilterItem) => {
    this._filterArray.push(item);
    let caml = this._buildCAML(["ID","Manager", "Created"]);
    this._getJobIdForFiltering(caml);
  }

  private _removeFromFilterArray = (item: IFilterItem) => {
    this._filterArray = this._filterArray.filter(i => {
      return (i.label !== item.label && i.section !== item.section);
    });
    let caml = this._buildCAML(["ID","Manager", "Created"]);
    this._getJobIdForFiltering(caml);
  }

  private _buildCAML = (viewFields: string[]) : string  => {
    let camlQuery = new CamlBuilder().View(viewFields).Query();
    if (this._filterArray.length > 0) {
      let xml = camlQuery.Where().DateTimeField("Deadline").GreaterThan(new Date()).OrderByDesc("Created").ToString();
      let query = CamlBuilder.FromXml(xml).ModifyWhere().AppendOr();
      this._filterArray.forEach(item => {
        switch (item.section) {
          case 'Manager':
            query.UserField("Manager").Id().EqualTo(6);
            break;
          case 'Level':
            query.ChoiceField("Job_x0020_Level").Contains(item.label);
            break;
          case 'Tag':
            query.LookupMultiField("JobTags").Includes(item.label);
            break;
          default:
            break;
        }
      });
      return query.DateField("Created").LessThanOrEqualTo(new Date()).ToString();
    } else {
      return camlQuery.Where()
        .DateTimeField("Deadline")
        .GreaterThan(new Date())
        .OrderByDesc("Created").ToString();
    }
  }

  private _getJobIdForFiltering = async (caml: string) => {
    try {
      let jobs = await this._web.lists.getByTitle('Jobs').select('Id').getItemsByCAMLQuery({
        ViewXml: caml,
      });
      console.log(jobs);
    } catch (error) {
      console.log(error);
    }
  }

  private _onLevelCheckboxChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
    let label = '';
    label = ev.currentTarget.children[0].textContent.replace(/([^A-Za-z0-9 _])+/, '');
    console.log(label);
    if (checked) {
      this._addToFilterArray({
        label: label,
        section: 'Level'
      });
    } else {
      this._removeFromFilterArray({
        label: label,
        section: 'Level'
      });
    }
  }

  private _onTagCheckboxChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
    let label = '';
    label = ev.currentTarget.children[0].textContent.replace(/([^A-Za-z0-9 _])+/, '');
    console.log(label);
    if (checked) {
      this._addToFilterArray({
        label: label,
        section: 'Tag'
      });
    } else {
      this._removeFromFilterArray({
        label: label,
        section: 'Tag'
      });
    }
  }

  private _onShowPanel = (): void => {
    this.setState({ showPanel: true });
  }
}

export default JobFilterPanel;

