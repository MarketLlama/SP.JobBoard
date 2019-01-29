import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { sp, Web } from '@pnp/pnpjs';
import CVSGenerator from '../global/CSVGenerator';
import { DefaultButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import JobApplicationView from './JobApplicationView';
import * as moment from 'moment';

const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px'
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden'
      }
    }
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px'
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px'
  },
  exampleChild: {
    display: 'block',
    marginBottom: '10px'
  }
});


export interface IJobApplicationsViewProps {
  context : WebPartContext;
}

export interface IJobApplicationsViewState {
  columns: IColumn[];
  items: IJobApplication[];
  selectionDetails: string;
  showApplicationPanel : boolean;
  isModalSelection: boolean;
  isCompactMode: boolean;
  selectedItemId? : number;
}

export interface IJobApplication {
  ID: number;
  Title: string;
  Cover_x0020_Note: string;
  Manager_Name: string;
  Job_Title: string;
  Location: string;
  JobId: number;
  Created : Date;
  Deadline : Date;
}



export class JobApplicationsView extends React.Component<IJobApplicationsViewProps, IJobApplicationsViewState> {
  private _selection: Selection;
  private _items: IJobApplication[];
  private _web = new Web(this.props.context.pageContext.web.absoluteUrl);

  constructor(props: IJobApplicationsViewProps) {
    super(props);

    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'Job Title',
        fieldName: 'JobTitle',
        minWidth: 150,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      },
      {
        key: 'column2',
        name: 'Details',
        fieldName: 'Detail',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      },
      {
        key: 'column3',
        name: 'Date Applied',
        fieldName: 'ApplicationDate',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        data: 'number',
        onRender: (item: any) => {
          return <span>{moment(item.ApplicationDate).format('DD/MM/YYYY')}</span>;
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'Manager',
        fieldName: 'ManagerName',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: any) => {
          return <span>{item.ManagerName}</span>;
        },
        isPadded: true
      },
      {
        key: 'column5',
        name: 'Job Location',
        fieldName: 'JobLocation',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'number',
        onColumnClick: this._onColumnClick,
        onRender: (item: any) => {
          return <span>{item.JobLocation}</span>;
        }
      },
      {
        key: 'column6',
        name: ' ',
        fieldName: ' ',
        minWidth: 20,
        maxWidth: 50,
        isResizable: true,
        onRender: (item: any) => {
          return <span><IconButton iconProps={{ iconName: 'OpenInNewWindow' }} onClick={ () => this._setSelectedItem(item)}
          title="OpenInNewWindow" ariaLabel="OpenInNewWindow" /></span>;
        }
      }
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails()
        });
      }
    });

    this.state = {
      items: [],
      columns: columns,
      selectionDetails: this._getSelectionDetails(),
      isModalSelection: false,
      isCompactMode: false,
      showApplicationPanel : false
    };
  }

  public render() {
    const { columns, isCompactMode, items, selectionDetails, isModalSelection } = this.state;

    return (
      <div>
        <DefaultButton
            text="Export to CSV"
            onClick={this._exportToCSV}
            iconProps={{ iconName: 'ExcelLogo16' }}
            style={{backgroundColor : '#007c45', color:'white',margin : '10px', marginLeft : '20px'}}
          />
        <TextField className={classNames.exampleChild} label="Filter by Job Title:" onChange={this._onChangeText.bind(this)} />
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={this.state.items}
            compact={isCompactMode}
            columns={columns}
            selectionMode={isModalSelection ? SelectionMode.multiple : SelectionMode.none}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            onItemInvoked={this._onItemInvoked}
            enterModalSelectionOnTouch={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          />
        </MarqueeSelection>
        <JobApplicationView context={this.props.context} parent={this} jobId={this.state.selectedItemId}/>
      </div>
    );
  }

  public componentDidUpdate(previousProps: any, previousState: IJobApplicationsViewState) {
    if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
      this._selection.setAllSelected(false);
    }
  }

  private _onChangeCompactMode = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
    this.setState({ isCompactMode: checked });
  }

  private _onChangeModalSelection = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
    this.setState({ isModalSelection: checked });
  }

  private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({ items: text ? this._items.filter(i => i.Job_Title.toLowerCase().indexOf(text) > -1) : this._items });
  }

  private _onItemInvoked(item: any): void {
    alert(`Item invoked: ${item.name}`);
  }

  public componentDidMount() {
    this._getItems();
  }

  private _exportToCSV = () =>{
    let exporter= new CVSGenerator();
    exporter.generateCSV(this.state.items);
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IJobApplication).Job_Title;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this._copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
  }


  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }


  private _getItems = async () => {

    try {
      let items: any = await this._web.lists.getByTitle('Job Applications').items
        .select('Id', 'Cover_x0020_Note', 'Title', 'Created', 'Job/Manager_x0020_Name', 'Job/Title', 'Job/Location', 'Job/Deadline').expand('Job').get();

      let flatItems = [];
      items.forEach(item => {
        flatItems.push({
          Detail: item.Title,
          Id: item.Id,
          CoverNote: item.Cover_x0020_Note,
          ManagerName: item.Job.Manager_x0020_Name,
          JobTitle: item.Job.Title,
          JobLocation: item.Job.Location,
          JobDeadline: item.Job.Deadline,
          ApplicationDate: item.Created
        });
      });

      this.setState({
        items: flatItems
      });
    } catch (err) {
      console.log(err);
    }
  }

  private _setSelectedItem = (item) =>{
    this.setState({
      selectedItemId : item.Id,
      showApplicationPanel : true
    });
  }
}
export default JobApplicationsView;
