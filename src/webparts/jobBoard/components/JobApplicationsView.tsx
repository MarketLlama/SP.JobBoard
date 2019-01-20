import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import * as CamlBuilder from 'camljs';
import { CamlQuery, sp } from '@pnp/pnpjs';
import CVSGenerator from '../global/CSVGenerator';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';

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

const fileIcons: { name: string }[] = [
  { name: 'accdb' },
  { name: 'csv' },
  { name: 'docx' },
  { name: 'dotx' },
  { name: 'mpt' },
  { name: 'odt' },
  { name: 'one' },
  { name: 'onepkg' },
  { name: 'onetoc' },
  { name: 'pptx' },
  { name: 'pub' },
  { name: 'vsdx' },
  { name: 'xls' },
  { name: 'xlsx' },
  { name: 'xsn' }
];

export interface IJobApplicationsViewState {
  columns: IColumn[];
  items: IJobApplication[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
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



export class JobApplicationsView extends React.Component<{}, IJobApplicationsViewState> {
  private _selection: Selection;
  private _items: IJobApplication[];

  constructor(props: {}) {
    super(props);

    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'Name',
        fieldName: 'Job_Title',
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
        key: 'column2',
        name: 'Name',
        fieldName: 'Title',
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
        fieldName: 'Created',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        data: 'number',
        onRender: (item: IJobApplication) => {
          return <span>{item.Created}</span>;
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'Manager',
        fieldName: 'Manager_Name',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IJobApplication) => {
          return <span>{item.Manager_Name}</span>;
        },
        isPadded: true
      },
      {
        key: 'column5',
        name: 'Location',
        fieldName: 'Location',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'number',
        onColumnClick: this._onColumnClick,
        onRender: (item: IJobApplication) => {
          return <span>{item.Location}</span>;
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
      isCompactMode: false
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
            style={{backgroundColor : '#007c45', color:'#ffff'}}
          />
        <TextField className={classNames.exampleChild} label="Filter by name:" onChange={this._onChangeText.bind(this)} />
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
    exporter.generateCSV([]);
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
    let caml = this._buildCAML();
    let items: any = await sp.web.lists.getByTitle('Job Applications').getItemsByCAMLQuery(caml);
    this.setState({
      items: items
    });
  }

  private _buildCAML = (): CamlQuery => {
    let query = new CamlBuilder().View(["ID", "Title", "Cover_x0020_Note"
      , "Manager_Name", "Job_Title", "Created", "Deadline"])
      .LeftJoin("Job", "Id")
      .Select("ID", "JobId")
      .Select("Manager_x0020_Name", "Manager_Name")
      .Select("Title", "Job_Title")
      .Select("Location", "Location")
      .Select("Deadline", "Deadline")
      .Query().ToString();

    const caml: CamlQuery = {
      ViewXml: query,
    };

    return caml;
  }

}
export default JobApplicationsView;
