import * as React from 'react';
import styles from './RequestList.module.scss';
import { IRequestListProps } from './IRequestListProps';
import { IRequestListState } from './IRequestListState';
import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  Panel,
  PanelType,
  TextField,
  Label,
  extendComponent,
  DetailsList,
  SelectionMode,
  IColumn,
  IDetailsListProps,
  IDetailsRowStyles,
  DetailsRow
} from 'office-ui-fabric-react';
import { EnvironmentType } from '@microsoft/sp-core-library';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { IEventData } from '../../services/IEventData';
import { IUserPermissions } from '../../services/IUserPermissions';
import {
  DatePicker,
  DayOfWeek,
  IDatePickerStrings,
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
  DefaultButton,
  PrimaryButton,
  IPersonaProps,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Dialog,
  DialogType,
  DialogFooter,
  Toggle

}
  from 'office-ui-fabric-react';
import { addMonths, addYears } from 'office-ui-fabric-react/lib/utilities/dateMath/DateMath';
import { _ComponentBaseKillSwitches } from '@microsoft/sp-component-base';
import { EditorState, convertToRaw, ContentState } from 'draft-js';
import draftToHtml from 'draftjs-to-html';
import htmlToDraft from 'html-to-draftjs';
import 'react-draft-wysiwyg/dist/react-draft-wysiwyg.css';
import spservices from '../../services/spservices';


function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}

export class RequestList extends React.Component<IRequestListProps, IRequestListState> {
  private spService: spservices = null;
  private attendees: IPersonaProps[] = [];
  private managers: IPersonaProps[] = [];

  private categoryDropdownOption: IDropdownOption[] = [];

  public constructor(props) {
    super(props);

    this.state = {
        showPanel: false,
        siteRegionalSettings: undefined,
        isloading: false,
        hasError: false,
        displayDialog: false,
        errorMessage: "",
        eventData: [],
        columns: [] = [
            {
                key: 'title',
                name: 'Requestor',
                fieldName: 'ownerName',
                minWidth: 200,
                maxWidth: 200,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
            },            
            {
                key: 'Category',
                name: 'Category',
                fieldName: 'Category',
                minWidth: 100,
                maxWidth: 100,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'number',
                isPadded: true,
            },
            {
                key: 'Date',
                name: 'Date',
                fieldName: 'start',
                minWidth: 160,
                maxWidth: 160,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'number',        
                onRender: (item: IEventData) => {
                    return <span>{item.start.toDateString()}</span>;
                  },
                isPadded: true,
            },
            {
                key: 'Manager',
                name: 'Manager',
                fieldName: 'managerName',
                minWidth: 210,
                maxWidth: 300,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,     
                onRender: (item: IEventData) => {
                    return <span>{ item.managerApproved? item.managerName + " (Approved)": item.Status === "Pending" ? item.managerName + " (Pending)" : item.managerName + " (Rejected)"}</span>;
                  },
            },
            {
                key: 'Backup',
                name: 'Backup',
                fieldName: 'backupName',
                minWidth: 210,
                maxWidth: 300,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,     
                onRender: (item: IEventData) => {
                    return <span>{item.backupApproved? item.backupName + " (Accepted)": item.Status === "Pending" ? item.backupName + " (Pending)" : item.backupName + " (Declined)"}</span>;
                  },
            },
            {
                key: 'Status',
                name: 'Status',
                fieldName: 'Status',
                minWidth: 100,
                maxWidth: 100,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'number',
                isPadded: true,
            },
          ]
    };
    // local copia of props
    this.onRenderFooterContent = this.onRenderFooterContent.bind(this);
    this.loadEvents = this.loadEvents.bind(this);
    this.hidePanel = this.hidePanel.bind(this);
    this.closeDialog = this.closeDialog.bind(this);
    //this.enableSave = this.enableSave.bind(this);
    this.spService = new spservices(this.props.context);
  }

  private _onRenderRow: IDetailsListProps['onRenderRow'] = props => {
    const customStyles: Partial<IDetailsRowStyles> = {};
    console.log(props.item)
    if (props) {
        switch(props.item.Status){
            case("Pending"):
                customStyles.root = { backgroundColor: '#f4f5c4' };
              break;
            case("Rejected"):
                customStyles.root = { backgroundColor: '#f0c7c2' };
            break;
            case("Accepted"):
                customStyles.root = { backgroundColor: '#e3ffdb' };
            break;
            case("Cancelled"):
                customStyles.root = { backgroundColor: '#d9d9d9' };
            break;              
        }

      return <DetailsRow {...props} styles={customStyles} />;
    }
    return null;
  };

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, eventData } = this.state;
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
    const newItems : IEventData[] = _copyAndSort(eventData, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      eventData: newItems,
    });
  };
  /**
   *  Hide Panel
   *
   * @private
   * @memberof RequestList
   */
  private hidePanel() {
    this.props.onDissmissPanel(false);
  }
 
  /**
   *
   * @param {*} error
   * @param {*} errorInfo
   * @memberof RequestList
   */
  public componentDidCatch(error: any, errorInfo: any) {
    this.setState({ hasError: true, errorMessage: errorInfo.componentStack });
  }  
  
  /**
   * @private
   * @memberof Calendar
   */
  private async loadEvents() {
    try {
      // Teste Properties
      if (!this.props.list || !this.props.siteUrl || !this.props.eventStartDate.value || !this.props.eventEndDate.value) return;

      const eventsData: IEventData[] = await this.spService.getUserEvents(escape(this.props.siteUrl), escape(this.props.list), this.props.eventStartDate.value, this.props.eventEndDate.value);

      for (const event of eventsData){ 
        //correct issue with items listed as all day events
        if (event.allDayEvent){
        event.end.setTime(event.end.getTime() + (7*60*60*1000)); 
        event.start.setTime(event.start.getTime() + (7*60*60*1000)); 
        }
      }
      console.log(eventsData)
      this.setState({ eventData: eventsData, hasError: false, errorMessage: "" });
    } catch (error) {
      this.setState({ hasError: true, errorMessage: error.message, isloading: false });
    }
  }

  /**
   *
   *
   * @memberof RequestList
   */
  public async componentDidMount() {
    this.setState({ isloading: true });
    let editorState:EditorState;
    // Load Regional Settings
    const siteRegionalSettigns = await this.spService.getSiteRegionalSettingsTimeZone(this.props.siteUrl);
    // chaeck User list Permissions
    const userListPermissions: IUserPermissions = await this.spService.getUserPermissions(this.props.siteUrl, this.props.listId);
    await this.loadEvents();
    this.setState({ isloading: false });
  }

  /**
   *
   * @memberof RequestList
   */
  public componentWillMount() {

  }

  /**
   *
   * @private
   * @param {React.MouseEvent<HTMLDivElement>} event
   * @memberof RequestList
   */
  private closeDialog(ev: React.MouseEvent<HTMLDivElement>) {
    ev.preventDefault();
    this.setState({ displayDialog: false });
  }

 
  /**
   * @private
   * @returns
   * @memberof RequestList
   */
  private onRenderFooterContent() {
    return (
      <div >
        <DefaultButton onClick={this.hidePanel} style={{ marginBottom: '15px', float: 'right' }}>
          {"Cancel"}
        </DefaultButton>
      </div>
    );
  }
  
  public render(): React.ReactElement<IRequestListProps> {
    const { columns} = this.state;
    return (
      <div>
        <Panel
          isOpen={this.props.showPanel}
          onDismiss={this.hidePanel}
          type={PanelType.large}
          headerText={"My Requests"}
          isFooterAtBottom={true}
          onRenderFooterContent={this.onRenderFooterContent}
        >
          <div style={{ width: '100%' }}>
            {
              this.state.hasError &&
              <MessageBar messageBarType={MessageBarType.error}>
                {this.state.errorMessage}
              </MessageBar>
            }
            {
              this.state.isloading && (
                <Spinner size={SpinnerSize.large} />
              )
            }
            {
              !this.state.isloading &&
              <div>
                  <DetailsList items={this.state.eventData} 
                    onRenderRow={this._onRenderRow}
                    columns={columns}
                    selectionMode={SelectionMode.none} />
              </div>
            }
          </div>
        </Panel>
      </div>
    );
  }
}
