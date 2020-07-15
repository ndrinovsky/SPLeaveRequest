import { IEventData } from '../../services/IEventData';
import { IUserPermissions } from '../../services/IUserPermissions';
import { DayOfWeek} from 'office-ui-fabric-react/lib/DatePicker';
import {  IDropdownOption } from 'office-ui-fabric-react/';
import { IColumn } from 'office-ui-fabric-react';
export interface IRequestListState {
  showPanel: boolean;
  isloading:boolean;
  siteRegionalSettings: any;
  hasError: boolean;
  displayDialog: false;
  errorMessage: string;
  eventData:  IEventData[];
  columns: IColumn[];
}
