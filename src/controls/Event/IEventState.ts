import { IEventData } from '../../services/IEventData';
import { IUserPermissions } from '../../services/IUserPermissions';
import { DayOfWeek} from 'office-ui-fabric-react/lib/DatePicker';
import {  IDropdownOption } from 'office-ui-fabric-react/';
export interface IEventState {
  showPanel: boolean;
  eventData:IEventData;
  firstDayOfWeek?: DayOfWeek;
  startSelectedHour: IDropdownOption ;
  startSelectedMin: IDropdownOption ;
  endSelectedHour: IDropdownOption ;
  endSelectedMin: IDropdownOption ;
  allDayEventState?: boolean;
  startDate?: Date;
  endDate?: Date;
  editorState?: any;
  selectedUsers: string[];
  managers?: string[];
  errorMessage?:string;
  hasError?:boolean;
  disableButton?: boolean;
  isSaving?:boolean;
  isDeleting?:boolean;
  displayDialog:boolean;
  userPermissions?: IUserPermissions;
  isloading:boolean;
  siteRegionalSettings: any;
  //noManagerRequired: boolean;
}
