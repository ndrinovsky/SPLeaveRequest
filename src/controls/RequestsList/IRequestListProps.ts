import { IEventData } from '../../services/IEventData';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
export interface IRequestListProps {
  onDissmissPanel: (refresh:boolean) => void;
  showPanel: boolean;
  context:WebPartContext;
  siteUrl: string;
  listId:string;
  list: string;  
  eventStartDate:  IDateTimeFieldValue;
  eventEndDate: IDateTimeFieldValue;
}