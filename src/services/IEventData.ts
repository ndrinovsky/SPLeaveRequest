export interface IEventData {
  id?:number;
  title: string;
  Description?: any;
  start: Date;
  end: Date;
  color?:string;
  ownerInitial?: string;
  ownerPhoto?:string;
  ownerEmail?:string;
  ownerName?:string;
  allDayEvent?: boolean;
  backup?: number;
  backupName?: string;
  backupApproved: boolean;
  manager?: number;
  managerName?: string;
  managerApproved: boolean;
  Category?: string;
}
