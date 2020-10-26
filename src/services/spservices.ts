import { WebPartContext } from "@microsoft/sp-webpart-base";
import { graph, } from "@pnp/graph";
import { IEventData } from './IEventData';
import * as moment from 'moment';
import * as moment_timezone from 'moment-timezone';
import { IUserPermissions } from './IUserPermissions';
import {sp, SiteUsers, SiteUser } from "@pnp/sp/presets/all";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { event } from "jquery";
import { Web, PermissionKind, RegionalSettings } from "@pnp/pnpjs";

moment.utc();
moment_timezone.tz.setDefault("Etc/UTC");

const ADMIN_ROLETEMPLATE_ID = "62e90394-69f5-4237-9190-012177145e10"; // Global Admin TemplateRoleId
// Class Services
export default class spservices {

  constructor(private context: WebPartContext) {
    // Setuo Context to PnPjs and MSGraph
    sp.setup({
      spfxContext: this.context,
      ie11 :true
    });
    
    graph.setup({
      spfxContext: this.context
    });
    // Init
    this.onInit();
  }
  // OnInit Function
  private async onInit() {
    //this.appCatalogUrl = await this.getAppCatalogUrl();

  }

  /**
   *
   * @private
   * @param {string} siteUrl
   * @returns {Promise<number>}
   * @memberof spservices
   */
  private async getSiteTimeZoneHoursToUtc(siteUrl: string): Promise<number> {
    let numberHours: number = 0;
    let siteTimeZoneHoursToUTC: any;
    let siteTimeZoneBias: number;
    let siteTimeZoneDaylightBias: number;
    let currentDateTimeOffSet: number = new Date().getTimezoneOffset() / 60;

    try {
      const siteRegionalSettings: any = await this.getSiteRegionalSettingsTimeZone(siteUrl);
      // Calculate  hour to current site
      siteTimeZoneBias = siteRegionalSettings.Information.Bias;
      siteTimeZoneDaylightBias = siteRegionalSettings.Information.DaylightBias;

      // Formula to calculate the number of  hours need to get UTC Date.
      // numberHours = (siteTimeZoneBias / 60) + (siteTimeZoneDaylightBias / 60) - currentDateTimeOffSet;
      if ( siteTimeZoneBias >= 0 ){
        numberHours = ((siteTimeZoneBias / 60) - currentDateTimeOffSet) + siteTimeZoneDaylightBias/60 ;
      }else {
        numberHours = ((siteTimeZoneBias / 60) - currentDateTimeOffSet)  ;
      }
    }
    catch (error) {
      return Promise.reject(error);
    }
    return numberHours;
  }

  /**
   *
   * @param {IEventData} newEvent
   * @param {string} siteUrl
   * @param {string} listId
   * @returns
   * @memberof spservices
   */
  public async addEvent(newEvent: IEventData, siteUrl: string, listId: string) {
    let results = null;
    try {
      const web = new Web(siteUrl);

      const siteTimeZoneHoursToUTC: number = await this.getSiteTimeZoneHoursToUtc(siteUrl);
      //"Title","fRecurrence", "fAllDayEvent","EventDate", "EndDate", "Description","ID",
      //"Backup"
      const startDate = new Date(moment(newEvent.start).add(siteTimeZoneHoursToUTC, 'hours').toISOString());
      const endDate = new Date(moment(newEvent.end).add(siteTimeZoneHoursToUTC, 'hours').toISOString());
      
      let requestor = await web.siteUsers.getByEmail(this.context.pageContext.user.loginName).get();
      results = await web.lists.getById(listId).items.add({
        Title:  requestor.Title + " - " + newEvent.Category,
        Description: newEvent.Description,
        ParticipantsPickerId: newEvent.backup,
        EventDate: startDate,
        EndDate: endDate,
        fAllDayEvent: newEvent.allDayEvent,
        fRecurrence: false,
        Category: newEvent.Category,
        ManagerId: newEvent.manager,
        RequestorId: requestor.Id
      });
    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }
  /**
   *
   * @param {IEventData} newEvent
   * @param {string} siteUrl
   * @param {string} listId
   * @returns
   * @memberof spservices
   */
  public async updateEvent(updateEvent: IEventData, siteUrl: string, listId: string) {
    let results = null;
    try {

      const siteTimeZoneHoursToUTC: number = await this.getSiteTimeZoneHoursToUtc(siteUrl);

      const web = new Web(siteUrl);

      const startDate = new Date(moment(updateEvent.start).add(siteTimeZoneHoursToUTC, 'hours').toISOString());
      const endDate = new Date(moment(updateEvent.end).add(siteTimeZoneHoursToUTC, 'hours').toISOString());
      let requestor = await web.siteUsers.getByEmail(this.context.pageContext.user.loginName).get();
      //"Title","fRecurrence", "fAllDayEvent","EventDate", "EndDate", "Description","ID", "Location","ParticipantsPickerId"
      results = await web.lists.getById(listId).items.getById(updateEvent.id).update({
          Title: requestor.Title + " - " + updateEvent.Category,
          Description: updateEvent.Description,
          ParticipantsPickerId: updateEvent.backup,
          EventDate: startDate,
          EndDate: endDate,
          fAllDayEvent:  updateEvent.allDayEvent,
          fRecurrence: false,
          Category: updateEvent.Category,
          ManagerId: updateEvent.manager
      });
    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }

  /**
   *
   * @param {IEventData} event
   * @param {string} siteUrl
   * @param {string} listId
   * @returns
   * @memberof spservices
   */
  public async deleteEvent(event: IEventData, siteUrl: string, listId: string) {
    let results = null;
    try {
      const web = new Web(siteUrl);

      //"Title","fRecurrence", "fAllDayEvent","EventDate", "EndDate", "Description","ID", "Location","ParticipantsPickerId"
      results = await web.lists.getById(listId).items.getById(event.id).update({
        Status: "Cancelled"
      });
    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }

  /**
   *
   * @param {number} userId
   * @param {string} siteUrl
   * @returns {Promise<SiteUser>}
   * @memberof spservices
   */
  public async getUserById(userId: number, siteUrl: string): Promise<typeof SiteUser> {
    let results: typeof SiteUser = null;

    if (!userId && !siteUrl) {
      return null;
    }

    try {
      const web = new Web(siteUrl);
      results = await web.siteUsers.getById(userId).get();
      //results = await web.siteUsers.getByLoginName(userId).get();
    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }

  /**
   *
   *
   * @param {string} loginName
   * @param {string} siteUrl
   * @returns {Promise<SiteUser>}
   * @memberof spservices
   */
  public async getUserByLoginName(loginName: string, siteUrl: string): Promise<typeof SiteUser> {
    let results: typeof SiteUser = null;

    if (!loginName && !siteUrl) {
      return null;
    }

    try {
      const web = new Web(siteUrl);
      await web.ensureUser(loginName);
      results = await web.siteUsers.getByLoginName(loginName).get();
      //results = await web.siteUsers.getByLoginName(userId).get();
    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }
  /**
   *
   * @param {string} loginName
   * @returns
   * @memberof spservices
   */
  public async getUserProfilePictureUrl(loginName: string) {
    let results: any = null;
    try {
      results = await sp.profiles.getPropertiesFor(loginName);
    } catch (error) {
      results = null;
    }
    return results.PictureUrl;
  }

  /**
   *
   * @param {string} siteUrl
   * @param {string} listId
   * @returns {Promise<IUserPermissions>}
   * @memberof spservices
   */
  public async getUserPermissions(siteUrl: string, listId: string): Promise<IUserPermissions> {
    let hasPermissionAdd: boolean = false;
    let hasPermissionEdit: boolean = false;
    let hasPermissionDelete: boolean = false;
    let hasPermissionView: boolean = false;
    let userPermissions: IUserPermissions = undefined;
    try {
      const web = new Web(siteUrl);
      const  userEffectivePermissions = await web.lists.getById(listId).effectiveBasePermissions.get();
        // chaeck user permissions
        hasPermissionAdd = web.lists.getById(listId).hasPermissions(userEffectivePermissions, PermissionKind.AddListItems);
        hasPermissionEdit = web.lists.getById(listId).hasPermissions(userEffectivePermissions, PermissionKind.EditListItems);
        hasPermissionDelete =web.lists.getById(listId).hasPermissions(userEffectivePermissions, PermissionKind.DeleteListItems);
        hasPermissionView = web.lists.getById(listId).hasPermissions(userEffectivePermissions, PermissionKind.ViewListItems);
        userPermissions = { hasPermissionAdd: hasPermissionAdd, hasPermissionEdit: hasPermissionEdit, hasPermissionDelete: hasPermissionDelete, hasPermissionView: hasPermissionView };
    } catch (error) {
      return Promise.reject(error);
    }
    return userPermissions;
  }
  /**
   *
   * @param {string} siteUrl
   * @returns
   * @memberof spservices
   */
  public async getSiteLists(siteUrl: string) {

    let results: any[] = [];

    if (!siteUrl) {
      return [];
    }

    try {
      const web = new Web(siteUrl);
      results = await web.lists.select("Title", "ID").filter('BaseTemplate eq 106').get();

    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }

  /**
   *
   * @private
   * @returns
   * @memberof spservices
   */
  public async colorGenerate() {

    var hexValues = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e"];
    var newColor = "#";
    "OoO"
    "Arrive late"
    "Leave early"
    "Sick"
    "Training"
    "Vacation"
    for (var i = 0; i < 6; i++) {
      var x = Math.round(Math.random() * 14);
      var y = hexValues[x];
      newColor += y;
    }
    return newColor;
  }
    /**
   *
   * @private
   * @returns
   * @memberof spservices
   */
  public async colorGenerateEvents(event : IEventData) {

    var color : string;
    switch(event.Category){
      case "OoO":
        color="#ff230f";
        break;
      case "Arrive late":
        color="#27b300";
        break;
      case "Leave early":
        color="#448000";
        break;
      case "Sick":
        color="#c7c7c7";
        break;
      case "Training":
        color="#2492ff";
        break;
      case "Vacation":
        color="#ffaf24";
        break;
      default:
        color="#c7c7c7";
    }
    
    return color;
  }
  /**
   *
   * @param {string} siteUrl
   * @returns {Promise< any[]>}
   * @memberof spservices
   */
  public async getUserManagerPrincipalName(siteUrl: string): Promise<string[]> {
  if (!siteUrl) {
    return [];
  }
  try {
    const profile = await sp.profiles.myProperties.get();
    // Properties are stored in Key/Value pairs,
    // so parse into an object called userProperties
    var props = {};
      profile.UserProfileProperties.forEach((prop) => {
      props[prop.Key] = prop.Value;
    });
    profile.userProperties = props;
    var result = [];
    result.push((await sp.web.siteUsers.getByLoginName(profile.userProperties.Manager).get()).UserPrincipalName);
    return result;
  }catch (error) {
    return [];
    //return Promise.reject(error);
  }
}
  /**
   *
   * @param {string} siteUrl
   * @returns {Promise< any[]>}
   * @memberof spservices
   */
  public async getUserManager(siteUrl: string): Promise<string[]> {
    if (!siteUrl) {
      return [];
    }
    try {
      const profile = await sp.profiles.myProperties.get();
      // Properties are stored in Key/Value pairs,
      // so parse into an object called userProperties
      var props = {};
        profile.UserProfileProperties.forEach((prop) => {
        props[prop.Key] = prop.Value;
      });
      profile.userProperties = props;
      var result = [];
      result.push((await sp.web.siteUsers.getByLoginName(profile.userProperties.Manager).get()));
      return result;
    }catch (error) {
      return [];
      //return Promise.reject(error);
    }
  }
  /**
   *
   * @param {string} siteUrl
   * @param {string} listId
   * @param {Date} eventStartDate
   * @param {Date} eventEndDate
   * @returns {Promise< IEventData[]>}
   * @memberof spservices
   */
  public async getUserEvents(siteUrl: string, listId: string, eventStartDate: Date, eventEndDate: Date): Promise<IEventData[]> {
    let events: IEventData[] = [];
    if (!siteUrl) {
      return [];
    }
    try {
      // Get Regional Settings TimeZone Hours to UTC
      const siteTimeZoneHoursToUTC: number = await this.getSiteTimeZoneHoursToUtc(siteUrl);
      // Get Category Field Choices
      const categoryDropdownOption = await this.getChoiceFieldOptions(siteUrl, listId, 'Category');

      const web = new Web(siteUrl);
      const results = await web.lists.getById(listId).renderListDataAsStream(
        {
          DatesInUtc: true,
          ViewXml: `<View><ViewFields><FieldRef Name='Status'/><FieldRef Name='Requestor'/><FieldRef Name='Author'/><FieldRef Name='Manager'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='ManagerApproved'/><FieldRef Name='BackupApproved'/><FieldRef Name='Category'/><FieldRef Name='Description'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='ID'/><FieldRef Name='EndDate'/><FieldRef Name='EventDate'/><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='fAllDayEvent'/></ViewFields>
          <Query>
          <Where>             
            <Or>     
              <Or>            
                <Eq>
                  <FieldRef Name='Manager'/>
                    <Value Type='Text'>${this.context.pageContext.user.displayName}</Value>
                </Eq>                  
                <Eq>
                  <FieldRef Name='Requestor'/>
                    <Value Type='Text'>${this.context.pageContext.user.displayName}</Value>
                </Eq>
              </Or>                  
              <Eq>
                <FieldRef Name='ParticipantsPicker'/>
                  <Value Type='Text'>${this.context.pageContext.user.displayName}</Value>
              </Eq>
            </Or> 
          </Where>
          <OrderBy>
            <FieldRef Name="EventDate"  Ascending = "FALSE"/>
          </OrderBy>
          </Query>
          <RowLimit Paged=\"FALSE\">2000</RowLimit>
          </View>`
        }
      );
      if (results && results.Row.length > 0) {
        for (const event of results.Row) {

          const initialsArray: string[] = event.Requestor[0].title.split(' ');
          //const initials: string = "test";
          const initials: string = initialsArray[0].charAt(0) + initialsArray[initialsArray.length - 1].charAt(0);
          //const userPictureUrl = await this.getUserProfilePictureUrl(`i:0#.f|membership|${event.Requestor[0].email}`);
          const fAllDayEvent: boolean =  (event.fAllDayEvent == "Yes") ? true : false;         

          const backupId: number = ( event.ParticipantsPicker[0] != null) ? event.ParticipantsPicker[0].id : null;    
          //const backupObj : any = (backupId != null) ? (await web.siteUsers.getById(backupId).get()).Title : null;   
          const backupObj : any = (event.ParticipantsPicker[0] != null) ? event.ParticipantsPicker[0].title : null;
          const backupApproved: boolean =  (event.BackupApproved == "Yes") ? true : false;   
          const managerId: number = ( event.Manager[0] != null) ? event.Manager[0].id : null;    
          //const managerObj : any = (managerId != null) ? (await web.siteUsers.getById(managerId).get()).Title : null;      
          const managerObj : any = event.Manager[0].title;
          const managerApproved: boolean =  (event.ManagerApproved == "Yes") ? true : false;
          const startDate =  new Date(moment(event.EventDate).toISOString());
          const endDate = new Date(moment(event.EndDate).toISOString());
          events.push({
            Status: event.Status,
            id: event.ID,
            title: event.Requestor[0].title + " " + event.Category,
            Description: event.Description,
            start: startDate,
            end: endDate,
            ownerEmail: event.Requestor[0].email,
            ownerPhoto: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${event.Requestor[0].email}&UA=0&size=HR96x96`,
            ownerInitial: initials,
            //color: await this.colorGenerate(),
            color: await this.colorGenerateEvents(event),
            ownerName: event.Requestor[0].title,
            backup: backupId,      
            backupName: backupObj,
            backupApproved: backupApproved,      
            manager : managerId,       
            managerName: managerObj,
            managerApproved: managerApproved,   
            allDayEvent: fAllDayEvent,
            Category: event.Category
          });
        }
      }
      // Return Data
      return events;
    } catch (error) {
      console.log(error)
      return Promise.reject(error);
    }
  }

  /**
   *
   * @param {string} siteUrl
   * @param {string} listId
   * @param {string} fieldInternalName
   * @returns {Promise<{ key: string, text: string }[]>}
   * @memberof spservices
   */
  public async getChoiceFieldOptions(siteUrl: string, listId: string, fieldInternalName: string): Promise<{ key: string, text: string }[]> {
    let fieldOptions: { key: string, text: string }[] = [];
    try {
      const web = new Web(siteUrl);
      const results = await web.lists.getById(listId)
        .fields
        .getByInternalNameOrTitle(fieldInternalName)
        .select("Title", "InternalName", "Choices")
        .get();
      if (results && results.Choices.length > 0) {
        for (const option of results.Choices) {
          fieldOptions.push({
            key: option,
            text: option
          });
        }
      }
    } catch (error) {
      return Promise.reject(error);
    }
    return fieldOptions;
  }

  // /**
  //  *
  //  * @param {string} siteUrl
  //  * @param {string} listId
  //  * @param {Date} eventStartDate
  //  * @param {Date} eventEndDate
  //  * @returns {Promise< IEventData[]>}
  //  * @memberof spservices
  //  */
  // public async getEvents2(siteUrl: string, listId: string, eventStartDate: Date, eventEndDate: Date, allowPending: boolean): Promise<IEventData[]> {
  //     let events: IEventData[] = [];
  //     let results;
  //     if (!siteUrl) {
  //       return [];
  //     }
  //     try {
  //       const web = new Web(siteUrl);
  //         // Get Regional Settings TimeZone Hours to UTC
  //       const siteTimeZoneHoursToUTC: number = await this.getSiteTimeZoneHoursToUtc(siteUrl);
  //       // Get Category Field Choices
  //       const categoryDropdownOption = await this.getChoiceFieldOptions(siteUrl, listId, 'Category');
  //       let categoryColor: { category: string, color: string }[] = [];
  //       for (const cat of categoryDropdownOption) {
  //       categoryColor.push({ category: cat.text, color: await this.colorGenerate() });
  //       }
  //       var spRequest = new XMLHttpRequest(); 
  //       //spRequest.open('GET', "/sites/ITD/ApplicationsDiv/_api/web/lists/getbytitle('OoO Request List')/items",false); 
  //     //   spRequest.setRequestHeader("Accept","application/json");                          
  //     //   spRequest.onreadystatechange = function(){             
  //     //     if (spRequest.readyState === 4 && spRequest.status === 200){ 
  //     //       results = JSON.parse(spRequest.responseText); 
  //     //       console.log(results.Row); 
  //     //     } 
  //     //     else if (spRequest.readyState === 4 && spRequest.status !== 200){ 
  //     //         console.log('Error Occurred !'); 
  //     //     } 
  //     // }; 
  //     // spRequest.send();

  //   const restAPI = `${siteUrl}/_api/web/Lists(guid'${listId}')/RenderListDataAsStream`;
  //   this.context.spHttpClient.post(restAPI, SPHttpClient.configurations.v1 , {
  //     body: JSON.stringify({
  //       parameters: {
  //         RenderOptions: 2,
  //         DatesInUtc: true,
  //         ViewXml: `<View><ViewFields><FieldRef Name='Status'/><FieldRef Name='Requestor'/><FieldRef Name='Author'/><FieldRef Name='Manager'/><FieldRef Name='ManagerApproved'/><FieldRef Name='BackupApproved'/><FieldRef Name='Category'/><FieldRef Name='Description'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='ID'/><FieldRef Name='EndDate'/><FieldRef Name='EventDate'/><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='fAllDayEvent'/></ViewFields>
  //         <Query>
  //         <Where>
  //           <And>
  //            <And>
  //             <Geq>
  //               <FieldRef Name='EventDate' />
  //               <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventStartDate).format('YYYY-MM-DD')}</Value>
  //             </Geq>
  //             <Leq>
  //               <FieldRef Name='EventDate' />
  //               <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventEndDate).format('YYYY-MM-DD')}</Value>
  //             </Leq>
  //             </And>
  //             <Eq>
  //             <FieldRef Name='fRecurrence' />
  //               <Value Type='Recurrence'>0</Value>
  //             </Eq>
  //           </And>
  //         </Where>
  //         </Query>
  //         <RowLimit Paged=\"FALSE\">2000</RowLimit>
  //         </View>`
  //       }
  //     })
  //   })
  //   .then((data: SPHttpClientResponse) => data.json())
  //   .then(async (data: any) => {
  //     if (data && data.Row.length > 0) {
  //       for (const event of data.Row) {
  //         console.log(event);
  //         const initialsArray: string[] = event.Requestor[0].title.split(' ');
  //         //const initials: string = "test";
  //         const initials: string = initialsArray[0].charAt(0) + initialsArray[initialsArray.length - 1].charAt(0);
  //         const userPictureUrl = await this.getUserProfilePictureUrl(`i:0#.f|membership|${event.Requestor[0].email}`);
  
         
  //         const fAllDayEvent: boolean =  (event.fAllDayEvent == "Yes") ? true : false;         
  //         const CategoryColorValue: any[] = categoryColor.filter((value) => {
  //           return value.category == event.Category;
  //         });
  
  //         const backupId: number = ( event.ParticipantsPicker[0] != null) ? event.ParticipantsPicker[0].id : null;    
  //         const backupObj : any = (backupId != null) ? (await web.siteUsers.getById(backupId).get()).Title : null;                   
  //         const backupApproved: boolean =  (event.BackupApproved == "Yes") ? true : false;   
  //         const managerId: number = ( event.Manager[0] != null) ? event.Manager[0].id : null;    
  //         const managerObj : any = (managerId != null) ? (await web.siteUsers.getById(managerId).get()).Title : null;                   
  //         const managerApproved: boolean =  (event.ManagerApproved == "Yes") ? true : false;       
  
  //         const startDate =  new Date(moment(event.EventDate).toISOString());
  //         const endDate = new Date(moment(event.EndDate).toISOString());
  //         if ((allowPending || (!allowPending && managerApproved)) && event.Status !== "Cancelled" && event.Status !== "Rejected" )
  //         events.push({
  //           id: event.ID,
  //           title: event.Requestor[0].title + " " + event.Category,
  //           Description: event.Description,
  //           start: startDate,
  //           end: endDate,
  //           ownerEmail: event.Requestor[0].email,
  //           ownerPhoto: userPictureUrl ?
  //             `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${event.Requestor[0].email}&UA=0&size=HR96x96` : '',
  //           ownerInitial: initials,
  //           //color: await this.colorGenerate(),
  //           color: await this.colorGenerateEvents(event),
  //           ownerName: event.Requestor[0].title,
  //           backup: backupId,      
  //           backupName: backupObj,
  //           backupApproved: backupApproved,      
  //           manager : managerId,       
  //           managerName: managerObj,
  //           managerApproved: managerApproved,   
  //           allDayEvent: fAllDayEvent,
  //           Category: event.Category
  //         });
  //       }
  //     }
  //     console.log(events);
  //       return events;
  //   });
  //   } catch (error) {
  //     return Promise.reject(error);
  //   }
  // }
  /**
   *
   * @param {string} siteUrl
   * @param {string} listId
   * @param {Date} eventStartDate
   * @param {Date} eventEndDate
   * @returns {Promise< IEventData[]>}
   * @memberof spservices
   */
  public async getEvents(siteUrl: string, listId: string, eventStartDate: Date, eventEndDate: Date, allowPending: boolean): Promise<IEventData[]> {

    let events: IEventData[] = [];
    if (!siteUrl) {
      return [];
    }
    try {
      // Get Regional Settings TimeZone Hours to UTC
      const siteTimeZoneHoursToUTC: number = await this.getSiteTimeZoneHoursToUtc(siteUrl);
      // Get Category Field Choices
      const categoryDropdownOption = await this.getChoiceFieldOptions(siteUrl, listId, 'Category');
      let categoryColor: { category: string, color: string }[] = [];
      for (const cat of categoryDropdownOption) {
        categoryColor.push({ category: cat.text, color: await this.colorGenerate() });
      }

      const web = new Web(siteUrl);
      const results = await web.lists.getById(listId).renderListDataAsStream(
        {
          DatesInUtc: true,
          ViewXml: `<View><ViewFields><FieldRef Name='Status'/><FieldRef Name='Requestor'/><FieldRef Name='Author'/><FieldRef Name='Manager'/><FieldRef Name='ManagerApproved'/><FieldRef Name='BackupApproved'/><FieldRef Name='Category'/><FieldRef Name='Description'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='ID'/><FieldRef Name='EndDate'/><FieldRef Name='EventDate'/><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='fAllDayEvent'/></ViewFields>
          <Query>
          <Where>
            <And>
             <And>
              <Geq>
                <FieldRef Name='EventDate' />
                <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventStartDate).format('YYYY-MM-DD')}</Value>
              </Geq>
              <Leq>
                <FieldRef Name='EventDate' />
                <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventEndDate).format('YYYY-MM-DD')}</Value>
              </Leq>
              </And>
              <Eq>
              <FieldRef Name='fRecurrence' />
                <Value Type='Recurrence'>0</Value>
              </Eq>
            </And>
          </Where>
          </Query>
          <RowLimit Paged=\"FALSE\">2000</RowLimit>
          </View>`
        }
      );
      if (results && results.Row.length > 0) {
        for (const event of results.Row) {
          const initialsArray: string[] = event.Requestor[0].title.split(' ');
          //const initials: string = "test";
          const initials: string = initialsArray[0].charAt(0) + initialsArray[initialsArray.length - 1].charAt(0);
          //const userPictureUrl = await this.getUserProfilePictureUrl(`i:0#.f|membership|${event.Requestor[0].email}`);
          const fAllDayEvent: boolean =  (event.fAllDayEvent == "Yes") ? true : false;         
          const CategoryColorValue: any[] = categoryColor.filter((value) => {
            return value.category == event.Category;
          });

          const backupId: number = ( event.ParticipantsPicker[0] != null) ? event.ParticipantsPicker[0].id : null;    
          //const backupObj : any = (backupId != null) ? (await web.siteUsers.getById(backupId).get()).Title : null;    
          const backupObj : any = (event.ParticipantsPicker[0] != null) ? event.ParticipantsPicker[0].title : null;
          const backupApproved: boolean =  (event.BackupApproved == "Yes") ? true : false;   
          const managerId: number = ( event.Manager[0] != null) ? event.Manager[0].id : null;    
          //const managerObj : any = (managerId != null) ? (await web.siteUsers.getById(managerId).get()).Title : null;      
          const managerObj : any = event.Manager[0].title;
          const managerApproved: boolean =  (event.ManagerApproved == "Yes") ? true : false;       
          const img = new Image();
          img.src = `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${event.Requestor[0].email}&UA=0&size=HR96x96`;
          const startDate =  new Date(moment(event.EventDate).toISOString());
          const endDate = new Date(moment(event.EndDate).toISOString());
          if ((allowPending || (!allowPending && managerApproved)) && event.Status !== "Cancelled" && event.Status !== "Rejected" )
          events.push({
            id: event.ID,
            title: event.Requestor[0].title + " " + event.Category,
            Description: event.Description,
            start: startDate,
            end: endDate,
            ownerEmail: event.Requestor[0].email,
            ownerPhoto:img.width !== 1? `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${event.Requestor[0].email}&UA=0&size=HR96x96` : "",
            ownerInitial: initials,
            //color: await this.colorGenerate(),
            color: await this.colorGenerateEvents(event),
            ownerName: event.Requestor[0].title,
            backup: backupId,      
            backupName: backupObj,
            backupApproved: backupApproved,      
            manager : managerId,       
            managerName: managerObj,
            managerApproved: managerApproved,   
            allDayEvent: fAllDayEvent,
            Category: event.Category
          });
        }
      }
      // Return Data
      return events;
    } catch (error) {
      console.log(error)
      return Promise.reject(error);
    }
  }

  /**
   *
   * @private
   * @param {string} siteUrl
   * @returns
   * @memberof spservices
   */
  public async getSiteRegionalSettingsTimeZone(siteUrl: string) {
    let regionalSettings: RegionalSettings;
    try {
      const web = new Web(siteUrl);
      regionalSettings = await web.regionalSettings.timeZone.get();

    } catch (error) {
      return Promise.reject(error);
    }
    return regionalSettings;
  }

  public async enCodeHtmlEntities(string: string) {

    const HtmlEntitiesMap = {
      "'": "&apos;",
      "<": "&lt;",
      ">": "&gt;",
      " ": "&nbsp;",
      "¡": "&iexcl;",
      "¢": "&cent;",
      "£": "&pound;",
      "¤": "&curren;",
      "¥": "&yen;",
      "¦": "&brvbar;",
      "§": "&sect;",
      "¨": "&uml;",
      "©": "&copy;",
      "ª": "&ordf;",
      "«": "&laquo;",
      "¬": "&not;",
      "®": "&reg;",
      "¯": "&macr;",
      "°": "&deg;",
      "±": "&plusmn;",
      "²": "&sup2;",
      "³": "&sup3;",
      "´": "&acute;",
      "µ": "&micro;",
      "¶": "&para;",
      "·": "&middot;",
      "¸": "&cedil;",
      "¹": "&sup1;",
      "º": "&ordm;",
      "»": "&raquo;",
      "¼": "&frac14;",
      "½": "&frac12;",
      "¾": "&frac34;",
      "¿": "&iquest;",
      "À": "&Agrave;",
      "Á": "&Aacute;",
      "Â": "&Acirc;",
      "Ã": "&Atilde;",
      "Ä": "&Auml;",
      "Å": "&Aring;",
      "Æ": "&AElig;",
      "Ç": "&Ccedil;",
      "È": "&Egrave;",
      "É": "&Eacute;",
      "Ê": "&Ecirc;",
      "Ë": "&Euml;",
      "Ì": "&Igrave;",
      "Í": "&Iacute;",
      "Î": "&Icirc;",
      "Ï": "&Iuml;",
      "Ð": "&ETH;",
      "Ñ": "&Ntilde;",
      "Ò": "&Ograve;",
      "Ó": "&Oacute;",
      "Ô": "&Ocirc;",
      "Õ": "&Otilde;",
      "Ö": "&Ouml;",
      "×": "&times;",
      "Ø": "&Oslash;",
      "Ù": "&Ugrave;",
      "Ú": "&Uacute;",
      "Û": "&Ucirc;",
      "Ü": "&Uuml;",
      "Ý": "&Yacute;",
      "Þ": "&THORN;",
      "ß": "&szlig;",
      "à": "&agrave;",
      "á": "&aacute;",
      "â": "&acirc;",
      "ã": "&atilde;",
      "ä": "&auml;",
      "å": "&aring;",
      "æ": "&aelig;",
      "ç": "&ccedil;",
      "è": "&egrave;",
      "é": "&eacute;",
      "ê": "&ecirc;",
      "ë": "&euml;",
      "ì": "&igrave;",
      "í": "&iacute;",
      "î": "&icirc;",
      "ï": "&iuml;",
      "ð": "&eth;",
      "ñ": "&ntilde;",
      "ò": "&ograve;",
      "ó": "&oacute;",
      "ô": "&ocirc;",
      "õ": "&otilde;",
      "ö": "&ouml;",
      "÷": "&divide;",
      "ø": "&oslash;",
      "ù": "&ugrave;",
      "ú": "&uacute;",
      "û": "&ucirc;",
      "ü": "&uuml;",
      "ý": "&yacute;",
      "þ": "&thorn;",
      "ÿ": "&yuml;",
      "Œ": "&OElig;",
      "œ": "&oelig;",
      "Š": "&Scaron;",
      "š": "&scaron;",
      "Ÿ": "&Yuml;",
      "ƒ": "&fnof;",
      "ˆ": "&circ;",
      "˜": "&tilde;",
      "Α": "&Alpha;",
      "Β": "&Beta;",
      "Γ": "&Gamma;",
      "Δ": "&Delta;",
      "Ε": "&Epsilon;",
      "Ζ": "&Zeta;",
      "Η": "&Eta;",
      "Θ": "&Theta;",
      "Ι": "&Iota;",
      "Κ": "&Kappa;",
      "Λ": "&Lambda;",
      "Μ": "&Mu;",
      "Ν": "&Nu;",
      "Ξ": "&Xi;",
      "Ο": "&Omicron;",
      "Π": "&Pi;",
      "Ρ": "&Rho;",
      "Σ": "&Sigma;",
      "Τ": "&Tau;",
      "Υ": "&Upsilon;",
      "Φ": "&Phi;",
      "Χ": "&Chi;",
      "Ψ": "&Psi;",
      "Ω": "&Omega;",
      "α": "&alpha;",
      "β": "&beta;",
      "γ": "&gamma;",
      "δ": "&delta;",
      "ε": "&epsilon;",
      "ζ": "&zeta;",
      "η": "&eta;",
      "θ": "&theta;",
      "ι": "&iota;",
      "κ": "&kappa;",
      "λ": "&lambda;",
      "μ": "&mu;",
      "ν": "&nu;",
      "ξ": "&xi;",
      "ο": "&omicron;",
      "π": "&pi;",
      "ρ": "&rho;",
      "ς": "&sigmaf;",
      "σ": "&sigma;",
      "τ": "&tau;",
      "υ": "&upsilon;",
      "φ": "&phi;",
      "χ": "&chi;",
      "ψ": "&psi;",
      "ω": "&omega;",
      "ϑ": "&thetasym;",
      "ϒ": "&Upsih;",
      "ϖ": "&piv;",
      "–": "&ndash;",
      "—": "&mdash;",
      "‘": "&lsquo;",
      "’": "&rsquo;",
      "‚": "&sbquo;",
      "“": "&ldquo;",
      "”": "&rdquo;",
      "„": "&bdquo;",
      "†": "&dagger;",
      "‡": "&Dagger;",
      "•": "&bull;",
      "…": "&hellip;",
      "‰": "&permil;",
      "′": "&prime;",
      "″": "&Prime;",
      "‹": "&lsaquo;",
      "›": "&rsaquo;",
      "‾": "&oline;",
      "⁄": "&frasl;",
      "€": "&euro;",
      "ℑ": "&image;",
      "℘": "&weierp;",
      "ℜ": "&real;",
      "™": "&trade;",
      "ℵ": "&alefsym;",
      "←": "&larr;",
      "↑": "&uarr;",
      "→": "&rarr;",
      "↓": "&darr;",
      "↔": "&harr;",
      "↵": "&crarr;",
      "⇐": "&lArr;",
      "⇑": "&UArr;",
      "⇒": "&rArr;",
      "⇓": "&dArr;",
      "⇔": "&hArr;",
      "∀": "&forall;",
      "∂": "&part;",
      "∃": "&exist;",
      "∅": "&empty;",
      "∇": "&nabla;",
      "∈": "&isin;",
      "∉": "&notin;",
      "∋": "&ni;",
      "∏": "&prod;",
      "∑": "&sum;",
      "−": "&minus;",
      "∗": "&lowast;",
      "√": "&radic;",
      "∝": "&prop;",
      "∞": "&infin;",
      "∠": "&ang;",
      "∧": "&and;",
      "∨": "&or;",
      "∩": "&cap;",
      "∪": "&cup;",
      "∫": "&int;",
      "∴": "&there4;",
      "∼": "&sim;",
      "≅": "&cong;",
      "≈": "&asymp;",
      "≠": "&ne;",
      "≡": "&equiv;",
      "≤": "&le;",
      "≥": "&ge;",
      "⊂": "&sub;",
      "⊃": "&sup;",
      "⊄": "&nsub;",
      "⊆": "&sube;",
      "⊇": "&supe;",
      "⊕": "&oplus;",
      "⊗": "&otimes;",
      "⊥": "&perp;",
      "⋅": "&sdot;",
      "⌈": "&lceil;",
      "⌉": "&rceil;",
      "⌊": "&lfloor;",
      "⌋": "&rfloor;",
      "⟨": "&lang;",
      "⟩": "&rang;",
      "◊": "&loz;",
      "♠": "&spades;",
      "♣": "&clubs;",
      "♥": "&hearts;",
      "♦": "&diams;"
    };

      var entityMap = HtmlEntitiesMap;
      string = string.replace(/&/g, '&amp;');
      string = string.replace(/"/g, '&quot;');
      for (var key in entityMap) {
        var entity = entityMap[key];
        var regex = new RegExp(key, 'g');
        string = string.replace(regex, entity);
      }
      return string;
  }

  public async deCodeHtmlEntities(string: string) {

    const HtmlEntitiesMap = {
      "'": "&#39;",
      "<": "&lt;",
      ">": "&gt;",
      " ": "&nbsp;",
      "¡": "&iexcl;",
      "¢": "&cent;",
      "£": "&pound;",
      "¤": "&curren;",
      "¥": "&yen;",
      "¦": "&brvbar;",
      "§": "&sect;",
      "¨": "&uml;",
      "©": "&copy;",
      "ª": "&ordf;",
      "«": "&laquo;",
      "¬": "&not;",
      "®": "&reg;",
      "¯": "&macr;",
      "°": "&deg;",
      "±": "&plusmn;",
      "²": "&sup2;",
      "³": "&sup3;",
      "´": "&acute;",
      "µ": "&micro;",
      "¶": "&para;",
      "·": "&middot;",
      "¸": "&cedil;",
      "¹": "&sup1;",
      "º": "&ordm;",
      "»": "&raquo;",
      "¼": "&frac14;",
      "½": "&frac12;",
      "¾": "&frac34;",
      "¿": "&iquest;",
      "À": "&Agrave;",
      "Á": "&Aacute;",
      "Â": "&Acirc;",
      "Ã": "&Atilde;",
      "Ä": "&Auml;",
      "Å": "&Aring;",
      "Æ": "&AElig;",
      "Ç": "&Ccedil;",
      "È": "&Egrave;",
      "É": "&Eacute;",
      "Ê": "&Ecirc;",
      "Ë": "&Euml;",
      "Ì": "&Igrave;",
      "Í": "&Iacute;",
      "Î": "&Icirc;",
      "Ï": "&Iuml;",
      "Ð": "&ETH;",
      "Ñ": "&Ntilde;",
      "Ò": "&Ograve;",
      "Ó": "&Oacute;",
      "Ô": "&Ocirc;",
      "Õ": "&Otilde;",
      "Ö": "&Ouml;",
      "×": "&times;",
      "Ø": "&Oslash;",
      "Ù": "&Ugrave;",
      "Ú": "&Uacute;",
      "Û": "&Ucirc;",
      "Ü": "&Uuml;",
      "Ý": "&Yacute;",
      "Þ": "&THORN;",
      "ß": "&szlig;",
      "à": "&agrave;",
      "á": "&aacute;",
      "â": "&acirc;",
      "ã": "&atilde;",
      "ä": "&auml;",
      "å": "&aring;",
      "æ": "&aelig;",
      "ç": "&ccedil;",
      "è": "&egrave;",
      "é": "&eacute;",
      "ê": "&ecirc;",
      "ë": "&euml;",
      "ì": "&igrave;",
      "í": "&iacute;",
      "î": "&icirc;",
      "ï": "&iuml;",
      "ð": "&eth;",
      "ñ": "&ntilde;",
      "ò": "&ograve;",
      "ó": "&oacute;",
      "ô": "&ocirc;",
      "õ": "&otilde;",
      "ö": "&ouml;",
      "÷": "&divide;",
      "ø": "&oslash;",
      "ù": "&ugrave;",
      "ú": "&uacute;",
      "û": "&ucirc;",
      "ü": "&uuml;",
      "ý": "&yacute;",
      "þ": "&thorn;",
      "ÿ": "&yuml;",
      "Œ": "&OElig;",
      "œ": "&oelig;",
      "Š": "&Scaron;",
      "š": "&scaron;",
      "Ÿ": "&Yuml;",
      "ƒ": "&fnof;",
      "ˆ": "&circ;",
      "˜": "&tilde;",
      "Α": "&Alpha;",
      "Β": "&Beta;",
      "Γ": "&Gamma;",
      "Δ": "&Delta;",
      "Ε": "&Epsilon;",
      "Ζ": "&Zeta;",
      "Η": "&Eta;",
      "Θ": "&Theta;",
      "Ι": "&Iota;",
      "Κ": "&Kappa;",
      "Λ": "&Lambda;",
      "Μ": "&Mu;",
      "Ν": "&Nu;",
      "Ξ": "&Xi;",
      "Ο": "&Omicron;",
      "Π": "&Pi;",
      "Ρ": "&Rho;",
      "Σ": "&Sigma;",
      "Τ": "&Tau;",
      "Υ": "&Upsilon;",
      "Φ": "&Phi;",
      "Χ": "&Chi;",
      "Ψ": "&Psi;",
      "Ω": "&Omega;",
      "α": "&alpha;",
      "β": "&beta;",
      "γ": "&gamma;",
      "δ": "&delta;",
      "ε": "&epsilon;",
      "ζ": "&zeta;",
      "η": "&eta;",
      "θ": "&theta;",
      "ι": "&iota;",
      "κ": "&kappa;",
      "λ": "&lambda;",
      "μ": "&mu;",
      "ν": "&nu;",
      "ξ": "&xi;",
      "ο": "&omicron;",
      "π": "&pi;",
      "ρ": "&rho;",
      "ς": "&sigmaf;",
      "σ": "&sigma;",
      "τ": "&tau;",
      "υ": "&upsilon;",
      "φ": "&phi;",
      "χ": "&chi;",
      "ψ": "&psi;",
      "ω": "&omega;",
      "ϑ": "&thetasym;",
      "ϒ": "&Upsih;",
      "ϖ": "&piv;",
      "–": "&ndash;",
      "—": "&mdash;",
      "‘": "&lsquo;",
      "’": "&rsquo;",
      "‚": "&sbquo;",
      "“": "&ldquo;",
      "”": "&rdquo;",
      "„": "&bdquo;",
      "†": "&dagger;",
      "‡": "&Dagger;",
      "•": "&bull;",
      "…": "&hellip;",
      "‰": "&permil;",
      "′": "&prime;",
      "″": "&Prime;",
      "‹": "&lsaquo;",
      "›": "&rsaquo;",
      "‾": "&oline;",
      "⁄": "&frasl;",
      "€": "&euro;",
      "ℑ": "&image;",
      "℘": "&weierp;",
      "ℜ": "&real;",
      "™": "&trade;",
      "ℵ": "&alefsym;",
      "←": "&larr;",
      "↑": "&uarr;",
      "→": "&rarr;",
      "↓": "&darr;",
      "↔": "&harr;",
      "↵": "&crarr;",
      "⇐": "&lArr;",
      "⇑": "&UArr;",
      "⇒": "&rArr;",
      "⇓": "&dArr;",
      "⇔": "&hArr;",
      "∀": "&forall;",
      "∂": "&part;",
      "∃": "&exist;",
      "∅": "&empty;",
      "∇": "&nabla;",
      "∈": "&isin;",
      "∉": "&notin;",
      "∋": "&ni;",
      "∏": "&prod;",
      "∑": "&sum;",
      "−": "&minus;",
      "∗": "&lowast;",
      "√": "&radic;",
      "∝": "&prop;",
      "∞": "&infin;",
      "∠": "&ang;",
      "∧": "&and;",
      "∨": "&or;",
      "∩": "&cap;",
      "∪": "&cup;",
      "∫": "&int;",
      "∴": "&there4;",
      "∼": "&sim;",
      "≅": "&cong;",
      "≈": "&asymp;",
      "≠": "&ne;",
      "≡": "&equiv;",
      "≤": "&le;",
      "≥": "&ge;",
      "⊂": "&sub;",
      "⊃": "&sup;",
      "⊄": "&nsub;",
      "⊆": "&sube;",
      "⊇": "&supe;",
      "⊕": "&oplus;",
      "⊗": "&otimes;",
      "⊥": "&perp;",
      "⋅": "&sdot;",
      "⌈": "&lceil;",
      "⌉": "&rceil;",
      "⌊": "&lfloor;",
      "⌋": "&rfloor;",
      "⟨": "&lang;",
      "⟩": "&rang;",
      "◊": "&loz;",
      "♠": "&spades;",
      "♣": "&clubs;",
      "♥": "&hearts;",
      "♦": "&diams;"
    };

    var entityMap = HtmlEntitiesMap;
    for (var key in entityMap) {
      var entity = entityMap[key];
      var regex = new RegExp(entity, 'g');
      string = string.replace(regex, key);
    }
    string = string.replace(/&quot;/g, '"');
    string = string.replace(/&amp;/g, '&');
    return string;
  }



}
