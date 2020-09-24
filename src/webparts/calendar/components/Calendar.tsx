import * as React from 'react';
import styles from './Calendar.module.scss';
import { ICalendarProps } from './ICalendarProps';
import { ICalendarState } from './ICalendarState';
import { escape } from '@microsoft/sp-lodash-subset';
import BigCalendar from 'react-big-calendar';
import * as moment from 'moment';
import * as moment_timezone from 'moment-timezone';
import * as strings from 'CalendarWebPartStrings';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import {

  IPersonaSharedProps,
  Persona,
  PersonaSize,
  PersonaPresence,
  HoverCard, HoverCardType,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardDetails,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DefaultButton

} from 'office-ui-fabric-react';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { DisplayMode } from '@microsoft/sp-core-library';
import spservices from '../../../services/spservices';
import { Event } from '../../../controls/Event/event';
import { RequestList } from '../../../controls/RequestsList/requestList';
import { IPanelModelEnum } from '../../../controls/Event/IPanelModeEnum';
import { IEventData } from './../../../services/IEventData';
import { IUserPermissions } from './../../../services/IUserPermissions';
import { Item } from '@pnp/sp';
import { event } from 'jquery';
/*
moment.tz('UTC')
const localizer = BigCalendar.momentLocalizer(moment);
*/
moment_timezone.tz.setDefault("America/Los Angeles");
const localizer = BigCalendar.momentLocalizer(moment_timezone);
/**
 * @export
 * @class Calendar
 * @extends {React.Component<ICalendarProps, ICalendarState>}
 */
export default class Calendar extends React.Component<ICalendarProps, ICalendarState> {
  private spService: spservices = null;
  private userListPermissions: IUserPermissions = undefined;
  public constructor(props) {
    super(props);

    this.state = {
      showDialog: false,
      eventData: [],
      selectedEvent: undefined,
      isloading: true,
      hasError: false,
      errorMessage: '',
      showRequests: false
    };

    this.onDismissPanel = this.onDismissPanel.bind(this);
    this.onSelectEvent = this.onSelectEvent.bind(this);
    this.onSelectSlot = this.onSelectSlot.bind(this);
    this.onShowRequests = this.onShowRequests.bind(this);
    this.spService = new spservices(this.props.context);
    moment.locale(this.props.context.pageContext.cultureInfo.currentUICultureName);

  }

  private onDocumentCardClick(ev: React.SyntheticEvent<HTMLElement, Event>) {
    ev.preventDefault();
    ev.stopPropagation();
  }
  /**
   * @private
   * @param {*} event
   * @memberof Calendar
   */
  private onSelectEvent(event: any) {
    this.setState({ showDialog: true, selectedEvent: event, panelMode: IPanelModelEnum.edit });
  }
  /**
   * @private
   * @param {*} event
   * @memberof Calendar
   */
  private onShowRequests() {
    this.setState({ showRequests: true});
  }

  /**
   *
   * @private
   * @param {boolean} refresh
   * @memberof Calendar
   */
  private async onDismissPanel(refresh: boolean) {

    this.setState({ showDialog: false, showRequests: false });
    if (refresh === true) {
      this.setState({ isloading: true });
      await this.loadEvents();
      this.setState({ isloading: false });
    }
  }
  
  /**
   * @private
   * @memberof Calendar
   */
  private async loadEvents() {
    try {
      // Teste Properties
      if (!this.props.list || !this.props.siteUrl || !this.props.eventStartDate.value || !this.props.eventEndDate.value) return;

      this.userListPermissions = await this.spService.getUserPermissions(this.props.siteUrl, this.props.list);
      const eventsData: IEventData[] = await this.spService.getEvents(escape(this.props.siteUrl), escape(this.props.list), this.props.eventStartDate.value, this.props.eventEndDate.value, this.props.allowPending);
      
      for (const event of eventsData){ 
        let startOffset = (event.start.getTimezoneOffset());
        let endOffset = (event.end.getTimezoneOffset());
        //correct issue with items listed as all day events
        if (event.allDayEvent){
        event.end.setTime(event.end.getTime() + (7*60*60*1000)); 
        event.start.setTime(event.start.getTime() + (7*60*60*1000)); 
        }
        //offset for DST
        event.end.setHours(event.end.getHours() + ((endOffset - 420)/60));
        event.start.setHours(event.start.getHours() + ((startOffset - 420)/60));
      }
      this.setState({ eventData: eventsData, hasError: false, errorMessage: "" });

    } catch (error) {
      this.setState({ hasError: true, errorMessage: error.message, isloading: false });
    }
  }
  /**
   * @memberof Calendar
   */
  public async componentDidMount() {
    this.setState({ isloading: true });
    await this.loadEvents();
    this.setState({ isloading: false });
  }

  /**
   *
   *
   * @param {*} error
   * @param {*} errorInfo
   * @memberof Calendar
   */
  public componentDidCatch(error: any, errorInfo: any) {
    this.setState({ hasError: true, errorMessage: errorInfo.componentStack });
  }
  /**
   *
   *
   * @param {ICalendarProps} prevProps
   * @param {ICalendarState} prevState
   * @memberof Calendar
   */
  public async componentDidUpdate(prevProps: ICalendarProps, prevState: ICalendarState) {

    if (!this.props.list || !this.props.siteUrl || !this.props.eventStartDate.value || !this.props.eventEndDate.value) return;
    // Get  Properties change
    if (prevProps.list !== this.props.list || this.props.eventStartDate.value !== prevProps.eventStartDate.value || this.props.eventEndDate.value !== prevProps.eventEndDate.value) {
      this.setState({ isloading: true });
      await this.loadEvents();
      this.setState({ isloading: false });
    }
  }

  /**
   * @private
   * @param {*} { event }
   * @param {boolean} allowPending
   * @returns
   * @memberof Calendar
   */
  private renderEvent({ event }) {
    const previewEventIcon: IDocumentCardPreviewProps = {
      previewImages: [
        {
          // previewImageSrc: event.ownerPhoto,
          previewIconProps: { iconName: 'Calendar', styles: { root: { color: event.color } }, className: styles.previewEventIcon },
          height: 43,
        }
      ]
    };
    const EventInfo: IPersonaSharedProps = {
      imageInitials: event.ownerInitial,
      imageUrl: event.ownerPhoto,
      text: event.title
    };
    const allowBackup : boolean = (event.backupName != '' &&  event.backupName != null) ? true : false;
    const approvalColor : number = (event.backupApproved) ? 4 : 13;
    const approvalStatus : string = (event.managerApproved) ? "Approved by " + event.managerName : (event.status) === "Rejected" ? "Rejected" : "Requested";
    const approvalColorHex : string = (event.managerApproved) ? "#96e884" : "#e88484";
    /**
     * @returns {JSX.Element}
     */
    const onRenderPlainCard = (): JSX.Element => {
        return (
          <div className={styles.plainCard}>
            <DocumentCard className={styles.Documentcard}   >
              <div>
                <DocumentCardPreview {...previewEventIcon} />
              </div>
              <DocumentCardDetails>
                <div className={styles.DocumentCardDetails}>
                  <DocumentCardTitle title={event.title} shouldTruncate={true} className={styles.DocumentCardTitle} styles={{ root: { color: event.color} }} />
                </div>
                {<div className={styles.ApprovalStatus} style={{ backgroundColor: approvalColorHex} }>{approvalStatus}</div>}
                {
                  moment(event.start).format('YYYY/MM/DD') !== moment(event.end).format('YYYY/MM/DD') ?
                    <span className={styles.DocumentCardTitleTime}>{moment(event.start).format('dddd')} - {moment(event.end).format('dddd')} </span>
                    :
                    <span className={styles.DocumentCardTitleTime}>{moment(event.start).format('dddd')} </span>
                }
                <span className={styles.DocumentCardTitleTime}>{moment(event.start).format('HH:mm')}H - {moment(event.end).format('HH:mm')}H</span>
                {allowBackup && 
                <div style={{ marginTop: 20 }}>                
                  <DocumentCardActivity
                    activity={strings.BackupLabel}
                    people={[{ name: event.backupName, profileImageSrc: null, initialsColor: approvalColor }]}
                  />
                </div>}
              </DocumentCardDetails>
            </DocumentCard>
          </div>
        );      
    };

    return (
      <div style={{ height: 22 }}>
        <HoverCard
          cardDismissDelay={1000}
          type={HoverCardType.plain}
          plainCardProps={{ onRenderPlainCard: onRenderPlainCard }}
          onCardHide={(): void => {
          }}
        >
          <Persona
            {...EventInfo}
            size={PersonaSize.size24}
            presence={PersonaPresence.none}
            coinSize={22}
            initialsColor={event.color}
          />
        </HoverCard>
      </div>
    );
  }
  /**
   *
   *
   * @private
   * @memberof Calendar
   */
  private onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

  /**
   * @param {*} { start, end }
   * @memberof Calendar
   */
  public async onSelectSlot({ start, end }) {
    if (!this.userListPermissions.hasPermissionAdd) return;
    this.setState({ showDialog: true, startDateSlot: start, endDateSlot: end, selectedEvent: undefined, panelMode: IPanelModelEnum.add });
  }

  /**
   *
   * @param {*} event
   * @param {*} start
   * @param {*} end
   * @param {*} isSelected
   * @returns {*}
   * @memberof Calendar
   */
  public eventStyleGetter(event, start, end, isSelected): any {
    let style: any = {
      backgroundColor: 'white',
      borderRadius: '0px',
      opacity: 1,
      color: 'black',
      borderWidth: '1.1px',
      borderStyle: 'solid',
      borderColor: event.color,
      borderLeftWidth: '5px',
      display: 'block'
    };

    return {
      style: style
    };
  }

  /**
   *
   * @returns {React.ReactElement<ICalendarProps>}
   * @memberof Calendar
   */
  public render(): React.ReactElement<ICalendarProps> {

    return (
      <div className={styles.calendar}>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty} />
        {
          (!this.props.list || !this.props.eventStartDate.value || !this.props.eventEndDate.value) ?
            <Placeholder iconName='Edit'
              iconText={strings.WebpartConfigIconText}
              description={strings.WebpartConfigDescription}
              buttonLabel={strings.WebPartConfigButtonLabel}
              hideButton={this.props.displayMode === DisplayMode.Read}
              onConfigure={this.onConfigure.bind(this)} />
            :
            // test if has errors
            this.state.hasError ?
              <MessageBar messageBarType={MessageBarType.error}>
                {this.state.errorMessage}
              </MessageBar>
              :
              // show Calendar
              // Test if is loading Events
              <div>
                {this.state.isloading ? <Spinner size={SpinnerSize.large} label={strings.LoadingEventsLabel} /> :
                  <>
                  <div className={styles.container}>
                    <BigCalendar
                      localizer={localizer}
                      selectable
                      length={0}
                      popup
                      views={['month', 'week', 'day']}
                      events={this.state.eventData}
                      startAccessor="start"
                      endAccessor="end"
                      eventPropGetter={this.eventStyleGetter}
                      onSelectSlot={this.onSelectSlot}
                      components={{
                        event: this.renderEvent
                      }}
                      onSelectEvent={this.onSelectEvent}
                      defaultDate={moment().startOf('day').toDate()}
                      messages={
                        {
                          'today': strings.todayLabel,
                          'previous': strings.previousLabel,
                          'next': strings.nextLabel,
                          'month': strings.monthLabel,
                          'week': strings.weekLabel,
                          'day': strings.dayLable,
                          'agenda': "Daily Agenda",
                          'showMore': total => `+${total} ${strings.showMore}`
                        }
                      }
                    />
                  </div>
                  <div className={styles.calendarFooterControls}>
                    <DefaultButton type="button" onClick={() => this.onShowRequests()}>View My Requests</DefaultButton>
                  </div>
                  </>
                }
              </div>
        }
        {
          this.state.showDialog &&
          <Event
            event={this.state.selectedEvent}
            panelMode={this.state.panelMode}
            onDissmissPanel={this.onDismissPanel}
            showPanel={this.state.showDialog}
            startDate={this.state.startDateSlot}
            endDate={this.state.endDateSlot}
            context={this.props.context}
            siteUrl={this.props.siteUrl}
            listId={this.props.list}
            allowEdit={this.props.allowEdit}
            allowBackup={this.props.allowBackup}
          />
        }
        {
          this.state.showRequests &&
          <RequestList
            onDissmissPanel={this.onDismissPanel}
            showPanel={this.state.showRequests}
            context={this.props.context}
            siteUrl={this.props.siteUrl}
            listId={this.props.list}
            list={this.props.list}
            eventStartDate={this.props.eventStartDate}
            eventEndDate={this.props.eventEndDate}
          />
        }
      </div>
    );
  }
}
