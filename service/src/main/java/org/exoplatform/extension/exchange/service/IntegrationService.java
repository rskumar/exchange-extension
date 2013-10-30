package org.exoplatform.extension.exchange.service;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.BasePropertySet;
import microsoft.exchange.webservices.data.CalendarFolder;
import microsoft.exchange.webservices.data.EventType;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.FindFoldersResults;
import microsoft.exchange.webservices.data.FindItemsResults;
import microsoft.exchange.webservices.data.Folder;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.FolderView;
import microsoft.exchange.webservices.data.Item;
import microsoft.exchange.webservices.data.ItemEvent;
import microsoft.exchange.webservices.data.ItemId;
import microsoft.exchange.webservices.data.ItemSchema;
import microsoft.exchange.webservices.data.ItemView;
import microsoft.exchange.webservices.data.PropertySet;
import microsoft.exchange.webservices.data.SearchFilter;
import microsoft.exchange.webservices.data.ServiceResponseException;
import microsoft.exchange.webservices.data.WellKnownFolderName;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.calendar.service.CalendarSetting;
import org.exoplatform.container.PortalContainer;
import org.exoplatform.container.component.ComponentRequestLifecycle;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;
import org.exoplatform.services.organization.UserProfile;

/**
 * 
 * @author Boubaker KHANFIR
 * 
 */
public class IntegrationService {

  private final static Log LOG = ExoLogger.getLogger(IntegrationService.class);

  private final static SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMM yyyy HH:mm:ss");

  private static final String USER_EXCHANGE_HANDLED_ATTRIBUTE = "exchange.check.date";
  private static final Map<String, IntegrationService> instances = new HashMap<String, IntegrationService>();

  private final String username;
  private final ExchangeService service;
  private final ExoStorageService exoStorageService;
  private final ExchangeStorageService exchangeStorageService;
  private final CorrespondenceService correspondenceService;
  private final OrganizationService organizationService;
  private final CalendarService calendarService;

  public IntegrationService(OrganizationService organizationService, CalendarService calendarService, ExoStorageService exoStorageService, ExchangeStorageService exchangeStorageService,
      CorrespondenceService correspondenceService, ExchangeService service, String username) {
    this.organizationService = organizationService;
    this.calendarService = calendarService;
    this.exoStorageService = exoStorageService;
    this.exchangeStorageService = exchangeStorageService;
    this.correspondenceService = correspondenceService;
    this.service = service;
    this.username = username;

    // Set corresponding service to each username.
    instances.put(username, this);
  }

  /**
   * Gets user exchange instance service.
   * 
   * @param username
   * @return
   */
  public static IntegrationService getInstance(String username) {
    return instances.get(username);
  }

  /**
   * 
   * Synchronize Exchange Calendar identified by 'folderId' with eXo Calendar.
   * 
   * @param folderId
   * @param lastSyncDate
   * @param diffTimeZone
   * @throws Exception
   * @return List of event IDs
   */
  public List<String> synchronizeFullCalendar(FolderId folderId) throws Exception {
    List<String> events = new ArrayList<String>();
    CalendarFolder folder = null;
    try {
      folder = CalendarFolder.bind(service, folderId);
    } catch (ServiceResponseException e) {
      LOG.warn("Can't get Folder identified by id: " + folderId.getUniqueId());
      return events;
    }
    exoStorageService.getOrCreateUserCalendar(username, folder);

    Iterable<Item> items = searchAllAppointments(folderId);
    for (Item item : items) {
      if (item instanceof Appointment) {
        CalendarEvent event = exoStorageService.createOrUpdateEvent((Appointment) item, folder, username, getUserExoCalenarTimeZoneSetting());
        events.add(event.getId());
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
    return events;
  }

  /**
   * 
   * Synchronize Exchange Calendar identified by 'folderId' with eXo Calendar.
   * The check is done for events modified since 'lastSyncDate'.
   * 
   * @param folderId
   * @param lastSyncDate
   * @param updatedExoEventIDs
   * @param diffTimeZone
   * @throws Exception
   */
  public void synchronizeModificationsOfCalendar(FolderId folderId, Date lastSyncDate, List<String> updatedExoEventIDs, int diffTimeZone) throws Exception {
    Iterable<Item> items = searchAllAppointmentsModifiedSince(folderId, lastSyncDate, diffTimeZone);
    // Search for modified Appointments in Exchange, since last check date.
    List<String> updatedEventIDs = updatedExoEventIDs;
    if (updatedEventIDs == null) {
      updatedEventIDs = new ArrayList<String>();
    }
    CalendarEvent event = null;
    Folder folder = null;
    for (Item item : items) {
      if (item instanceof Appointment) {
        folder = Folder.bind(service, item.getParentFolderId());
        // Test if there is a modification conflict
        event = exoStorageService.getEventByAppointmentId(username, item.getId().getUniqueId());
        if (event != null) {
          Date eventModifDate = convertDateToUTC(event.getLastUpdatedTime());
          Date itemModifDate = item.getLastModifiedTime();
          if (itemModifDate.after(eventModifDate)) {
            event = exoStorageService.updateEvent((Appointment) item, folder, username, getUserExoCalenarTimeZoneSetting());
            if (event != null) {
              updatedEventIDs.add(event.getId());
            }
          }
        } else {
          event = exoStorageService.createEvent((Appointment) item, folder, username, getUserExoCalenarTimeZoneSetting());
          updatedEventIDs.add(event.getId());
        }
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
    // Search for unsynchronized Appointments with Exchange
    Iterable<CalendarEvent> unsynchronizedEvents = searchUnsynchronizedAppointments(username, folderId.getUniqueId());
    for (CalendarEvent calendarEvent : unsynchronizedEvents) {
      if (calendarEvent.getLastUpdatedTime() != null) {
        if (calendarEvent.getLastUpdatedTime().after(lastSyncDate)) {
          LOG.info("Synchronize exo CalendarEvent with Exchange: " + calendarEvent.getSummary());
          exchangeStorageService.updateOrCreateExchangeAppointment(username, service, calendarEvent, getUserExoCalenarTimeZoneSetting());
          updatedEventIDs.add(calendarEvent.getId());
        } else {
          LOG.info("Delete user calendar event: " + calendarEvent.getSummary());
          exoStorageService.deleteEvent(calendarEvent, folderId.getUniqueId(), username);
        }
      }
    }
    // Serach modified eXo Calendar events since this date
    List<CalendarEvent> modifiedCalendarEvents = searchCalendarEventsModifiedSince(getUserCalendarByExchangeFolderId(folderId), lastSyncDate);
    for (CalendarEvent calendarEvent : modifiedCalendarEvents) {
      // If modified with synchronization, ignore
      if (updatedEventIDs.contains(calendarEvent.getId())) {
        continue;
      }
      LOG.info("Synchronize exo CalendarEvent with Exchange: " + calendarEvent.getSummary());
      exchangeStorageService.updateOrCreateExchangeAppointment(username, service, calendarEvent, getUserExoCalenarTimeZoneSetting());
    }
  }

  /**
   * 
   * Checks if eXo Calendar is present but not in Exchange. If so, this calendar
   * was deleted from Exchange and should be deleted from eXo.
   * 
   * @param calendarFolderIds
   * @throws Exception
   */
  public void synchronizeDeletedFolder(List<FolderId> calendarFolderIds) throws Exception {
    List<String> folderIds = new ArrayList<String>();
    for (FolderId folderId : calendarFolderIds) {
      folderIds.add(folderId.getUniqueId());
    }
    exoStorageService.deleteUnregisteredExchangeCalendars(username, folderIds);
  }

  /**
   * 
   * Gets list of personnal Exchange Calendars.
   * 
   * @return list of FolderId
   * @throws Exception
   */
  public List<FolderId> getExchangeCalendars() throws Exception {
    List<FolderId> calendarFolderIds = new ArrayList<FolderId>();
    CalendarFolder calendarRootFolder = CalendarFolder.bind(service, WellKnownFolderName.Calendar);

    calendarFolderIds.add(calendarRootFolder.getId());
    List<Folder> calendarfolders = searchSubFolders(calendarRootFolder.getId());

    if (calendarfolders != null && !calendarfolders.isEmpty()) {
      for (Folder tmpFolder : calendarfolders) {
        calendarFolderIds.add(tmpFolder.getId());
      }
    }
    return calendarFolderIds;
  }

  /**
   * Checks if eXo associated Calendar is present.
   * 
   * @param usename
   * @return true if present.
   * @throws Exception
   */
  public boolean isCalendarPresent(FolderId folderId) throws Exception {
    return exoStorageService.getUserCalendar(username, folderId.getUniqueId()) != null;
  }

  /**
   * 
   * Handle Exchange Calendar Deletion by deleting associated eXo Calendar.
   * 
   * @param folderId
   *          Exchange Calendar folderId
   * @return
   * @throws Exception
   */
  public boolean deleteCalendar(FolderId folderId) throws Exception {
    try {
      Folder folder = Folder.bind(service, folderId);
      if (folder != null) {
        LOG.info("Folder was found, but event seems saying that it was deleted.");
        return false;
      }
    } catch (Exception e) {
      // Folder doesn't exist
    }
    return exoStorageService.deleteCalendar(username, folderId.getUniqueId());
  }

  /**
   * 
   * Creates or updates or deletes eXo Calendar Event associated to Item, switch
   * state in Exchange.
   * 
   * @param itemEvent
   * @return
   * @throws Exception
   */
  public CalendarEvent createOrUpdateOrDelete(ItemEvent itemEvent) throws Exception {
    Appointment appointment = null;
    try {
      appointment = Appointment.bind(service, itemEvent.getItemId());
    } catch (ServiceResponseException e) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Item was not bound, it was deleted:" + itemEvent.getItemId());
      }
    }
    CalendarEvent event = null;
    if (appointment == null) {
      LOG.info("Delete user calendar event: " + itemEvent.getItemId().getUniqueId());
      Folder folder = Folder.bind(service, itemEvent.getParentFolderId());
      exoStorageService.deleteEventByAppointmentID(itemEvent.getItemId().getUniqueId(), folder.getId().getUniqueId(), username);
    } else {
      Folder folder = Folder.bind(service, appointment.getParentFolderId());
      if (itemEvent.getEventType() == EventType.Modified) {
        event = exoStorageService.updateEvent((Appointment) appointment, folder, username, getUserExoCalenarTimeZoneSetting());
      } else if (itemEvent.getEventType() == EventType.Created) {
        event = exoStorageService.createEvent((Appointment) appointment, folder, username, getUserExoCalenarTimeZoneSetting());
      }
    }
    return event;
  }

  /**
   * 
   * Sets last check operation date.
   * 
   * @param username
   * @param time
   * @throws Exception
   */
  public void setUserLastCheckDate(long time) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
    userProfile.setAttribute(USER_EXCHANGE_HANDLED_ATTRIBUTE, "" + time);
    organizationService.getUserProfileHandler().saveUserProfile(userProfile, false);
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
    }
  }

  /**
   * 
   * Gets last check operation date
   * 
   * @return
   * @throws Exception
   */
  public Date getUserLastCheckDate() throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
    long time = userProfile.getAttribute(USER_EXCHANGE_HANDLED_ATTRIBUTE) == null ? 0 : Long.valueOf(userProfile.getAttribute(USER_EXCHANGE_HANDLED_ATTRIBUTE));
    Date lastSyncDate = null;
    if (time > 0) {
      lastSyncDate = new Date(time);
    }
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
    }
    return lastSyncDate;
  }

  /**
   * 
   * Get corresponding User Calenar from Exchange Folder Id
   * 
   * @param folderId
   * @return
   * @throws Exception
   */
  public Calendar getUserCalendarByExchangeFolderId(FolderId folderId) throws Exception {
    return exoStorageService.getUserCalendar(username, folderId.getUniqueId());
  }

  @Override
  public void finalize() throws Throwable {
    LOG.info("Stop Exchange Integration Service for user: " + username);
    instances.remove(username);
  }

  public void updateOrCreateExchangeCalendarEvent(CalendarEvent event) throws Exception {
    exchangeStorageService.updateOrCreateExchangeAppointment(username, service, event, getUserExoCalenarTimeZoneSetting());
  }

  public void deleteExchangeCalendarEvent(CalendarEvent event) throws Exception {
    exchangeStorageService.deleteExchangeAppointment(username, service, event);
  }

  /**
   * This method returns User exo calendar TimeZone settings. This have to be
   * called each synchronization, the timeZone may be changed from call to
   * another.
   * 
   * @return User exo calendar TimeZone settings
   */
  private TimeZone getUserExoCalenarTimeZoneSetting() {
    try {
      CalendarSetting calendarSetting = calendarService.getCalendarSetting(username);
      return TimeZone.getTimeZone(calendarSetting.getTimeZone());
    } catch (Exception e) {
      LOG.error("Error while getting user '" + username + "'Calendar TimeZone setting, use default, this may cause some inconsistance.");
      return TimeZone.getDefault();
    }
  }

  private Iterable<CalendarEvent> searchUnsynchronizedAppointments(String username, String folderId) throws Exception {
    List<CalendarEvent> calendarEvents = exoStorageService.getUserCalendarEvents(username, folderId);
    Iterator<CalendarEvent> calendarEventsIterator = calendarEvents.iterator();
    while (calendarEventsIterator.hasNext()) {
      CalendarEvent calendarEvent = calendarEventsIterator.next();
      String itemId = correspondenceService.getCorrespondingId(username, calendarEvent.getId());
      if (itemId == null) {
        // New created appointment
        continue;
      }
      try {
        Appointment.bind(service, ItemId.getItemIdFromString(itemId));
        calendarEventsIterator.remove();
      } catch (ServiceResponseException e) {
        if (LOG.isTraceEnabled()) {
          LOG.trace("Item will be removed by  synchronization with exchange state: " + calendarEvent.getDescription());
        }
      }
    }
    return calendarEvents;
  }

  private List<Item> searchAllAppointments(FolderId parentFolderId) throws Exception {
    ItemView view = new ItemView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId, view);
    if (LOG.isTraceEnabled()) {
      LOG.trace("Exchange user calendar '" + username + "', items found: " + findResults.getTotalCount());
    }
    return findResults.getItems();
  }

  private List<Folder> searchSubFolders(FolderId parentFolderId) throws Exception {
    FolderView view = new FolderView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindFoldersResults findResults = service.findFolders(parentFolderId, view);
    return findResults.getFolders();
  }

  private List<CalendarEvent> searchCalendarEventsModifiedSince(Calendar calendar, Date date) throws Exception {
    if (date == null) {
      return exoStorageService.searchAllEvents(username, calendar);
    }
    return exoStorageService.searchEventsModifiedSince(username, calendar, date);
  }

  private List<Item> searchAllAppointmentsModifiedSince(FolderId parentFolderId, Date date, int diffTimeZone) throws Exception {
    if (date == null) {
      return searchAllAppointments(parentFolderId);
    }

    // Exchange system dates are saved using UTC timezone independing of User
    // Calendar timezone, so we have to get the diff with eXo Server TimeZone
    // and Exchange to make search queries
    java.util.Calendar calendar = java.util.Calendar.getInstance();
    calendar.setTime(date);
    calendar.add(java.util.Calendar.HOUR_OF_DAY, diffTimeZone);
    calendar.add(java.util.Calendar.SECOND, 1);

    ItemView view = new ItemView(100);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId, new SearchFilter.IsGreaterThan(ItemSchema.LastModifiedTime, calendar.getTime()), view);
    if (LOG.isTraceEnabled()) {
      LOG.trace("Exchange user calendar '" + username + "', items found: " + findResults.getTotalCount());
    }
    return findResults.getItems();
  }

  private static Date convertDateToUTC(Date date) throws ParseException {
    dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
    String time = dateFormat.format(date);
    dateFormat.setTimeZone(TimeZone.getDefault());
    return dateFormat.parse(time);
  }

}
