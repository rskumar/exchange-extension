package org.exoplatform.extension.exchange.service;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;

import javax.jcr.Node;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.BasePropertySet;
import microsoft.exchange.webservices.data.CalendarFolder;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.FindItemsResults;
import microsoft.exchange.webservices.data.Folder;
import microsoft.exchange.webservices.data.FolderId;
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
import org.exoplatform.calendar.service.impl.CalendarServiceImpl;
import org.exoplatform.container.PortalContainer;
import org.exoplatform.container.component.ComponentRequestLifecycle;
import org.exoplatform.extension.exchange.service.util.CalendarConverterService;
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

  public static final String USER_EXCHANGE_SERVER_URL_ATTRIBUTE = "exchange.server.url";
  public static final String USER_EXCHANGE_SERVER_DOMAIN_ATTRIBUTE = "exchange.server.domain";
  public static final String USER_EXCHANGE_USERNAME_ATTRIBUTE = "exchange.username";
  public static final String USER_EXCHANGE_PASSWORD_ATTRIBUTE = "exchange.password";

  private final static Log LOG = ExoLogger.getLogger(IntegrationService.class);

  private static final String USER_EXCHANGE_HANDLED_ATTRIBUTE = "exchange.check.date";
  private static final String USER_EXO_HANDLED_ATTRIBUTE = "exo.check.date";
  private static final Map<String, IntegrationService> instances = new HashMap<String, IntegrationService>();

  private final String username;
  private final ExchangeService service;
  private final ExoStorageService exoStorageService;
  private final ExchangeStorageService exchangeStorageService;
  private final CorrespondenceService correspondenceService;
  private final OrganizationService organizationService;
  private final CalendarService calendarService;

  private boolean synchIsCurrentlyRunning = false;

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
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public CalendarFolder getExchangeCalendar(String folderId) throws Exception {
    CalendarFolder folder = null;
    try {
      folder = CalendarFolder.bind(service, FolderId.getFolderIdFromString(folderId));
    } catch (ServiceResponseException e) {
      LOG.warn("Can't get Folder identified by id: " + folderId);
    }
    return folder;
  }

  /**
   * 
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public CalendarFolder getExchangeCalendar(FolderId folderId) throws Exception {
    CalendarFolder folder = null;
    try {
      folder = CalendarFolder.bind(service, folderId);
    } catch (ServiceResponseException e) {
      LOG.warn("Can't get Folder identified by id: " + folderId.getUniqueId());
    }
    return folder;
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
    List<String> eventIds = new ArrayList<String>();
    CalendarFolder folder = null;
    try {
      folder = CalendarFolder.bind(service, folderId);
    } catch (ServiceResponseException e) {
      LOG.warn("Can't get Folder identified by id: " + folderId.getUniqueId());
      return eventIds;
    }
    // Create Calendar if not present
    exoStorageService.getOrCreateUserCalendar(username, folder);

    Iterable<Item> items = searchAllItems(folderId);
    for (Item item : items) {
      if (item instanceof Appointment) {
        List<CalendarEvent> updatedEvents = exoStorageService.createOrUpdateEvent((Appointment) item, folder, username, getUserExoCalenarTimeZoneSetting());
        if (updatedEvents != null && !updatedEvents.isEmpty()) {
          for (CalendarEvent calendarEvent : updatedEvents) {
            eventIds.add(calendarEvent.getId());
          }
        }
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }

    List<CalendarEvent> events = exoStorageService.getUserCalendarEvents(username, folderId.getUniqueId());
    for (CalendarEvent calendarEvent : events) {
      if (correspondenceService.getCorrespondingId(username, calendarEvent.getId()) == null) {
        exoStorageService.deleteEvent(username, calendarEvent);
      } else {
        String itemId = correspondenceService.getCorrespondingId(username, calendarEvent.getId());
        Item item = null;
        try {
          item = Item.bind(service, ItemId.getItemIdFromString(itemId));
        } catch (Exception e) {
          item = null;
        }
        if (item == null) {
          exoStorageService.deleteEvent(username, calendarEvent);
        }
      }
    }

    return eventIds;
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
    // Serach modified eXo Calendar events since this date, this is used to
    // force synchronization
    Date exoLastSyncDate = getUserExoLastCheckDate();
    if (exoLastSyncDate == null) {
      exoLastSyncDate = lastSyncDate;
    }

    Iterable<Item> items = searchAllAppointmentsModifiedSince(folderId, lastSyncDate, diffTimeZone);
    // Search for modified Appointments in Exchange, since last check date.
    Folder folder = null;
    for (Item item : items) {
      if (item instanceof Appointment) {
        folder = Folder.bind(service, item.getParentFolderId());
        // Test if there is a modification conflict
        CalendarEvent event = exoStorageService.getEventByAppointmentId(username, item.getId().getUniqueId());
        if (event != null) {
          if (updatedExoEventIDs != null && updatedExoEventIDs.contains(event.getId())) {
            // Already updated by previous operation
            continue;
          }
          Date eventModifDate = CalendarConverterService.convertDateToUTC(event.getLastUpdatedTime());
          Date itemModifDate = item.getLastModifiedTime();
          if (itemModifDate.after(eventModifDate)) {
            List<CalendarEvent> updatedEvents = exoStorageService.updateEvent((Appointment) item, folder, username, getUserExoCalenarTimeZoneSetting());
            if (updatedEvents != null && !updatedEvents.isEmpty() && updatedExoEventIDs != null) {
              for (CalendarEvent calendarEvent : updatedEvents) {
                updatedExoEventIDs.add(calendarEvent.getId());
              }
            }
          }
        } else {
          List<CalendarEvent> updatedEvents = exoStorageService.createEvent((Appointment) item, folder, username, getUserExoCalenarTimeZoneSetting());
          if (updatedEvents != null && !updatedEvents.isEmpty() && updatedExoEventIDs != null) {
            for (CalendarEvent calendarEvent : updatedEvents) {
              updatedExoEventIDs.add(calendarEvent.getId());
            }
          }
        }
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
    // Search for existant Appointments in Exchange but not in eXo
    Iterable<CalendarEvent> unsynchronizedEvents = searchUnsynchronizedAppointments(username, folderId.getUniqueId());
    for (CalendarEvent calendarEvent : unsynchronizedEvents) {
      // To not have redendance
      if (updatedExoEventIDs != null && updatedExoEventIDs.contains(calendarEvent.getId())) {
        continue;
      }
      if (calendarEvent.getLastUpdatedTime() != null) {
        if (calendarEvent.getLastUpdatedTime().after(exoLastSyncDate)) {
          String exoMasterId = null;
          if (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence()) {
            exoMasterId = exoStorageService.getRecurrentEventIdByOriginalUUID(calendarEvent.getOriginalReference());
            if (exoMasterId == null) {
              LOG.error("No master Id was found for occurence: " + calendarEvent.getSummary() + " with recurrenceId = " + calendarEvent.getRecurrenceId() + ". The event will not be updated.");
              continue;
            }
          }
          boolean deleteEvent = exchangeStorageService.updateOrCreateExchangeAppointment(username, service, calendarEvent, exoMasterId, getUserExoCalenarTimeZoneSetting(), null);
          if (deleteEvent) {
            exoStorageService.deleteEvent(username, calendarEvent);
          }
          if (updatedExoEventIDs != null) {
            updatedExoEventIDs.add(calendarEvent.getId());
          }
        } else {
          exoStorageService.deleteEvent(username, calendarEvent);
        }
      }
    }

    List<CalendarEvent> modifiedCalendarEvents = searchCalendarEventsModifiedSince(getUserCalendarByExchangeFolderId(folderId), exoLastSyncDate);
    for (CalendarEvent calendarEvent : modifiedCalendarEvents) {
      // If modified with synchronization, ignore
      if (updatedExoEventIDs.contains(calendarEvent.getId())) {
        continue;
      }
      String exoMasterId = null;
      if (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence()) {
        exoMasterId = exoStorageService.getRecurrentEventIdByOriginalUUID(calendarEvent.getOriginalReference());
        if (exoMasterId == null) {
          LOG.error("No master Id was found for occurence: " + calendarEvent.getSummary() + " with recurrenceId = " + calendarEvent.getRecurrenceId() + ". The event will not be updated.");
        }
      }
      boolean deleteEvent = exchangeStorageService.updateOrCreateExchangeAppointment(username, service, calendarEvent, exoMasterId, getUserExoCalenarTimeZoneSetting(), null);
      if (deleteEvent) {
        exoStorageService.deleteEvent(username, calendarEvent);
      }
      updatedExoEventIDs.add(calendarEvent.getId());
    }
  }

  /**
   * 
   * Gets list of personnal Exchange Calendars.
   * 
   * @return list of FolderId
   * @throws Exception
   */
  public List<FolderId> getAllExchangeCalendars() throws Exception {
    return exchangeStorageService.getAllExchangeCalendars(service);
  }

  /**
   * Checks if eXo associated Calendar is present.
   * 
   * @param usename
   * @return true if present.
   * @throws Exception
   */
  public boolean isCalendarSynchronized(String id) throws Exception {
    return correspondenceService.getCorrespondingId(username, id) != null;
  }

  /**
   * Checks if eXo associated Calendar is present.
   * 
   * @param usename
   * @return true if present.
   * @throws Exception
   */
  public boolean isCalendarPresentInExo(FolderId folderId) throws Exception {
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
  public List<CalendarEvent> createOrUpdateOrDelete(ItemEvent itemEvent) throws Exception {
    Appointment appointment = null;
    try {
      appointment = Appointment.bind(service, itemEvent.getItemId());
    } catch (ServiceResponseException e) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Item was not bound, it was deleted:" + itemEvent.getItemId());
      }
    }
    List<CalendarEvent> updatedEvents = null;
    if (appointment == null) {
      exoStorageService.deleteEventByAppointmentID(itemEvent.getItemId().getUniqueId(), username);
    } else {
      Folder folder = Folder.bind(service, appointment.getParentFolderId());
      String eventId = correspondenceService.getCorrespondingId(username, appointment.getId().getUniqueId());
      if (eventId == null) {
        updatedEvents = exoStorageService.createEvent((Appointment) appointment, folder, username, getUserExoCalenarTimeZoneSetting());
      } else {
        updatedEvents = exoStorageService.updateEvent((Appointment) appointment, folder, username, getUserExoCalenarTimeZoneSetting());
      }
    }
    return updatedEvents;
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

  /**
   * 
   * @param eventId
   * @throws Exception
   */
  public void updateOrCreateExchangeCalendarEvent(Node eventNode) throws Exception {
    CalendarEvent event = exoStorageService.getEvent(eventNode);
    if (isCalendarSynchronized(event.getCalendarId())) {
      List<CalendarEvent> calendarEventsToUpdateModifiedTime = new ArrayList<CalendarEvent>();
      updateOrCreateExchangeCalendarEvent(event, calendarEventsToUpdateModifiedTime);
      if (!calendarEventsToUpdateModifiedTime.isEmpty()) {
        // This is done to not have a cyclic updates between eXo and Exchange
        exoStorageService.updateModifiedDateOfEvent(username, event);
      }
    }
  }

  /**
   * 
   * @param eventId
   * @throws Exception
   */
  public boolean updateOrCreateExchangeCalendarEvent(String eventId) throws Exception {
    CalendarEvent event = ((CalendarServiceImpl) calendarService).getDataStorage().getEvent(username, eventId);
    return updateOrCreateExchangeCalendarEvent(event);
  }

  /**
   * 
   * @param event
   * @throws Exception
   */
  public boolean updateOrCreateExchangeCalendarEvent(CalendarEvent event, List<CalendarEvent> eventsToUpdate) throws Exception {
    String exoMasterId = null;
    if (event.getIsExceptionOccurrence() != null && event.getIsExceptionOccurrence()) {
      exoMasterId = exoStorageService.getRecurrentEventIdByOriginalUUID(event.getOriginalReference());
      if (exoMasterId == null) {
        LOG.error("No master Id was found for occurence: " + event.getSummary() + " with recurrenceId = " + event.getRecurrenceId() + ". The event will not be updated.");
      }
    }
    return exchangeStorageService.updateOrCreateExchangeAppointment(username, service, event, exoMasterId, getUserExoCalenarTimeZoneSetting(), eventsToUpdate);
  }

  /**
   * 
   * @param event
   * @throws Exception
   */
  public boolean updateOrCreateExchangeCalendarEvent(CalendarEvent event) throws Exception {
    return updateOrCreateExchangeCalendarEvent(event, null);
  }

  /**
   * 
   * @param eventId
   * @throws Exception
   */
  public void deleteExchangeCalendarEvent(String eventId, String calendarId) throws Exception {
    exchangeStorageService.deleteExchangeAppointment(username, service, eventId, calendarId);
  }

  /**
   * 
   * @param calendarId
   * @throws Exception
   */
  public void deleteExchangeCalendar(String calendarId) throws Exception {
    exchangeStorageService.deleteExchangeFolderByCalenarId(username, service, calendarId);
  }

  public void removeInstance() {
    LOG.info("Stop Exchange Integration Service for user: " + username);
    instances.remove(username);
  }

  public List<String> synchronizeExchangeFolderState(List<FolderId> calendarFolderIds, boolean synchronizeAllExchangeFolders, boolean deleteExoCalendarOnUnsync) throws Exception {
    Iterator<FolderId> iterator = calendarFolderIds.iterator();
    while (iterator.hasNext()) {
      FolderId folderId = (FolderId) iterator.next();
      try {
        CalendarFolder.bind(service, folderId);
      } catch (Exception e) {
        // Test if the connection is ok, else the exception is thrown because of
        // interrupted connection
        CalendarFolder.bind(service, FolderId.getFolderIdFromWellKnownFolderName(WellKnownFolderName.Calendar));

        Calendar calendar = exoStorageService.getUserCalendar(username, folderId.getUniqueId());
        if (calendar != null) {
          LOG.info("Folder '" + folderId.getUniqueId() + "' was deleted from Exchange, stopping synchronization for this folder.");
          if (deleteExoCalendarOnUnsync) {
            exoStorageService.deleteCalendar(username, folderId.getUniqueId());
          } else {
            correspondenceService.deleteCorrespondingId(username, folderId.getUniqueId());
          }
          // Remove FolderId from synchronized folders in Scheduled Job
          iterator.remove();
        }
      }
    }

    List<String> updatedCalendarEventIds = null;
    // synchronize added Folders
    if (synchronizeAllExchangeFolders) {
      List<FolderId> folderIds = exchangeStorageService.getAllExchangeCalendars(service);
      for (FolderId folderId : folderIds) {
        if (!calendarFolderIds.contains(folderId)) {
          updatedCalendarEventIds = synchronizeFullCalendar(folderId);
          calendarFolderIds.add(folderId);
        }
      }
    } else {
      // Checks if synchronized Exchange Folders have been modified
      List<FolderId> folderIds = getSynchronizedExchangeCalendars();
      for (FolderId folderId : folderIds) {
        if (!calendarFolderIds.contains(folderId)) {
          updatedCalendarEventIds = synchronizeFullCalendar(folderId);
          calendarFolderIds.add(folderId);
        }
      }
      Iterator<FolderId> folderIdIterator = calendarFolderIds.iterator();
      while (folderIdIterator.hasNext()) {
        FolderId folderId = (FolderId) folderIdIterator.next();
        if (!folderIds.contains(folderId)) {
          folderIdIterator.remove();
        }
      }
    }
    return updatedCalendarEventIds;
  }

  public List<FolderId> getSynchronizedExchangeCalendars() throws Exception {
    List<FolderId> folderIds = new ArrayList<FolderId>();
    List<String> folderIdsString = correspondenceService.getSynchronizedExchangeFolderIds(username);
    for (String folderIdString : folderIdsString) {
      folderIds.add(FolderId.getFolderIdFromString(folderIdString));
    }
    return folderIds;
  }

  public void addFolderToSynchronization(String folderIdString) throws Exception {
    String calendarId = CalendarConverterService.getCalendarId(folderIdString);
    correspondenceService.setCorrespondingId(username, calendarId, folderIdString);
  }

  public void deleteFolderFromSynchronization(String folderIdString) throws Exception {
    correspondenceService.deleteCorrespondingId(username, folderIdString);
  }

  public synchronized void setSynchronizationStarted() {
    synchIsCurrentlyRunning = true;
  }

  public synchronized void setSynchronizationStopped() {
    synchIsCurrentlyRunning = false;
  }

  public synchronized boolean isSynchronizationStarted() {
    return synchIsCurrentlyRunning;
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
        // Item was detected, and will be created
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

  private List<Item> searchAllItems(FolderId parentFolderId) throws Exception {
    ItemView view = new ItemView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId, view);
    if (LOG.isTraceEnabled()) {
      LOG.trace("Exchange user calendar '" + username + "', items found: " + findResults.getTotalCount());
    }
    return findResults.getItems();
  }

  private List<CalendarEvent> searchCalendarEventsModifiedSince(Calendar calendar, Date date) throws Exception {
    if (date == null) {
      return exoStorageService.searchAllEvents(username, calendar);
    }
    return exoStorageService.searchEventsModifiedSince(username, calendar, date);
  }

  private List<Item> searchAllAppointmentsModifiedSince(FolderId parentFolderId, Date date, int diffTimeZone) throws Exception {
    if (date == null) {
      return searchAllItems(parentFolderId);
    }

    // Exchange system dates are saved using UTC timezone independing of User
    // Calendar timezone, so we have to get the diff with eXo Server TimeZone
    // and Exchange to make search queries
    java.util.Calendar calendar = java.util.Calendar.getInstance();
    calendar.setTime(date);
    calendar.add(java.util.Calendar.MINUTE, diffTimeZone);
    calendar.add(java.util.Calendar.SECOND, 1);

    ItemView view = new ItemView(100);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId, new SearchFilter.IsGreaterThan(ItemSchema.LastModifiedTime, calendar.getTime()), view);
    if (LOG.isTraceEnabled()) {
      LOG.trace("Exchange user calendar '" + username + "', items found: " + findResults.getTotalCount());
    }
    return findResults.getItems();
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
    long savedTime = userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE) == null ? 0 : Long.valueOf(userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE));
    if (time > savedTime) {
      userProfile.setAttribute(USER_EXO_HANDLED_ATTRIBUTE, "" + time);
    }
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
   * Sets last check operation date.
   * 
   * @param username
   * @param time
   * @throws Exception
   */
  public void setUserExoLastCheckDate(long time) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
    long savedTime = userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE) == null ? 0 : Long.valueOf(userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE));
    if (savedTime <= 0) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("User '" + username + "' exo last check time was not set before, may be the synhronization was not run before or an error occured in the meantime.");
      }
    } else {
      userProfile.setAttribute(USER_EXO_HANDLED_ATTRIBUTE, "" + time);
      organizationService.getUserProfileHandler().saveUserProfile(userProfile, false);
    }
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
  public Date getUserExoLastCheckDate() throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
    long time = userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE) == null ? 0 : Long.valueOf(userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE));
    Date lastSyncDate = null;
    if (time > 0) {
      lastSyncDate = new Date(time);
    }
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
    }
    return lastSyncDate;
  }

  public ExchangeService getService() {
    return service;
  }

  /**
   * 
   * @param name
   * @param value
   * @throws Exception
   */
  public void setUserArrtibute(String name, String value) throws Exception {
    setUserArrtibute(organizationService, username, name, value);
  }

  /**
   * 
   * @param name
   * @return
   * @throws Exception
   */
  public String getUserArrtibute(String name) throws Exception {
    return getUserArrtibute(organizationService, username, name);
  }

  /**
   * 
   * @param name
   * @param value
   * @throws Exception
   */
  public static void setUserArrtibute(OrganizationService organizationService, String username, String name, String value) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
    userProfile.setAttribute(name, value);
    organizationService.getUserProfileHandler().saveUserProfile(userProfile, false);
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
    }
  }

  /**
   * 
   * @param name
   * @return
   * @throws Exception
   */
  public static String getUserArrtibute(OrganizationService organizationService, String username, String name) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    try {
      UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
      return userProfile.getAttribute(name);
    } finally {
      if (organizationService instanceof ComponentRequestLifecycle) {
        ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
      }
    }
  }
}
