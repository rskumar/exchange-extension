package org.exoplatform.extension.exchange.service;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
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
public class ExchangeIntegrationService {

  private final static Log LOG = ExoLogger.getLogger(ExchangeService.class);

  private static final String USER_EXCHANGE_HANDLED_ATTRIBUTE = "exchange.check.date";

  private final String username;
  private final ExchangeService service;
  private final CalendarStorageService calendarStorageService;
  private final OrganizationService organizationService;

  public ExchangeIntegrationService(OrganizationService organizationService, CalendarStorageService calendarStorageService, ExchangeService service, String username) {
    this.calendarStorageService = calendarStorageService;
    this.organizationService = organizationService;
    this.service = service;
    this.username = username;
  }

  /**
   * 
   * Synchronize Exchange Calendar identified by 'folderId' with eXo Calendar.
   * 
   * @param folderId
   * @param lastSyncDate
   * @param diffTimeZone
   * @throws Exception
   */
  public void synchronizeFullCalendar(FolderId folderId) throws Exception {
    CalendarFolder folder = CalendarFolder.bind(service, folderId);
    calendarStorageService.getOrCreateUserCalendar(username, folder);

    Iterable<Item> items = searchAllAppointments(folderId);
    for (Item item : items) {
      if (item instanceof Appointment) {
        calendarStorageService.createOrUpdateEvent((Appointment) item, folder, username);
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
  }

  public static void main(String[] args) {
    System.out.println(TimeZone.getTimeZone("UTC").getRawOffset() / 3600000);
    System.out.println(TimeZone.getDefault().getRawOffset() / 3600000);
    System.out.println(TimeZone.getDefault().useDaylightTime());
    System.out.println(TimeZone.getDefault().inDaylightTime(new Date()));
    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS");
    sdf.setTimeZone(TimeZone.getDefault());
    System.out.println(sdf.format(new Date()));
    sdf.setTimeZone(TimeZone.getTimeZone("UTC"));
    System.out.println(sdf.format(new Date()));
  }

  /**
   * 
   * Synchronize Exchange Calendar identified by 'folderId' with eXo Calendar.
   * The check is done for events modified since 'lastSyncDate'.
   * 
   * @param folderId
   * @param lastSyncDate
   * @param diffTimeZone
   * @throws Exception
   */
  public void synchronizeModificationsOfCalendar(FolderId folderId, Date lastSyncDate, int diffTimeZone) throws Exception {
    Iterable<Item> items = searchAllAppointmentsModifiedSince(folderId, lastSyncDate, diffTimeZone);
    // Search for modified Appointments in Exchange, since last check date.
    for (Item item : items) {
      if (item instanceof Appointment) {
        Folder folder = Folder.bind(service, item.getParentFolderId());
        calendarStorageService.createOrUpdateEvent((Appointment) item, folder, username);
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
    // Search for deleted Appointments from Exchange
    Iterable<CalendarEvent> deletedEvents = searchDeletedAppointments(username, folderId.getUniqueId());
    for (CalendarEvent calendarEvent : deletedEvents) {
      LOG.info("Delete user calendar event: " + calendarEvent.getSummary());
      calendarStorageService.deleteEvent(calendarEvent, folderId.getUniqueId(), username);
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
    calendarStorageService.deleteUnregisteredExchangeCalendars(username, folderIds);
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
    return calendarStorageService.getUserCalendar(username, folderId.getUniqueId()) != null;
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
    return calendarStorageService.deleteCalendar(username, folderId.getUniqueId());
  }

  /**
   * 
   * Creates or updates or deletes eXo Calendar Event associated to Item, switch
   * state in Exchange.
   * 
   * @param itemEvent
   * @throws Exception
   */
  public void createOrUpdateOrDelete(ItemEvent itemEvent) throws Exception {
    Item item = null;
    try {
      item = Appointment.bind(service, itemEvent.getItemId());
    } catch (ServiceResponseException e) {
      if (LOG.isDebugEnabled()) { // Just in case
        LOG.debug("Item was not bound, it was deleted:" + itemEvent.getItemId());
      }
    }
    if (item == null) {
      LOG.info("Delete user calendar event: " + itemEvent.getItemId().getUniqueId());
      Folder folder = Folder.bind(service, itemEvent.getParentFolderId());
      calendarStorageService.deleteEvent(itemEvent.getItemId().getUniqueId(), folder.getId().getUniqueId(), username);
    } else if (item instanceof Appointment) {
      Folder folder = Folder.bind(service, item.getParentFolderId());
      if (itemEvent.getEventType() == EventType.Modified) {
        LOG.info("Update user calendar event: " + itemEvent.getItemId().getUniqueId());
        calendarStorageService.updateEvent((Appointment) item, folder, username);
      } else if (itemEvent.getEventType() == EventType.Created) {
        LOG.info("Create user calendar event: " + itemEvent.getItemId().getUniqueId());
        calendarStorageService.createEvent((Appointment) item, folder, username);
      }
    } else {
      LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
    }
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

  private Iterable<CalendarEvent> searchDeletedAppointments(String username, String folderId) throws Exception {
    List<CalendarEvent> calendarEvents = calendarStorageService.getUserCalendarEvents(username, folderId);
    Iterator<CalendarEvent> calendarEventsIterator = calendarEvents.iterator();
    while (calendarEventsIterator.hasNext()) {
      CalendarEvent calendarEvent = calendarEventsIterator.next();
      String itemId = calendarEvent.getMessage();
      if (itemId == null) {
        LOG.warn("Can't synchronize state of Calendar Event : " + calendarEvent.getDescription());
        continue;
      }
      try {
        Appointment.bind(service, ItemId.getItemIdFromString(itemId));
        calendarEventsIterator.remove();
      } catch (ServiceResponseException e) {
        if (LOG.isDebugEnabled()) {
          LOG.debug("Item will be removed by  synchronization with exchange state: " + calendarEvent.getDescription());
        }
      }
    }
    return calendarEvents;
  }

  private List<Item> searchAllAppointments(FolderId parentFolderId) throws Exception {
    ItemView view = new ItemView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId, view);
    LOG.info("Exchange user calendar '" + username + "', items found: " + findResults.getTotalCount());
    return findResults.getItems();
  }

  private List<Folder> searchSubFolders(FolderId parentFolderId) throws Exception {
    FolderView view = new FolderView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindFoldersResults findResults = service.findFolders(parentFolderId, view);
    return findResults.getFolders();
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
    calendar.set(java.util.Calendar.HOUR_OF_DAY, calendar.get(java.util.Calendar.HOUR_OF_DAY) + diffTimeZone);

    ItemView view = new ItemView(100);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId, new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.LastModifiedTime, calendar.getTime()), view);
    LOG.info("Exchange user calendar '" + username + "', items found: " + findResults.getTotalCount());
    return findResults.getItems();
  }

  public Calendar getUserCalendar(FolderId folderId) throws Exception {
    return calendarStorageService.getUserCalendar(username, folderId.getUniqueId());
  }

}
