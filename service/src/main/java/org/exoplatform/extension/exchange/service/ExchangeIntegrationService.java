package org.exoplatform.extension.exchange.service;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

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

import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.extension.exchange.service.util.CalendarConverterService;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;

public class ExchangeIntegrationService {

  private final static Log LOG = ExoLogger.getLogger(ExchangeService.class);

  private final String username;
  private final ExchangeService service;
  private final CalendarStorageService calendarStorageService;

  public ExchangeIntegrationService(CalendarStorageService calendarStorageService, ExchangeService service, String username) {
    this.calendarStorageService = calendarStorageService;
    this.service = service;
    this.username = username;
  }

  public Iterable<Item> synchronizeModificationsOfCalendar(FolderId folderId, Date lastSyncDate, int diffTimeZone) throws Exception {
    Iterable<Item> items = searchAllAppointmentsModifiedSince(folderId, lastSyncDate, diffTimeZone);
    for (Item item : items) {
      if (item instanceof Appointment) {
        Folder folder = Folder.bind(service, item.getParentFolderId());
        calendarStorageService.createOrUpdateEvent((Appointment) item, folder, username);
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
    Iterable<CalendarEvent> deletedEvents = searchDeletedAppointments(username, folderId.getUniqueId());
    for (CalendarEvent calendarEvent : deletedEvents) {
      LOG.info("Delete user calendar event: " + calendarEvent.getSummary());
      calendarStorageService.deleteEvent(calendarEvent, folderId.getUniqueId(), username);
    }
    return items;
  }

  public Iterable<CalendarEvent> searchDeletedAppointments(String username, String folderId) throws Exception {
    List<CalendarEvent> calendarEvents = calendarStorageService.findUserCalendarEvents(username, folderId);
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

  public String synchronizeFullCalendar(FolderId folderId) throws Exception {
    Iterable<Item> items = searchAllAppointments(folderId);
    for (Item item : items) {
      if (item instanceof Appointment) {
        Folder folder = Folder.bind(service, item.getParentFolderId());
        calendarStorageService.createOrUpdateEvent((Appointment) item, folder, username);
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
    return CalendarConverterService.getCalendarId(username, folderId.getUniqueId());
  }

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

  public List<Folder> searchSubFolders(FolderId parentFolderId) throws Exception {
    FolderView view = new FolderView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindFoldersResults findResults = service.findFolders(parentFolderId, view);
    return findResults.getFolders();
  }

  public List<Item> searchAllAppointments(FolderId parentFolderId) throws Exception {
    ItemView view = new ItemView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId, view);
    LOG.info("Exchange user calendar '" + username + "', items found: " + findResults.getTotalCount());
    return findResults.getItems();
  }

  public List<Item> searchAllAppointmentsModifiedSince(FolderId parentFolderId, Date date, int diffTimeZone) throws Exception {
    if (date == null) {
      return searchAllAppointments(parentFolderId);
    }
    // TODO delete those lines once diff timezone calculation become automatic
    java.util.Calendar calendar = java.util.Calendar.getInstance();
    calendar.setTime(date);
    calendar.set(java.util.Calendar.HOUR_OF_DAY, calendar.get(java.util.Calendar.HOUR_OF_DAY) + diffTimeZone);

    ItemView view = new ItemView(100);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId, new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.LastModifiedTime, calendar.getTime()), view);
    LOG.info("Exchange user calendar '" + username + "', items found: " + findResults.getTotalCount());
    return findResults.getItems();
  }

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

  public boolean isCalendarPresent(String username, String folderId) throws Exception {
    return calendarStorageService.findUserCalendar(username, folderId) != null;
  }
}
