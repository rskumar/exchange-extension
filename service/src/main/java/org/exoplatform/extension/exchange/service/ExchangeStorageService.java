package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.TimeZone;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.AppointmentSchema;
import microsoft.exchange.webservices.data.BasePropertySet;
import microsoft.exchange.webservices.data.CalendarFolder;
import microsoft.exchange.webservices.data.ConflictResolutionMode;
import microsoft.exchange.webservices.data.DeleteMode;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.FindFoldersResults;
import microsoft.exchange.webservices.data.Folder;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.FolderView;
import microsoft.exchange.webservices.data.ItemId;
import microsoft.exchange.webservices.data.PropertySet;
import microsoft.exchange.webservices.data.ServiceResponseException;
import microsoft.exchange.webservices.data.TimeZoneDefinition;
import microsoft.exchange.webservices.data.WellKnownFolderName;

import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.extension.exchange.service.util.CalendarConverterService;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;

/**
 * 
 * @author Boubaker Khanfir
 * 
 */
public class ExchangeStorageService implements Serializable {
  private static final long serialVersionUID = 6348129698208975430L;

  private final static Log LOG = ExoLogger.getLogger(ExchangeStorageService.class);

  private OrganizationService organizationService;
  private CorrespondenceService correspondenceService;

  public ExchangeStorageService(OrganizationService organizationService, CorrespondenceService correspondenceService) {
    this.organizationService = organizationService;
    this.correspondenceService = correspondenceService;
  }

  /**
   * 
   * Gets list of personnal Exchange Calendars.
   * 
   * @return list of FolderId
   * @throws Exception
   */
  public List<FolderId> getAllExchangeCalendars(ExchangeService service) throws Exception {
    List<FolderId> calendarFolderIds = new ArrayList<FolderId>();
    CalendarFolder calendarRootFolder = CalendarFolder.bind(service, WellKnownFolderName.Calendar);

    calendarFolderIds.add(calendarRootFolder.getId());
    List<Folder> calendarfolders = searchSubFolders(service, calendarRootFolder.getId());

    if (calendarfolders != null && !calendarfolders.isEmpty()) {
      for (Folder tmpFolder : calendarfolders) {
        calendarFolderIds.add(tmpFolder.getId());
      }
    }
    return calendarFolderIds;
  }

  /**
   * 
   * @param username
   * @param service
   * @param event
   * @param exoMasterId
   * @param userCalendarTimeZone
   * @return true if the CalenarEvent have to be deleted
   * @throws Exception
   */
  public boolean updateOrCreateExchangeAppointment(String username, ExchangeService service, CalendarEvent event, String exoMasterId, TimeZone userCalendarTimeZone,
      List<CalendarEvent> eventsToUpdateModifiedTime) throws Exception {
    if (event == null) {
      return false;
    }
    String folderIdString = correspondenceService.getCorrespondingId(username, event.getCalendarId());
    if (folderIdString == null || folderIdString.isEmpty()) {
      LOG.trace("eXo Calendar with id '" + event.getCalendarId() + "' is not synhronized with Exchange, ignore Event:" + event.getSummary());
      return false;
    }

    String itemId = correspondenceService.getCorrespondingId(username, event.getId());
    boolean isNew = true;
    Appointment appointment = null;
    if (itemId != null) {
      try {
        appointment = Appointment.bind(service, ItemId.getItemIdFromString(itemId));
        isNew = false;
      } catch (ServiceResponseException e) {
        if (LOG.isTraceEnabled()) {
          LOG.trace("Item was not bound, it was deleted or not yet created:" + event.getId());
        }
        correspondenceService.deleteCorrespondingId(username, event.getId());
      }
    }

    if (event.getRecurrenceId() == null && (event.getRepeatType() == null || event.getRepeatType().equals(CalendarEvent.RP_NOREPEAT))) {
      if (isNew) {
        // Checks if this event was already in Exchange, if it's the case, it
        // means that the item was not found because the user has removed it
        // from
        // Exchange
        if (CalendarConverterService.isExchangeEventId(event.getId())) {
          LOG.error("Conflict in modification, inconsistant data, the event was deleted in Exchange but seems always in eXo, the event will be deleted from Exchange.");
          deleteExchangeAppointment(username, service, event.getId(), event.getCalendarId());
          return false;
        }
        appointment = new Appointment(service);
      }
      CalendarConverterService
          .convertExoToExchangeEvent(appointment, event, username, organizationService.getUserHandler(), getTimeZoneDefinition(service, userCalendarTimeZone), userCalendarTimeZone);
    } else {
      if ((event.getRecurrenceId() != null && !event.getRecurrenceId().isEmpty()) || (event.getIsExceptionOccurrence() != null && event.getIsExceptionOccurrence())) {
        if (isNew) {
          String exchangeMasterId = correspondenceService.getCorrespondingId(username, exoMasterId);
          appointment = getAppointmentOccurence(service, exchangeMasterId, event.getRecurrenceId());
          if (appointment == null) {
            LOG.error("Cannot find Appointment occurence '" + event.getSummary() + "' with recurenceId: " + event.getRecurrenceId() + ", delete event.");
            return true;
          }
          isNew = false;
        }
        // TODO make this exception occurence
        CalendarConverterService.convertExoToExchangeOccurenceEvent(appointment, event, username, organizationService.getUserHandler(), getTimeZoneDefinition(service, userCalendarTimeZone),
            userCalendarTimeZone);
      } else {
        if (isNew) {
          // Checks if this event was already in Exchange, if it's the case, it
          // means that the item was not found because the user has removed it
          // from
          // Exchange
          if (CalendarConverterService.isExchangeEventId(event.getId())) {
            LOG.error("Conflict in modification, inconsistant data, the event was deleted in Exchange but seems always in eXo, the event will be deleted from Exchange.");
            deleteExchangeAppointment(username, service, event.getId(), event.getCalendarId());
            return false;
          }
          appointment = new Appointment(service);
        }
        CalendarConverterService.convertExoToExchangeMasterRecurringCalendarEvent(appointment, event, username, organizationService.getUserHandler(),
            getTimeZoneDefinition(service, userCalendarTimeZone), userCalendarTimeZone);
      }
    }
    if (isNew) {
      LOG.info("Create Exchange Appointment: " + event.getSummary());
      FolderId folderId = FolderId.getFolderIdFromString(folderIdString);
      appointment.save(folderId);
    } else {
      LOG.info("Update Exchange Appointment: " + event.getSummary());
      appointment.update(ConflictResolutionMode.AlwaysOverwrite);
    }
    if (eventsToUpdateModifiedTime != null) {
      eventsToUpdateModifiedTime.add(event);
    }
    correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
    return false;
  }

  /**
   * 
   * @param username
   * @param service
   * @param eventId
   * @param calendarId
   * @throws Exception
   */
  public void deleteExchangeAppointment(String username, ExchangeService service, String eventId, String calendarId) throws Exception {
    String itemId = correspondenceService.getCorrespondingId(username, eventId);
    if (itemId == null) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("The event was deleted from eXo but seems don't have corresponding Event in Exchange, ignore.");
      }
    } else {
      // Verify that calendar is synchronized
      if (correspondenceService.getCorrespondingId(username, calendarId) == null) {
        LOG.warn("Calendar with id '" + calendarId + "' seems not synchronized with exchange.");
      } else {
        Appointment appointment = null;
        try {
          appointment = Appointment.bind(service, ItemId.getItemIdFromString(itemId));
          LOG.info("Delete Exchange appointment: " + appointment.getSubject());
          appointment.delete(DeleteMode.HardDelete);
        } catch (ServiceResponseException e) {
          if (LOG.isTraceEnabled()) {
            LOG.trace("Exchange Item was not bound, it was deleted or not yet created:" + eventId);
          }
        }
      }
      correspondenceService.deleteCorrespondingId(username, itemId, eventId);
    }
  }

  /**
   * 
   * @param username
   * @param service
   * @param calendarId
   * @throws Exception
   */
  public void deleteExchangeFolderByCalenarId(String username, ExchangeService service, String calendarId) throws Exception {
    if (CalendarConverterService.isExchangeCalendarId(calendarId)) {
      LOG.warn("Can't delete Exchange Calendar, because it was created on Exchange: " + calendarId);
      return;
    }
    String folderId = correspondenceService.getCorrespondingId(username, calendarId);
    if (folderId == null) {
      LOG.warn("Conflict in modification, inconsistant data, the Calendar was deleted from eXo but seems don't have corresponding Folder in Exchange, ignore.");
      return;
    } else {
      Folder folder = null;
      try {
        folder = Folder.bind(service, FolderId.getFolderIdFromString(folderId));
        LOG.trace("Delete Exchange folder: " + folder.getDisplayName());
        folder.delete(DeleteMode.MoveToDeletedItems);
      } catch (ServiceResponseException e) {
        if (LOG.isTraceEnabled()) {
          LOG.trace("Exchange Folder was not bound, it was deleted or not yet created:" + folderId);
        }
      }
      correspondenceService.deleteCorrespondingId(username, calendarId);
    }
  }

  private TimeZoneDefinition getTimeZoneDefinition(ExchangeService service, TimeZone userCalendarTimeZone) {
    TimeZoneDefinition serverTimeZoneDefinition = null;
    Iterator<TimeZoneDefinition> timeZoneDefinitions = service.getServerTimeZones().iterator();
    while (timeZoneDefinitions.hasNext()) {
      TimeZoneDefinition timeZoneDefinition = (TimeZoneDefinition) timeZoneDefinitions.next();
      if (timeZoneDefinition.getId().equals(userCalendarTimeZone.getID())) {
        serverTimeZoneDefinition = timeZoneDefinition;
        break;
      }
    }
    return serverTimeZoneDefinition;
  }

  private List<Folder> searchSubFolders(ExchangeService service, FolderId parentFolderId) throws Exception {
    FolderView view = new FolderView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindFoldersResults findResults = service.findFolders(parentFolderId, view);
    return findResults.getFolders();
  }

  private Appointment getAppointmentOccurence(ExchangeService service, String exchangeMasterId, String recurrenceId) throws Exception {
    Appointment appointment = null;
    ItemId exchangeMasterItemId = ItemId.getItemIdFromString(exchangeMasterId);
    Date occDate = CalendarConverterService.RECURRENCE_ID_FORMAT.parse(recurrenceId);
    {
      Calendar calendar = Calendar.getInstance();
      calendar.setTime(occDate);
      calendar.set(Calendar.HOUR_OF_DAY, 0);
      calendar.set(Calendar.MINUTE, 0);
      calendar.set(Calendar.SECOND, 0);
      calendar.set(Calendar.MILLISECOND, 0);
      occDate = calendar.getTime();
    }
    Appointment masterAppointment = Appointment.bind(service, exchangeMasterItemId, new PropertySet(AppointmentSchema.Recurrence));

    int i = 1;
    Date endDate = masterAppointment.getRecurrence().getEndDate();
    if (endDate != null && occDate.getTime() > endDate.getTime()) {
      return null;
    }

    Calendar indexCalendar = Calendar.getInstance();
    indexCalendar.setTime(masterAppointment.getRecurrence().getStartDate());

    boolean continueSearch = true;
    while (continueSearch) {
      Appointment tmpAppointment = null;
      try {
        tmpAppointment = Appointment.bindToOccurrence(service, exchangeMasterItemId, i, new PropertySet(AppointmentSchema.Start));
        Date date = CalendarConverterService.getExoDateFromExchangeFormat(tmpAppointment.getStart());
        if (CalendarConverterService.isSameDate(occDate, date)) {
          appointment = Appointment.bindToOccurrence(service, exchangeMasterItemId, i, new PropertySet(BasePropertySet.FirstClassProperties));
          continueSearch = false;
        }
      } catch (Exception e) {
        // Recurence not found, can be deleted from Exchange.
      }
      i++;
      indexCalendar.add(Calendar.DATE, 1);
      if (continueSearch && (occDate.before(indexCalendar.getTime()))) {
        continueSearch = false;
      }
    }

    return appointment;
  }
}