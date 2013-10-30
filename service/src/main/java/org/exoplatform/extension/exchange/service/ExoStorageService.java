package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.TimeZone;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.Folder;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.calendar.service.impl.CalendarServiceImpl;
import org.exoplatform.calendar.service.impl.JCRDataStorage;
import org.exoplatform.extension.exchange.service.util.CalendarConverterService;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;

/**
 * 
 * @author Boubaker Khanfir
 * 
 */
public class ExoStorageService implements Serializable {
  private static final long serialVersionUID = 6614108102985034995L;

  private final static Log LOG = ExoLogger.getLogger(ExoStorageService.class);

  private JCRDataStorage storage;
  private OrganizationService organizationService;
  private CorrespondenceService correspondenceService;

  public ExoStorageService(OrganizationService organizationService, CalendarService calendarService, CorrespondenceService correspondenceService) {
    this.storage = ((CalendarServiceImpl) calendarService).getDataStorage();
    this.organizationService = organizationService;
    this.correspondenceService = correspondenceService;
  }

  /**
   * 
   * Deletes eXo Calendar Event.
   * 
   * @param appointmentId
   * @param folderId
   * @param username
   * @throws Exception
   */
  protected void deleteEventByAppointmentID(String appointmentId, String folderId, String username) throws Exception {
    CalendarEvent calendarEvent = getEventByAppointmentId(username, appointmentId);
    if (calendarEvent != null) {
      deleteEvent(calendarEvent, folderId, username);
    }
  }

  /**
   * 
   * Deletes eXo Calendar Event.
   * 
   * @param calendarEvent
   * @param folderId
   * @param username
   * @throws Exception
   */
  protected void deleteEvent(CalendarEvent calendarEvent, String folderId, String username) throws Exception {
    Calendar calendar = getUserCalendar(username, folderId);
    if ((calendarEvent.getRepeatType() == null || calendarEvent.getRepeatType().equals(CalendarEvent.RP_NOREPEAT))
        && (calendarEvent.getIsExceptionOccurrence() == null || !calendarEvent.getIsExceptionOccurrence())) {
      storage.removeUserEvent(username, calendar.getId(), calendarEvent.getId());
    } else if (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence()) {
      storage.removeOccurrenceInstance(username, calendarEvent);
    } else {
      storage.removeRecurrenceSeries(username, calendarEvent);
    }

    // Remove correspondence between exo and exchange IDs
    correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
  }

  /**
   * 
   * Creates or updates eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @throws Exception
   */
  protected CalendarEvent createOrUpdateEvent(Appointment appointment, Folder folder, String username, TimeZone timeZone) throws Exception {
    boolean isNew = correspondenceService.getCorrespondingId(username, appointment.getId().getUniqueId()) == null;
    if (!isNew) {
      CalendarEvent event = getEventByAppointmentId(username, appointment.getId().getUniqueId());
      if (event == null) {
        isNew = true;
        correspondenceService.deleteCorrespondingId(username, appointment.getId().getUniqueId());
      }
    }
    return createOrUpdateEvent(appointment, folder, username, isNew, timeZone);
  }

  /**
   * 
   * Delete eXo Calendar.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  protected boolean deleteCalendar(String username, String folderId) throws Exception {
    Calendar calendar = storage.removeUserCalendar(username, CalendarConverterService.getCalendarId(folderId));
    if (calendar != null) {
      LOG.info("User Calendar" + calendar.getId() + " is deleted, because it was deleted from Exchange.");
    }
    correspondenceService.deleteCorrespondingId(username, folderId, CalendarConverterService.getCalendarId(folderId));
    return false;
  }

  /**
   * 
   * Gets User Calendar identified by Exchange folder Id.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  protected Calendar getUserCalendar(String username, String folderId) throws Exception {
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    if (calendarId == null) {
      calendarId = CalendarConverterService.getCalendarId(folderId);
    }
    Calendar calendar = storage.getUserCalendar(username, calendarId);
    if (calendar == null && calendarId != null) {
      correspondenceService.deleteCorrespondingId(username, folderId);
    }
    return calendar;
  }

  /**
   * 
   * Gets User Calendar identified by Exchange folder Id, or creates it if not
   * existing.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  protected Calendar getOrCreateUserCalendar(String username, Folder folder) throws Exception {
    Calendar calendar = getUserCalendar(username, folder.getId().getUniqueId());
    if (calendar == null) {
      LOG.info("Create user calendar from Exchange: " + folder.getDisplayName());

      calendar = new Calendar();
      calendar.setId(CalendarConverterService.getCalendarId(folder.getId().getUniqueId()));
      calendar.setName(CalendarConverterService.getCalendarName(folder.getDisplayName()));
      calendar.setCalendarOwner(username);
      calendar.setDataInit(false);
      calendar.setEditPermission(new String[] { "any read" });
      calendar.setCalendarColor(Calendar.COLORS[9]);
      storage.saveUserCalendar(username, calendar, true);

      // Set IDs correspondence
      correspondenceService.setCorrespondingId(username, calendar.getId(), folder.getId().getUniqueId());
    }
    return calendar;
  }

  /**
   * 
   * Gets Events from User Calendar identified by Exchange folder Id.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  protected List<CalendarEvent> getUserCalendarEvents(String username, String folderId) throws Exception {
    List<CalendarEvent> userEvents = null;
    String calendarId = CalendarConverterService.getCalendarId(folderId);
    Calendar calendar = storage.getUserCalendar(username, CalendarConverterService.getCalendarId(folderId));
    if (calendar != null) {
      List<String> calendarIds = new ArrayList<String>();
      calendarIds.add(calendarId);
      userEvents = storage.getUserEventByCalendar(username, calendarIds);
    }
    return userEvents;
  }

  /**
   * 
   * Updates existing eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @param timeZone 
   * @throws Exception
   */
  protected CalendarEvent updateEvent(Appointment appointment, Folder folder, String username, TimeZone timeZone) throws Exception {
    return createOrUpdateEvent(appointment, folder, username, false, timeZone);
  }

  /**
   * 
   * Create non existing eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @param timeZone 
   * @throws Exception
   */
  protected CalendarEvent createEvent(Appointment appointment, Folder folder, String username, TimeZone timeZone) throws Exception {
    return createOrUpdateEvent(appointment, folder, username, true, timeZone);
  }

  /**
   * 
   * Search for existing calendars in eXo but not in Exchange.
   * 
   * @param username
   * @param folderIds
   * @throws Exception
   */
  protected void deleteUnregisteredExchangeCalendars(String username, List<String> folderIds) throws Exception {
    List<String> registeredCalendarIds = new ArrayList<String>();
    for (String folderId : folderIds) {
      registeredCalendarIds.add(CalendarConverterService.getCalendarId(folderId));
    }

    List<Calendar> calendars = storage.getUserCalendars(username, true);
    Iterator<Calendar> iterator = calendars.iterator();
    while (iterator.hasNext()) {
      Calendar calendar = (Calendar) iterator.next();
      String calendarId = calendar.getId();
      if (CalendarConverterService.isExchangeCalendarId(calendarId) && !registeredCalendarIds.contains(calendarId)) {
        LOG.info("Remove user calendar because it was deleted from Exchange: " + calendar.getName());
        storage.removeUserCalendar(username, calendarId);
      }
    }
  }

  private CalendarEvent createOrUpdateEvent(Appointment appointment, Folder folder, String username, boolean isNew, TimeZone timeZone) throws Exception {
    Calendar calendar = getOrCreateUserCalendar(username, folder);
    CalendarEvent event = null;

    if (isNew) {
      LOG.info("Create user calendar event: " + appointment.getSubject());
    } else {
      LOG.info("Update user calendar event: " + appointment.getSubject());
    }

    if (appointment.getAppointmentType() != null) {
      switch (appointment.getAppointmentType()) {
      case Single: {
        event = null;
        if (isNew) {
          event = new CalendarEvent();
          event.setCalendarId(calendar.getId());
        } else {
          event = getEventByAppointmentId(username, appointment.getId().getUniqueId());
        }
        CalendarConverterService.convertExchangeToExoEvent(event, appointment, username, storage, organizationService.getUserHandler(), timeZone);
        event.setRepeatType(CalendarEvent.RP_NOREPEAT);
        storage.saveUserEvent(username, calendar.getId(), event, isNew);
        correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
      }
        break;
      case Exception:
        throw new IllegalStateException("The appointment is an exception occurence of this event >> '" + appointment.getSubject() + "'. start:" + appointment.getStart() + ", end : "
            + appointment.getEnd() + ", occurence: " + appointment.getAppointmentSequenceNumber());
      case RecurringMaster: {
        // Master recurring event
        CalendarEvent masterEvent = null;
        Date orginialStartDate = null;
        if (isNew) {
          masterEvent = new CalendarEvent();
        } else {
          masterEvent = getEventByAppointmentId(username, appointment.getId().getUniqueId());
          orginialStartDate = masterEvent.getFromDateTime();
        }
        masterEvent.setCalendarId(calendar.getId());
        CalendarConverterService.convertExchangeToExoMasterRecurringCalendarEvent(masterEvent, appointment, username, storage, organizationService.getUserHandler(), timeZone);
        if (isNew) {
          correspondenceService.setCorrespondingId(username, masterEvent.getId(), appointment.getId().getUniqueId());
        } else if (!CalendarConverterService.isSameDate(orginialStartDate, masterEvent.getFromDateTime())) {
          masterEvent.setExcludeId(new String[0]);
        }
        storage.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
        List<CalendarEvent> exceptionalEventsToUpdate = new ArrayList<CalendarEvent>();
        List<String> occAppointmentIDs = new ArrayList<String>();
        // Deleted execptional occurences events.
        List<CalendarEvent> toDeleteEvents = CalendarConverterService.convertExchangeToExoOccurenceEvent(masterEvent, exceptionalEventsToUpdate, occAppointmentIDs, appointment, username, storage,
            organizationService.getUserHandler(), timeZone);
        if (exceptionalEventsToUpdate != null && !exceptionalEventsToUpdate.isEmpty()) {
          storage.updateOccurrenceEvent(calendar.getId(), calendar.getId(), masterEvent.getCalType(), masterEvent.getCalType(), exceptionalEventsToUpdate, username);

          // Set correspondance IDs
          Iterator<CalendarEvent> eventsIterator = exceptionalEventsToUpdate.iterator();
          Iterator<String> occAppointmentIdIterator = occAppointmentIDs.iterator();
          while (eventsIterator.hasNext()) {
            CalendarEvent calendarEvent = eventsIterator.next();
            String occAppointmentId = occAppointmentIdIterator.next();
            correspondenceService.setCorrespondingId(username, calendarEvent.getId(), occAppointmentId);
          }
        }
        if (toDeleteEvents != null && !toDeleteEvents.isEmpty()) {
          for (CalendarEvent calendarEvent : toDeleteEvents) {
            storage.removeUserEvent(username, calendar.getId(), calendarEvent.getId());
            // Only if dates was modified
            if (!isNew && !CalendarConverterService.isSameDate(orginialStartDate, masterEvent.getFromDateTime()) && masterEvent.getExcludeId() != null && masterEvent.getExcludeId().length > 0) {
              String[] occIds = masterEvent.getExcludeId();
              List<String> newOccIds = new ArrayList<String>();
              for (String occId : occIds) {
                if (occId != null && !occId.equals(calendarEvent.getRecurrenceId())) {
                  newOccIds.add(occId);
                }
              }
              if (newOccIds.size() < occIds.length) {
                storage.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
              }
            }
          }
        }
        event = masterEvent;
      }
        break;
      case Occurrence:
        LOG.warn("The appointment is an occurence of this event >> '" + appointment.getSubject() + "'. start:" + appointment.getStart() + ", end : " + appointment.getEnd() + ", occurence: "
            + appointment.getAppointmentSequenceNumber());
      }
    }
    return event;
  }

  public CalendarEvent getEventByAppointmentId(String username, String appointmentId) throws Exception {
    String calEventId = correspondenceService.getCorrespondingId(username, appointmentId);
    CalendarEvent event = storage.getEvent(username, calEventId);
    if (event == null && calEventId != null) {
      correspondenceService.deleteCorrespondingId(username, appointmentId);
    }
    return event;
  }

  public List<CalendarEvent> searchAllEvents(String username, Calendar calendar) throws Exception {
    List<String> calendarIds = Collections.singletonList(calendar.getId());
    return storage.getUserEventByCalendar(username, calendarIds);
  }

  public List<CalendarEvent> searchEventsModifiedSince(String username, Calendar calendar, Date date) throws Exception {
    List<CalendarEvent> resultEvents = new ArrayList<CalendarEvent>();
    List<CalendarEvent> calendarEvents = searchAllEvents(username, calendar);
    for (CalendarEvent calendarEvent : calendarEvents) {
      if (calendarEvent.getLastUpdatedTime() == null) {
        continue;
      }
      if (calendarEvent.getLastUpdatedTime().after(date)) {
        resultEvents.add(calendarEvent);
      }
    }
    return resultEvents;
  }

}
