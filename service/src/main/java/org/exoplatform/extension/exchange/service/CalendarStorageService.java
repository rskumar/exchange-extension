package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.Folder;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.extension.exchange.service.util.CalendarConverterService;
import org.mortbay.log.Log;

public class CalendarStorageService implements Serializable {
  private static final long serialVersionUID = 6614108102985034995L;

  private CalendarService calendarService;

  public CalendarStorageService(CalendarService calendarService) {
    this.calendarService = calendarService;
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
  protected void deleteEvent(String appointmentId, String folderId, String username) throws Exception {
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
    calendarService.removeUserEvent(username, calendar.getId(), calendarEvent.getId());
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
  protected void createOrUpdateEvent(Appointment appointment, Folder folder, String username) throws Exception {
    CalendarEvent calendarEvent = getEventByAppointmentId(username, appointment.getId().getUniqueId());
    boolean isNew = (calendarEvent == null);
    createOrUpdateEvent(appointment, folder, username, isNew);
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
    Calendar calendar = calendarService.removeUserCalendar(username, CalendarConverterService.getCalendarId(folderId));
    if (calendar != null) {
      Log.info("User Calendar" + calendar.getId() + " is deleted, because it was deleted from Exchange.");
    }
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
    return calendarService.getUserCalendar(username, CalendarConverterService.getCalendarId(folderId));
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
      calendar = new Calendar();
      calendar.setId(CalendarConverterService.getCalendarId(folder.getId().getUniqueId()));
      calendar.setName(CalendarConverterService.getCalendarName(folder.getDisplayName()));
      calendar.setCalendarOwner(username);
      calendar.setDataInit(false);
      calendar.setEditPermission(new String[] { "any read" });
      calendar.setCalendarColor(Calendar.COLORS[9]);
      calendarService.saveUserCalendar(username, calendar, true);
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
    Calendar calendar = calendarService.getUserCalendar(username, CalendarConverterService.getCalendarId(folderId));
    if (calendar != null) {
      List<String> calendarIds = new ArrayList<String>();
      calendarIds.add(calendarId);
      userEvents = calendarService.getUserEventByCalendar(username, calendarIds);
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
   * @throws Exception
   */
  protected void updateEvent(Appointment appointment, Folder folder, String username) throws Exception {
    createOrUpdateEvent(appointment, folder, username, false);
  }

  /**
   * 
   * Create non existing eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @throws Exception
   */
  protected void createEvent(Appointment appointment, Folder folder, String username) throws Exception {
    createOrUpdateEvent(appointment, folder, username, true);
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

    List<Calendar> calendars = calendarService.getUserCalendars(username, true);
    Iterator<Calendar> iterator = calendars.iterator();
    while (iterator.hasNext()) {
      Calendar calendar = (Calendar) iterator.next();
      String calendarId = calendar.getId();
      if (CalendarConverterService.isExchangeCalendarId(calendarId) && !registeredCalendarIds.contains(calendarId)) {
        calendarService.removeUserCalendar(username, calendarId);
        Log.info("User Calendar '" + calendarId + "' is deleted, because it was deleted from Exchange.");
      }
    }
  }

  private void createOrUpdateEvent(Appointment appointment, Folder folder, String username, boolean isNew) throws Exception {
    Calendar calendar = getOrCreateUserCalendar(username, folder);

    if (appointment.getAppointmentType() != null) {
      switch (appointment.getAppointmentType()) {
      case Single: {
        CalendarEvent event = new CalendarEvent();
        event.setCalendarId(calendar.getId());
        event.setId(CalendarConverterService.getEventId(appointment.getId().getUniqueId()));
        CalendarConverterService.convertSingleCalendarEvent(event, appointment, username, calendarService);
        event.setRepeatType(CalendarEvent.RP_NOREPEAT);
        calendarService.saveUserEvent(username, calendar.getId(), event, isNew);
      }
        break;
      case Exception:
        Log.warn("The appointment is an exception occurence of this event >> '" + appointment.getSubject() + "'. start:" + appointment.getStart() + ", end : " + appointment.getEnd() + ", occurence: "
            + appointment.getAppointmentSequenceNumber());
        break;
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
        CalendarConverterService.convertMasterRecurringCalendarEvent(masterEvent, appointment, username, calendarService);
        if (isNew) {
          calendarService.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
        } else {
          if (!isNew && !CalendarConverterService.isSameDate(orginialStartDate, masterEvent.getFromDateTime())) {
            masterEvent.setExcludeId(new String[0]);
          }
          calendarService.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
        }
        List<CalendarEvent> exceptionalEventsToUpdate = new ArrayList<CalendarEvent>();
        // Deleted execptional occurences events.
        List<CalendarEvent> toDeleteEvents = CalendarConverterService.convertExceptionOccurencesOfRecurringEvent(masterEvent, exceptionalEventsToUpdate, appointment, username, calendarService);
        if (exceptionalEventsToUpdate != null && !exceptionalEventsToUpdate.isEmpty()) {
          calendarService.updateOccurrenceEvent(calendar.getId(), calendar.getId(), masterEvent.getCalType(), masterEvent.getCalType(), exceptionalEventsToUpdate, username);
        }
        if (toDeleteEvents != null && !toDeleteEvents.isEmpty()) {
          for (CalendarEvent calendarEvent : toDeleteEvents) {
            calendarService.removeUserEvent(username, calendar.getId(), calendarEvent.getId());
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
                calendarService.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
              }
            }
          }
        }
      }
        break;
      case Occurrence:
        Log.warn("The appointment is an occurence of this event >> '" + appointment.getSubject() + "'. start:" + appointment.getStart() + ", end : " + appointment.getEnd() + ", occurence: "
            + appointment.getAppointmentSequenceNumber());
      }
    }
  }

  private CalendarEvent getEventByAppointmentId(String username, String appointmentId) throws Exception {
    String calEventId = CalendarConverterService.getEventId(appointmentId);
    return calendarService.getEvent(username, calEventId);
  }

}
