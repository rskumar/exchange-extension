package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.ArrayList;
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

  public void deleteEvent(String appointmentId, String folderId, String username) throws Exception {
    CalendarEvent calendarEvent = getEventByAppointmentId(username, appointmentId);
    if (calendarEvent != null) {
      deleteEvent(calendarEvent, folderId, username);
    }
  }

  public void deleteEvent(CalendarEvent calendarEvent, String folderId, String username) throws Exception {
    Calendar calendar = findUserCalendar(username, folderId);
    calendarService.removeUserEvent(username, calendar.getId(), calendarEvent.getId());
  }

  public void createOrUpdateEvent(Appointment appointment, Folder folder, String username) throws Exception {
    CalendarEvent calendarEvent = getEventByAppointmentId(username, appointment.getId().getUniqueId());
    boolean isNew = (calendarEvent == null);
    createOrUpdateEvent(appointment, folder, username, isNew);
  }

  protected void updateEvent(Appointment appointment, Folder forlder, String username) throws Exception {
    createOrUpdateEvent(appointment, forlder, username, false);
  }

  protected void createEvent(Appointment appointment, Folder forlder, String username) throws Exception {
    createOrUpdateEvent(appointment, forlder, username, true);
  }

  public Calendar getOrCreateUserCalendar(String username, Folder forlder) throws Exception {
    Calendar calendar = findUserCalendar(username, forlder.getId().getUniqueId());
    if (calendar == null) {
      calendar = new Calendar();
      calendar.setId(CalendarConverterService.getCalendarId(username, forlder.getId().getUniqueId()));
      calendar.setName(CalendarConverterService.getCalendarName(forlder.getDisplayName()));
      calendar.setCalendarOwner(username);
      calendar.setDataInit(false);
      calendar.setEditPermission(new String[] { "any read" });
      calendar.setCalendarColor(Calendar.COLORS[9]);
      calendarService.saveUserCalendar(username, calendar, true);
    }
    return calendar;
  }

  public Calendar findUserCalendar(String username, String folderId) throws Exception {
    return calendarService.getUserCalendar(username, CalendarConverterService.getCalendarId(username, folderId));
  }

  public List<CalendarEvent> findUserCalendarEvents(String username, String folderId) throws Exception {
    List<CalendarEvent> userEvents = null;
    String calendarId = CalendarConverterService.getCalendarId(username, folderId);
    Calendar calendar = calendarService.getUserCalendar(username, CalendarConverterService.getCalendarId(username, folderId));
    if (calendar != null) {
      List<String> calendarIds = new ArrayList<String>();
      calendarIds.add(calendarId);
      userEvents = calendarService.getUserEventByCalendar(username, calendarIds);
    }
    return userEvents;
  }

  private void createOrUpdateEvent(Appointment appointment, Folder forlder, String username, boolean isNew) throws Exception {
    Calendar calendar = getOrCreateUserCalendar(username, forlder);

    if (appointment.getAppointmentType() != null) {
      switch (appointment.getAppointmentType()) {
      case Single: {
        CalendarEvent event = new CalendarEvent();
        event.setCalendarId(calendar.getId());
        event.setId(CalendarConverterService.getEventId(username, appointment.getId().getUniqueId()));
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
        CalendarEvent oldMasterEvent = null, masterEvent = null;
        if (isNew) {
          masterEvent = new CalendarEvent();
        } else {
          masterEvent = getEventByAppointmentId(username, appointment.getId().getUniqueId());
          oldMasterEvent = new CalendarEvent(masterEvent);
        }
        masterEvent.setCalendarId(calendar.getId());
        CalendarConverterService.convertMasterRecurringCalendarEvent(masterEvent, appointment, username, calendarService);
        if (isNew) {
          calendarService.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
        } else {
          if (!isNew && !CalendarConverterService.isSameDate(oldMasterEvent.getFromDateTime(), masterEvent.getFromDateTime())) {
            masterEvent.setExcludeId(new String[0]);
          }
          calendarService.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
        }
        List<CalendarEvent> listEvent = new ArrayList<CalendarEvent>();
        List<CalendarEvent> toDeleteEvents = CalendarConverterService.convertExceptionOccurencesOfRecurringEvent(masterEvent, listEvent, appointment, username, calendarService);
        if (listEvent != null && !listEvent.isEmpty()) {
          calendarService.updateOccurrenceEvent(calendar.getId(), calendar.getId(), masterEvent.getCalType(), masterEvent.getCalType(), listEvent, username);
        }
        if (toDeleteEvents != null && !toDeleteEvents.isEmpty()) {
          for (CalendarEvent calendarEvent : toDeleteEvents) {
            calendarService.removeUserEvent(username, calendar.getId(), calendarEvent.getId());
            // Only if dates was modified
            if (!isNew && !CalendarConverterService.isSameDate(oldMasterEvent.getFromDateTime(), masterEvent.getFromDateTime()) && masterEvent.getExcludeId() != null
                && masterEvent.getExcludeId().length > 0) {
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
    String calEventId = CalendarConverterService.getEventId(username, appointmentId);
    return calendarService.getEvent(username, calEventId);
  }

}
