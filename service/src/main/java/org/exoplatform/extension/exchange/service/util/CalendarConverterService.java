package org.exoplatform.extension.exchange.service.util;

import java.io.ByteArrayInputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.AppointmentSchema;
import microsoft.exchange.webservices.data.Attachment;
import microsoft.exchange.webservices.data.Attendee;
import microsoft.exchange.webservices.data.BasePropertySet;
import microsoft.exchange.webservices.data.FileAttachment;
import microsoft.exchange.webservices.data.OccurrenceInfo;
import microsoft.exchange.webservices.data.OccurrenceInfoCollection;
import microsoft.exchange.webservices.data.PropertySet;
import microsoft.exchange.webservices.data.Recurrence;
import microsoft.exchange.webservices.data.Recurrence.DailyPattern;
import microsoft.exchange.webservices.data.Recurrence.IntervalPattern;
import microsoft.exchange.webservices.data.Recurrence.MonthlyPattern;
import microsoft.exchange.webservices.data.Recurrence.WeeklyPattern;
import microsoft.exchange.webservices.data.Recurrence.YearlyPattern;
import microsoft.exchange.webservices.data.Sensitivity;
import microsoft.exchange.webservices.data.ServiceLocalException;
import microsoft.exchange.webservices.data.ServiceVersionException;

import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.calendar.service.EventCategory;
import org.exoplatform.extension.exchange.listener.ExchangeListenerService;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;

public class CalendarConverterService {

  private final static Log LOG = ExoLogger.getLogger(ExchangeListenerService.class);

  public static final String EXCHANGE_CALENDAR_NAME = "EXCH";

  public static final SimpleDateFormat recurrenceIdFormat = new SimpleDateFormat("yyyyMMdd'T'HHmmss'Z'");

  /**
   * 
   * @param event
   * @param appointment
   * @param username
   * @param calendarService
   * @throws Exception
   */
  public static void convertSingleCalendarEvent(CalendarEvent event, Appointment appointment, String username, CalendarService calendarService) throws Exception {
    event.setId(getEventId(username, appointment.getId().getUniqueId()));
    event.setMessage(appointment.getId().getUniqueId());
    event.setEventType(CalendarEvent.TYPE_EVENT);
    event.setCalType("" + org.exoplatform.calendar.service.Calendar.TYPE_PRIVATE);
    event.setLocation(appointment.getLocation());
    event.setLastUpdatedTime(appointment.getLastModifiedTime());
    event.setSummary(appointment.getSubject());
    setStatus(event, appointment);
    setAttachements(event, appointment);
    setDates(event, appointment);
    setPriority(event, appointment);
    setEventCategory(event, appointment, username, calendarService);
    setParticipants(event, appointment);
    if (appointment.getSensitivity() != null && !appointment.getSensitivity().equals(Sensitivity.Normal)) {
      event.setPrivate(true);
    } else {
      event.setPrivate(false);
    }
    // This have to be last thing to load because of BAD EWS API impl
    appointment.load(new PropertySet(AppointmentSchema.Body));
    event.setDescription(appointment.getBody().toString());
  }

  public static void convertMasterRecurringCalendarEvent(CalendarEvent event, Appointment appointment, String username, CalendarService calendarService) throws Exception {
    convertSingleCalendarEvent(event, appointment, username, calendarService);
    appointment = Appointment.bind(appointment.getService(), appointment.getId(), new PropertySet(AppointmentSchema.Recurrence));
    Recurrence recurrence = appointment.getRecurrence();
    if (recurrence instanceof DailyPattern) {
      event.setRepeatType(CalendarEvent.RP_DAILY);
    } else if (recurrence instanceof WeeklyPattern) {
      event.setRepeatType(CalendarEvent.RP_WEEKEND);
    } else if (recurrence instanceof MonthlyPattern) {
      event.setRepeatType(CalendarEvent.RP_MONTHLY);
    } else if (recurrence instanceof YearlyPattern) {
      event.setRepeatType(CalendarEvent.RP_YEARLY);
    }
    if (recurrence instanceof IntervalPattern) {
      if (((IntervalPattern) recurrence).getInterval() > 0) {
        event.setRepeatInterval(((IntervalPattern) recurrence).getInterval());
      }
    }
    if (recurrence.hasEnd()) {
      event.setRepeatUntilDate(recurrence.getEndDate());
    }
    if (recurrence.getNumberOfOccurrences() != null) {
      event.setRepeatCount(recurrence.getNumberOfOccurrences());
    }
  }

  public static List<CalendarEvent> convertExceptionOccurencesOfRecurringEvent(CalendarEvent masterEvent, List<CalendarEvent> listEvent, Appointment masterAppointment, String username,
      CalendarService calendarService) throws Exception {
    masterAppointment = Appointment.bind(masterAppointment.getService(), masterAppointment.getId(), new PropertySet(AppointmentSchema.ModifiedOccurrences));
    List<CalendarEvent> calendarEvents = calendarService.getExceptionEvents(username, masterEvent);
    OccurrenceInfoCollection occurrenceInfoCollection = masterAppointment.getModifiedOccurrences();
    if (occurrenceInfoCollection != null && occurrenceInfoCollection.getCount() > 0) {
      for (OccurrenceInfo occurrenceInfo : occurrenceInfoCollection) {
        Appointment occureceAppointment = Appointment.bind(masterAppointment.getService(), occurrenceInfo.getItemId(), new PropertySet(BasePropertySet.FirstClassProperties));
        CalendarEvent tmpEvent = new CalendarEvent();
        convertSingleCalendarEvent(tmpEvent, occureceAppointment, username, calendarService);
        tmpEvent.setCalendarId(masterEvent.getCalendarId());
        tmpEvent.setRepeatType(CalendarEvent.RP_NOREPEAT);
        tmpEvent.setId(masterEvent.getId());
        tmpEvent.setRecurrenceId(recurrenceIdFormat.format(tmpEvent.getFromDateTime()));
        tmpEvent.setMessage(occurrenceInfo.getItemId().getUniqueId());
        try {
          setOldEventId(masterEvent, calendarService, tmpEvent, calendarEvents == null ? null : calendarEvents.iterator());
        } catch (IllegalStateException e) {
          LOG.error(e);
          return new ArrayList<CalendarEvent>();
        }

        listEvent.add(tmpEvent);
      }
    }
    return calendarEvents;
  }

  public static String getCalendarName(String calendarName) {
    return EXCHANGE_CALENDAR_NAME + "-" + calendarName;
  }

  public static String getCalendarId(String username, String folderId) {
    return EXCHANGE_CALENDAR_NAME + "-" + folderId.hashCode();
  }

  public static String getEventId(String username, String appointmentId) throws Exception {
    return "ExcangeEvent-" + appointmentId.hashCode();
  }

  public static boolean isSameDate(Date value1, Date value2) {
    Calendar date1 = Calendar.getInstance();
    date1.setTime(value1);
    Calendar date2 = Calendar.getInstance();
    date2.setTime(value2);
    return isSameDate(date1, date2);
  }

  public static boolean isSameDate(java.util.Calendar date1, java.util.Calendar date2) {
    return (date1.get(java.util.Calendar.DATE) == date2.get(java.util.Calendar.DATE) && date1.get(java.util.Calendar.MONTH) == date2.get(java.util.Calendar.MONTH) && date1
        .get(java.util.Calendar.YEAR) == date2.get(java.util.Calendar.YEAR));
  }

  private static void setOldEventId(CalendarEvent masterEvent, CalendarService calendarService, CalendarEvent tmpEvent, Iterator<CalendarEvent> calendarEventIterator) throws Exception {
    if (calendarEventIterator != null && calendarEventIterator.hasNext()) {
      CalendarEvent originalOccEvent = null;
      while (calendarEventIterator.hasNext() && originalOccEvent == null) {
        CalendarEvent calendarEvent = calendarEventIterator.next();
        if (calendarEvent.getRecurrenceId() != null && calendarEvent.getRecurrenceId().equals(tmpEvent.getRecurrenceId())) {
          originalOccEvent = calendarEvent;
          calendarEventIterator.remove();
        }
      }
      tmpEvent.setId(originalOccEvent.getId());
      tmpEvent.setOriginalReference(originalOccEvent.getOriginalReference());
      tmpEvent.setIsExceptionOccurrence(true);
      tmpEvent.setRepeatInterval(0);
      tmpEvent.setRepeatCount(0);
      tmpEvent.setRepeatUntilDate(null);
      tmpEvent.setRepeatByDay(null);
      tmpEvent.setRepeatByMonthDay(null);
    }
  }

  /**
   * 
   * @param calendarEvent
   * @param appointment
   * @throws ServiceLocalException
   */
  private static void setParticipants(CalendarEvent calendarEvent, Appointment appointment) throws ServiceLocalException {
    List<String> participants = new ArrayList<String>();
    if (appointment.getOptionalAttendees() != null) {
      for (Attendee attendee : appointment.getRequiredAttendees()) {
        if (attendee.getName() != null) {
          participants.add(attendee.getName());
        }
      }
    }
    if (appointment.getOptionalAttendees() != null) {
      for (Attendee attendee : appointment.getOptionalAttendees()) {
        if (attendee.getName() != null) {
          participants.add(attendee.getName());
        }
      }
    }
    if (appointment.getResources() != null) {
      for (Attendee attendee : appointment.getResources()) {
        if (attendee.getName() != null) {
          participants.add(attendee.getName());
        }
      }
    }
    if (participants.size() > 0) {
      calendarEvent.setParticipant(participants.toArray(new String[0]));
    }
  }

  /**
   * 
   * @param calendarEvent
   * @param appointment
   * @throws ServiceLocalException
   */
  private static void setPriority(CalendarEvent calendarEvent, Appointment appointment) throws ServiceLocalException {
    if (appointment.getImportance() != null) {
      // Transform index 1,2,3 => 3,2,1. See CalendarEvent.PRIORITY and
      // Importance enum.
      int priority = 4 - appointment.getImportance().ordinal();
      calendarEvent.setPriority("" + priority);
    } else {
      calendarEvent.setPriority("0");
    }
  }

  /**
   * 
   * @param calendarEvent
   * @param appointment
   * @throws ServiceLocalException
   */
  private static void setDates(CalendarEvent calendarEvent, Appointment appointment) throws ServiceLocalException {
    calendarEvent.setFromDateTime(appointment.getStart());
    calendarEvent.setToDateTime(appointment.getEnd());
    if (appointment.getIsAllDayEvent()) {
      Calendar cal1 = Calendar.getInstance(), cal2 = Calendar.getInstance();
      cal1.setTime(appointment.getStart());
      if (cal1.get(Calendar.HOUR_OF_DAY) >= 22) {
        cal1.add(Calendar.DATE, 1);
      }
      cal1.set(Calendar.HOUR_OF_DAY, 0);
      cal1.set(Calendar.MINUTE, 0);

      cal2.setTime(appointment.getEnd());
      cal2.set(Calendar.HOUR_OF_DAY, cal2.getActualMaximum(Calendar.HOUR_OF_DAY));
      cal2.set(Calendar.MINUTE, cal2.getActualMaximum(Calendar.MINUTE));

      calendarEvent.setFromDateTime(cal1.getTime());
      calendarEvent.setToDateTime(cal2.getTime());
    }
  }

  /**
   * 
   * @param calendarEvent
   * @param appointment
   * @param username
   * @param calendarService
   * @throws ServiceLocalException
   * @throws Exception
   */
  private static void setEventCategory(CalendarEvent calendarEvent, Appointment appointment, String username, CalendarService calendarService) throws ServiceLocalException, Exception {
    if (appointment.getCategories() != null && appointment.getCategories().getSize() > 0) {
      String categoryName = appointment.getCategories().getString(0);
      if (categoryName != null && !categoryName.isEmpty()) {
        EventCategory category = calendarService.getEventCategoryByName(username, categoryName);
        if (category == null) {
          category = new EventCategory();
          category.setDataInit(false);
          category.setName(categoryName);
          category.setId(EXCHANGE_CALENDAR_NAME + "-" + categoryName);
          calendarService.saveEventCategory(username, category, true);
        }
        calendarEvent.setEventCategoryId(category.getId());
        calendarEvent.setEventCategoryName(category.getName());
      }
    }
  }

  /**
   * 
   * @param calendarEvent
   * @param appointment
   * @throws ServiceLocalException
   * @throws ServiceVersionException
   * @throws Exception
   */
  private static void setAttachements(CalendarEvent calendarEvent, Appointment appointment) throws ServiceLocalException, ServiceVersionException, Exception {
    if (appointment.getHasAttachments()) {
      Iterator<Attachment> attachmentIterator = appointment.getAttachments().iterator();
      List<org.exoplatform.calendar.service.Attachment> attachments = new ArrayList<org.exoplatform.calendar.service.Attachment>();
      while (attachmentIterator.hasNext()) {
        Attachment attachment = attachmentIterator.next();
        if (attachment instanceof FileAttachment) {
          FileAttachment fileAttachment = (FileAttachment) attachment;
          org.exoplatform.calendar.service.Attachment eXoAttachment = new org.exoplatform.calendar.service.Attachment();
          if (fileAttachment.getSize() == 0) {
            continue;
          }
          eXoAttachment.setInputStream(new ByteArrayInputStream(fileAttachment.getContent()));
          eXoAttachment.setMimeType(fileAttachment.getContentType());
          eXoAttachment.setName(fileAttachment.getFileName());
          eXoAttachment.setSize(fileAttachment.getSize());
          Calendar calendar = Calendar.getInstance();
          calendar.setTime(fileAttachment.getLastModifiedTime());
          eXoAttachment.setLastModified(calendar);
          attachments.add(eXoAttachment);
        }
      }
      calendarEvent.setAttachment(attachments);
    }
  }

  /**
   * 
   * @param calendarEvent
   * @param appointment
   * @throws ServiceLocalException
   */
  private static void setStatus(CalendarEvent calendarEvent, Appointment appointment) throws ServiceLocalException {
    if (appointment.getAppointmentState() != null) {
      switch (appointment.getAppointmentState()) {
      case 1:
        calendarEvent.setStatus(CalendarEvent.ST_AVAILABLE);
        calendarEvent.setEventState(CalendarEvent.ST_AVAILABLE);
        break;
      case 3:
        calendarEvent.setStatus(CalendarEvent.ST_BUSY);
        calendarEvent.setEventState(CalendarEvent.ST_BUSY);
        break;
      case 4:
        calendarEvent.setStatus(CalendarEvent.ST_OUTSIDE);
        calendarEvent.setEventState(CalendarEvent.ST_OUTSIDE);
        break;
      }
    }
  }
}