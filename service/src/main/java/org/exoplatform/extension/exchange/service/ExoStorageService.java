package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;
import java.util.List;
import java.util.TimeZone;

import javax.jcr.Node;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.Folder;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.calendar.service.impl.CalendarServiceImpl;
import org.exoplatform.calendar.service.impl.JCRDataStorage;
import org.exoplatform.extension.exchange.service.util.CalendarConverterService;
import org.exoplatform.services.jcr.ext.common.SessionProvider;
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
  protected void deleteEventByAppointmentID(String appointmentId, String username) throws Exception {
    CalendarEvent calendarEvent = getEventByAppointmentId(username, appointmentId);
    if (calendarEvent != null) {
      deleteEvent(username, calendarEvent);
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
  protected void deleteEvent(String username, CalendarEvent calendarEvent) throws Exception {
    if (calendarEvent == null) {
      LOG.warn("Event is null, can't delete it for username: " + username);
      return;
    }

    if ((calendarEvent.getRepeatType() == null || calendarEvent.getRepeatType().equals(CalendarEvent.RP_NOREPEAT))
        && (calendarEvent.getIsExceptionOccurrence() == null || !calendarEvent.getIsExceptionOccurrence())) {
      LOG.info("Delete user calendar event: " + calendarEvent.getSummary());
      storage.removeUserEvent(username, calendarEvent.getCalendarId(), calendarEvent.getId());
      // Remove correspondence between exo and exchange IDs
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    } else if ((calendarEvent.getRecurrenceId() != null && !calendarEvent.getRecurrenceId().isEmpty())
        || (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence())) {
      LOG.info("Delete user calendar event occurence: " + calendarEvent.getSummary() + ", id=" + calendarEvent.getRecurrenceId());
      storage.removeOccurrenceInstance(username, calendarEvent);
      if (storage.getEvent(username, calendarEvent.getId()) != null) {
        storage.removeUserEvent(username, calendarEvent.getCalendarId(), calendarEvent.getId());
      }
      if (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence()) {
        // Remove correspondence between exo and exchange IDs
        correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
      }
    } else {
      LOG.info("Delete user calendar event series: " + calendarEvent.getSummary());
      storage.removeRecurrenceSeries(username, calendarEvent);
      // Remove correspondence between exo and exchange IDs
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    }

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
  protected List<CalendarEvent> createOrUpdateEvent(Appointment appointment, Folder folder, String username, TimeZone timeZone) throws Exception {
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
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    if (calendarId == null) {
      calendarId = CalendarConverterService.getCalendarId(folderId);
    }
    Calendar calendar = storage.removeUserCalendar(username, calendarId);
    if (calendar != null) {
      LOG.info("User Calendar" + calendar.getName() + " is deleted, because it was deleted from Exchange.");
    }
    correspondenceService.deleteCorrespondingId(username, folderId, calendarId);
    return true;
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
    return getUserCalendar(username, folderId, true);
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
  protected Calendar getUserCalendar(String username, String folderId, boolean deleteIfCorrespondentExists) throws Exception {
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    Calendar calendar = null;
    if (calendarId != null) {
      calendar = storage.getUserCalendar(username, calendarId);
      if (calendar == null && deleteIfCorrespondentExists) {
        correspondenceService.deleteCorrespondingId(username, folderId);
      }
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
    Calendar calendar = getUserCalendar(username, folder.getId().getUniqueId(), false);
    String calendarId = CalendarConverterService.getCalendarId(folder.getId().getUniqueId());
    if (calendar == null) {
      Calendar tmpCalendar = storage.getUserCalendar(username, calendarId);
      if (tmpCalendar != null) {
        // Renew Calendar
        storage.removeUserCalendar(username, calendarId);
      }
    }
    if (calendar == null) {
      LOG.info("Create user calendar from Exchange: " + folder.getDisplayName());

      calendar = new Calendar();
      calendar.setId(calendarId);
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
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    if (calendarId == null) {
      calendarId = CalendarConverterService.getCalendarId(folderId);
    }
    Calendar calendar = storage.getUserCalendar(username, calendarId);
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
  protected List<CalendarEvent> updateEvent(Appointment appointment, Folder folder, String username, TimeZone timeZone) throws Exception {
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
  protected List<CalendarEvent> createEvent(Appointment appointment, Folder folder, String username, TimeZone timeZone) throws Exception {
    return createOrUpdateEvent(appointment, folder, username, true, timeZone);
  }

  private List<CalendarEvent> createOrUpdateEvent(Appointment appointment, Folder folder, String username, boolean isNew, TimeZone timeZone) throws Exception {
    Calendar calendar = getUserCalendar(username, folder.getId().getUniqueId());
    if (calendar == null) {
      LOG.warn("Attempting to synchronize an event without existing associated eXo Calendar.");
      return null;
    }
    List<CalendarEvent> updatedEvents = new ArrayList<CalendarEvent>();

    if (appointment.getAppointmentType() != null) {
      switch (appointment.getAppointmentType()) {
      case Single: {
        CalendarEvent event = null;
        if (isNew) {
          event = new CalendarEvent();
          event.setCalendarId(calendar.getId());
          updatedEvents.add(event);
        } else {
          event = getEventByAppointmentId(username, appointment.getId().getUniqueId());
          updatedEvents.add(event);
          if (CalendarConverterService.verifyModifiedDatesConflict(event, appointment)) {
            if (LOG.isTraceEnabled()) {
              LOG.trace("Attempting to update eXo Event with Exchange Event, but modification date of eXo is after, ignore updating.");
            }
            return updatedEvents;
          }
        }

        if (isNew) {
          LOG.info("Create user calendar event: " + appointment.getSubject());
        } else {
          LOG.info("Update user calendar event: " + appointment.getSubject());
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
          updatedEvents.add(masterEvent);
        } else {
          masterEvent = getEventByAppointmentId(username, appointment.getId().getUniqueId());
          updatedEvents.add(masterEvent);
          orginialStartDate = masterEvent.getFromDateTime();
        }

        if (!isNew && CalendarConverterService.verifyModifiedDatesConflict(masterEvent, appointment)) {
          if (LOG.isTraceEnabled()) {
            LOG.trace("Attempting to update eXo Event with Exchange Event, but modification date of eXo is after, ignore updating.");
          }
        } else {
          if (isNew) {
            LOG.info("Create recurrent user calendar event: " + appointment.getSubject());
          } else {
            LOG.info("Update recurrent user calendar event: " + appointment.getSubject());
          }

          masterEvent.setCalendarId(calendar.getId());
          CalendarConverterService.convertExchangeToExoMasterRecurringCalendarEvent(masterEvent, appointment, username, storage, organizationService.getUserHandler(), timeZone);
          if (isNew) {
            correspondenceService.setCorrespondingId(username, masterEvent.getId(), appointment.getId().getUniqueId());
          } else if (!CalendarConverterService.isSameDate(orginialStartDate, masterEvent.getFromDateTime())) {
            if (masterEvent.getExcludeId() == null) {
              masterEvent.setExcludeId(new String[0]);
            }
          }
          storage.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
        }
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
          updatedEvents.addAll(exceptionalEventsToUpdate);
        }
        if (toDeleteEvents != null && !toDeleteEvents.isEmpty()) {
          for (CalendarEvent calendarEvent : toDeleteEvents) {
            deleteEvent(username, calendarEvent);
          }
        }
      }
        break;
      case Occurrence:
        LOG.warn("The appointment is an occurence of this event >> '" + appointment.getSubject() + "'. start:" + appointment.getStart() + ", end : " + appointment.getEnd() + ", occurence: "
            + appointment.getAppointmentSequenceNumber());
      }
    }
    return updatedEvents;
  }

  /**
   * 
   * @param username
   * @param appointmentId
   * @return
   * @throws Exception
   */
  public CalendarEvent getEventByAppointmentId(String username, String appointmentId) throws Exception {
    String calEventId = correspondenceService.getCorrespondingId(username, appointmentId);
    CalendarEvent event = storage.getEvent(username, calEventId);
    if (event == null && calEventId != null) {
      correspondenceService.deleteCorrespondingId(username, appointmentId);
    }
    return event;
  }

  /**
   * 
   * @param username
   * @param calendar
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> searchAllEvents(String username, Calendar calendar) throws Exception {
    List<String> calendarIds = Collections.singletonList(calendar.getId());
    return storage.getUserEventByCalendar(username, calendarIds);
  }

  /**
   * 
   * @param username
   * @param calendar
   * @param date
   * @return
   * @throws Exception
   */
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

  public String getRecurrentEventIdByOriginalUUID(String uuid) throws Exception {
    Node node = storage.getSession(SessionProvider.createSystemProvider()).getNodeByUUID(uuid);
    if (node == null) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("No original recurrent node was found with UUID: " + uuid);
      }
      return null;
    } else {
      return node.getName();
    }
  }

  public void updateModifiedDateOfEvent(String username, CalendarEvent event) throws Exception {
    Node node = storage.getCalendarEventNode(username, event.getCalType(), event.getCalendarId(), event.getId());
    modifyUpdateDate(node);
    if (event.getOriginalReference() != null && !event.getOriginalReference().isEmpty()) {
      Node masterNode = storage.getSession(SessionProvider.createSystemProvider()).getNodeByUUID(event.getOriginalReference());
      modifyUpdateDate(masterNode);
    }
  }

  private void modifyUpdateDate(Node node) throws Exception {
    if (!node.isNodeType("exo:datetime")) {
      if (node.canAddMixin("exo:datetime")) {
        node.addMixin("exo:datetime");
      }
      node.setProperty("exo:dateCreated", new GregorianCalendar());
    }
    node.setProperty("exo:dateModified", new GregorianCalendar());
    node.save();
  }

  public CalendarEvent getEvent(Node eventNode) throws Exception {
    return storage.getEvent(eventNode);
  }
}