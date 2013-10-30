package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.Iterator;
import java.util.TimeZone;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.ConflictResolutionMode;
import microsoft.exchange.webservices.data.DeleteMode;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.ItemId;
import microsoft.exchange.webservices.data.ServiceResponseException;
import microsoft.exchange.webservices.data.TimeZoneDefinition;

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

  public void updateOrCreateExchangeAppointment(String username, ExchangeService service, CalendarEvent event, TimeZone userCalendarTimeZone) throws Exception {
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

    if (appointment == null) {
      // Checks if this event was already in Exchange, if it's the case, it
      // means that the item was not found because the user has removed it from
      // Exchange
      if (CalendarConverterService.isExchangeEventId(event.getId())) {
        LOG.error("Conflict in modification, inconsistant data, the event was deleted in Exchange but seems always in eXo, the event will be deleted from Exchange.");
        deleteExchangeAppointment(username, service, event);
        return;
      }
      appointment = new Appointment(service);
    }

    if (event.getRepeatType() == null || event.getRepeatType().equals(CalendarEvent.RP_NOREPEAT) || event.getRecurrenceId() == null) {
      CalendarConverterService.convertExoToExchangeEvent(appointment, event, username, organizationService.getUserHandler(), getTimeZoneDefinition(service, userCalendarTimeZone), userCalendarTimeZone);
    } else {
      if (event.getIsExceptionOccurrence() || event.getRecurrenceId() != null) {
        // TODO make this exception occurence
        CalendarConverterService.convertExoToExchangeOccurenceEvent(appointment, event, username, organizationService.getUserHandler(), getTimeZoneDefinition(service, userCalendarTimeZone), userCalendarTimeZone);
      } else {
        CalendarConverterService.convertExoToExchangeMasterRecurringCalendarEvent(appointment, event, username, organizationService.getUserHandler(),
            getTimeZoneDefinition(service, userCalendarTimeZone), userCalendarTimeZone);
      }
    }
    String folderIdString = correspondenceService.getCorrespondingId(username, event.getCalendarId());
    if (folderIdString == null || folderIdString.isEmpty()) {
      throw new IllegalStateException("No Folder Id was found that corresponds to calendar id: " + folderIdString);
    }
    if (isNew) {
      FolderId folderId = FolderId.getFolderIdFromString(folderIdString);
      appointment.save(folderId);
    } else {
      appointment.update(ConflictResolutionMode.AlwaysOverwrite);
    }

    correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
  }

  public void deleteExchangeAppointment(String username, ExchangeService service, CalendarEvent calendarEvent) throws Exception {
    String itemId = correspondenceService.getCorrespondingId(username, calendarEvent.getId());
    if (itemId == null) {
      LOG.warn("Conflict in modification, inconsistant data, the event was deleted from eXo but seems don't have corresponding Event in Exchange, ignore.");
    } else {
      Appointment appointment = null;
      try {
        appointment = Appointment.bind(service, ItemId.getItemIdFromString(itemId));
        appointment.delete(DeleteMode.HardDelete);
      } catch (ServiceResponseException e) {
        if (LOG.isTraceEnabled()) {
          LOG.trace("Item was not bound, it was deleted or not yet created:" + calendarEvent.getId());
        }
      }
      correspondenceService.deleteCorrespondingId(username, itemId, calendarEvent.getId());
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
  
}