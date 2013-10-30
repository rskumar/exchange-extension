package org.exoplatform.extension.exchange.listener;

import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.extension.exchange.service.IntegrationService;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.security.ConversationState;
import org.exoplatform.services.security.IdentityConstants;

public class CalendarEventListener extends org.exoplatform.calendar.service.impl.CalendarEventListener {

  private final static Log LOG = ExoLogger.getLogger(CalendarEventListener.class);

  @Override
  public void deletePublicEvent(CalendarEvent event, String calendarId) {
    ConversationState state = ConversationState.getCurrent();
    if (state == null || state.getIdentity() == null || state.getIdentity().getUserId().equals(IdentityConstants.ANONIM)) {
      LOG.info("No user was found while trying to synchronize with Exchange, eventId: " + event.getId());
    }

    IntegrationService integrationService = IntegrationService.getInstance(state.getIdentity().getUserId());
    if (integrationService == null) {
      LOG.warn("No Exchange integration service was found for authenticated user:" + state.getIdentity().getUserId() + " , for eXo Calendar eventId: " + event.getId());
    }
    try {
      integrationService.deleteExchangeCalendarEvent(event);
    } catch (Exception e) {
      LOG.error("Error while deleting Exchange event: " + event.getId(), e);
    }
  }

  @Override
  public void updatePublicEvent(CalendarEvent event, String calendarId) {
    ConversationState state = ConversationState.getCurrent();
    if (state == null || state.getIdentity() == null || state.getIdentity().getUserId().equals(IdentityConstants.ANONIM)) {
      LOG.info("No user was found while trying to synchronize with Exchange, eventId: " + event.getId());
    }

    IntegrationService integrationService = IntegrationService.getInstance(state.getIdentity().getUserId());
    if (integrationService == null) {
      LOG.warn("No Exchange integration service was found for authenticated user:" + state.getIdentity().getUserId() + " , for eXo Calendar eventId: " + event.getId());
    }
    try {
      integrationService.updateOrCreateExchangeCalendarEvent(event);
    } catch (Exception e) {
      LOG.error("Error while updating Exchange event: " + event.getId(), e);
    }
  }

  @Override
  public void updatePublicEvent(CalendarEvent oldEvent, CalendarEvent event, String calendarId) {
    ConversationState state = ConversationState.getCurrent();
    if (state == null || state.getIdentity() == null || state.getIdentity().getUserId().equals(IdentityConstants.ANONIM)) {
      LOG.info("No user was found while trying to synchronize with Exchange, eventId: " + event.getId());
    }

    IntegrationService integrationService = IntegrationService.getInstance(state.getIdentity().getUserId());
    if (integrationService == null) {
      LOG.warn("No Exchange integration service was found for authenticated user:" + state.getIdentity().getUserId() + " , for eXo Calendar eventId: " + event.getId());
    }
    try {
      integrationService.updateOrCreateExchangeCalendarEvent(event);
    } catch (Exception e) {
      LOG.error("Error while updating Exchange event: " + event.getId(), e);
    }
  }

  @Override
  public void savePublicEvent(CalendarEvent event, String calendarId) {
    ConversationState state = ConversationState.getCurrent();
    if (state == null || state.getIdentity() == null || state.getIdentity().getUserId().equals(IdentityConstants.ANONIM)) {
      LOG.info("No user was found while trying to synchronize with Exchange, eventId: " + event.getId());
    }

    IntegrationService integrationService = IntegrationService.getInstance(state.getIdentity().getUserId());
    if (integrationService == null) {
      LOG.warn("No Exchange integration service was found for authenticated user:" + state.getIdentity().getUserId() + " , for eXo Calendar eventId: " + event.getId());
    }
    try {
      integrationService.updateOrCreateExchangeCalendarEvent(event);
    } catch (Exception e) {
      LOG.error("Error while updating Exchange event: " + event.getId(), e);
    }
  }

}
