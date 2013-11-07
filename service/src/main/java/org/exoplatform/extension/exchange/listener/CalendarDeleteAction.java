package org.exoplatform.extension.exchange.listener;

import javax.jcr.Node;

import org.apache.commons.chain.Context;
import org.exoplatform.calendar.service.Utils;
import org.exoplatform.extension.exchange.service.IntegrationService;
import org.exoplatform.services.command.action.Action;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.security.ConversationState;
import org.exoplatform.services.security.IdentityConstants;

import com.ibm.icu.util.Calendar;

/**
 * 
 * @author Boubaker Khanfir
 * 
 */
public class CalendarDeleteAction implements Action {

  private final static Log LOG = ExoLogger.getLogger(CalendarDeleteAction.class);

  public boolean execute(Context context) throws Exception {
    Node node = (Node) context.get("currentItem");
    if (node.isNodeType("exo:calendarEvent")) {
      String eventId = node.getName();
      try {
        String userId = null;
        ConversationState state = ConversationState.getCurrent();
        if (state == null || state.getIdentity() == null || state.getIdentity().getUserId().equals(IdentityConstants.ANONIM)) {
          userId = node.getNode("../../../../..").getName();
        } else {
          userId = state.getIdentity().getUserId();
        }

        IntegrationService integrationService = IntegrationService.getInstance(userId);
        if (integrationService == null) {
          LOG.info("User '" + state.getIdentity().getUserId() + "' has no Exchange service, event will not be deleted from Exchange: eventId=" + eventId);
          return false;
        } else {
          try {
            if (!integrationService.isSynchronizationStarted()) {
              integrationService.setSynchronizationStarted();
              if (integrationService.getUserExoLastCheckDate() != null) {
                String calendarId = node.getProperty(Utils.CALENDAR_ID).getString();
                integrationService.deleteExchangeCalendarEvent(eventId, calendarId);
                integrationService.setUserExoLastCheckDate(Calendar.getInstance().getTime().getTime());
              }
              integrationService.setSynchronizationStopped();
            }
          } catch (Exception e) {
            LOG.error("Error while deleting Exchange event: " + eventId, e);
            // Integration is out of sync, so disable auto synchronization
            // until the scheduled job runs and try to fix this
            integrationService.setUserExoLastCheckDate(0);
          }
        }
      } catch (Exception e) {
        LOG.error("Error while deleting Exchange item corresponding event to eXo Event with Id: " + eventId, e);
      }
    }
    return false;
  }

}
