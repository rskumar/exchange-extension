package org.exoplatform.extension.exchange.listener;

import javax.jcr.Node;
import javax.jcr.Property;

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
public class CalendarCreateUpdateAction implements Action {

  private final static Log LOG = ExoLogger.getLogger(CalendarCreateUpdateAction.class);

  public boolean execute(Context context) throws Exception {
    Object object = context.get("currentItem");
    Node node = null;
    if (object instanceof Node) {
      node = (Node) object;
    } else if (object instanceof Property) {
      Property property = (Property) object;
      node = property.getParent();
    }
    if (node != null && node.isNodeType("exo:calendarEvent")) {
      String eventId = node.getName();
      try {
        String userId = null;
        ConversationState state = ConversationState.getCurrent();
        if (state == null || state.getIdentity() == null || state.getIdentity().getUserId().equals(IdentityConstants.ANONIM)) {
          userId = node.getNode("../../../../..").getName();
        } else {
          userId = state.getIdentity().getUserId();
        }

        if (userId == null) {
          LOG.warn("No user was found while trying to create/update eXo Calendar event with id: " + eventId);
          return false;
        }

        IntegrationService integrationService = IntegrationService.getInstance(userId);
        if (integrationService == null) {
          LOG.warn("No authenticated user was found while trying to create/update eXo Calendar event with id: '" + eventId + "' for user: " + userId);
          return false;
        } else {
          try {
            if (!integrationService.isSynchronizationStarted()) {
              integrationService.setSynchronizationStarted();
              if (isNodeValid(node) && integrationService.getUserExoLastCheckDate() != null) {
                integrationService.updateOrCreateExchangeCalendarEvent(node);
                integrationService.setUserExoLastCheckDate(Calendar.getInstance().getTime().getTime());
              }
              integrationService.setSynchronizationStopped();
            }
          } catch (Exception e) {
            // This can happen if the node was newly created, so not all
            // properties are in the node
            LOG.error("Error while create/update an Exchange item for eXo event: " + eventId, e);
            // Integration is out of sync, so disable auto synchronization
            // until the scheduled job runs and try to fix this
            integrationService.setUserExoLastCheckDate(0);
          }
        }
      } catch (Exception e) {
        LOG.error("Error while updating Exchange with the eXo Event with Id: " + eventId, e);
      }
    }
    return false;
  }

  private boolean isNodeValid(Node node) throws Exception {
    return node.hasProperty(Utils.EXO_PARTICIPANT_STATUS);
  }

}
