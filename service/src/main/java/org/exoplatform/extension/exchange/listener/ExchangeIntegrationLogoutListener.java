package org.exoplatform.extension.exchange.listener;

import org.exoplatform.services.listener.Event;
import org.exoplatform.services.listener.Listener;
import org.exoplatform.services.security.ConversationRegistry;
import org.exoplatform.services.security.ConversationState;

public class ExchangeIntegrationLogoutListener extends Listener<ConversationRegistry, ConversationState> {
  ExchangeListenerService exchangeListenerService;

  public ExchangeIntegrationLogoutListener(ExchangeListenerService exchangeListenerService) {
    this.exchangeListenerService = exchangeListenerService;
  }

  @Override
  public void onEvent(Event<ConversationRegistry, ConversationState> event) throws Exception {
    String username = event.getData().getIdentity().getUserId();
    exchangeListenerService.userLoggedOut(username);
  }
}
