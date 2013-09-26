package org.exoplatform.extension.exchange.listener;

import javax.security.auth.callback.Callback;
import javax.security.auth.callback.NameCallback;
import javax.security.auth.callback.PasswordCallback;
import javax.security.auth.login.LoginException;

import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.security.jaas.AbstractLoginModule;

public class ExchangeIntegrationLoginModule extends AbstractLoginModule {

  private static final Log LOG = ExoLogger.getLogger(ExchangeIntegrationLoginModule.class);

  ExchangeListenerService exchangeListenerService;

  public ExchangeIntegrationLoginModule() {
    try {
      this.exchangeListenerService  = (ExchangeListenerService) getContainer().getComponentInstanceOfType(ExchangeListenerService.class);
    } catch (Exception e) {
      LOG.error(e);
    }
  }

  @Override
  public boolean commit() throws LoginException {
    Callback[] callbacks = new Callback[2];
    callbacks[0] = new NameCallback("Username");
    callbacks[1] = new PasswordCallback("Password", false);
    try {
      callbackHandler.handle(callbacks);
      String username = ((NameCallback) callbacks[0]).getName();
      String password = new String(((PasswordCallback) callbacks[1]).getPassword());
      if (username == null || username.isEmpty()) {
        // Let other login modules handle this
        return true;
      }
      if (password == null || password.isEmpty()) {
        // Let other login modules handle this
        return true;
      }
      exchangeListenerService.userLoggedIn(username, password);
    } catch (Exception e) {
      LOG.error(e);
    }
    // Let other login modules run
    return true;
  }

  @Override
  public boolean login() throws LoginException {
    return true;
  }

  @Override
  public boolean abort() throws LoginException {
    return true;
  }

  @Override
  public boolean logout() throws LoginException {
    return false;
  }

  @Override
  protected Log getLogger() {
    return LOG;
  }
}
