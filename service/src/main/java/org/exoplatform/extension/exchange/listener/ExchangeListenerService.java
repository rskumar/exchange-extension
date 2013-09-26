package org.exoplatform.extension.exchange.listener;

import java.net.URI;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;

import microsoft.exchange.webservices.data.EventType;
import microsoft.exchange.webservices.data.ExchangeCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.ExchangeVersion;
import microsoft.exchange.webservices.data.FolderEvent;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.GetEventsResults;
import microsoft.exchange.webservices.data.ItemEvent;
import microsoft.exchange.webservices.data.PullSubscription;
import microsoft.exchange.webservices.data.WebCredentials;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.container.PortalContainer;
import org.exoplatform.container.component.ComponentRequestLifecycle;
import org.exoplatform.container.xml.InitParams;
import org.exoplatform.extension.exchange.service.CalendarStorageService;
import org.exoplatform.extension.exchange.service.ExchangeIntegrationService;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;
import org.exoplatform.services.organization.UserProfile;
import org.exoplatform.services.security.Identity;
import org.exoplatform.services.security.IdentityRegistry;
import org.picocontainer.Startable;

public class ExchangeListenerService implements Startable {

  private final static Log LOG = ExoLogger.getLogger(ExchangeListenerService.class);

  private static final String EXCHANGE_SERVER_URL_PARAM_NAME = "exchange.ews.url";
  private static final String EXCHANGE_DOMAIN_PARAM_NAME = "exchange.domain";
  // TODO This param have to be computed automatically between servers timezones
  // and calendar timezones
  private static final String EXCHANGE_DIFF_TIME_NAME = "exchange.diff.time";
  private static final String USER_EXCHANGE_HANDLED_ATTRIBUTE = "exchange.check.date";

  private static int threadIndex = 0;

  private static int diffTimeZone = 0;

  private final ScheduledExecutorService scheduledExecutor = Executors.newScheduledThreadPool(10);
  private final ExecutorService threadExecutor = Executors.newFixedThreadPool(10);
  private final Map<String, ScheduledFuture<?>> futures = new HashMap<String, ScheduledFuture<?>>();

  private final String exchangeServerURL;
  private final String exchangeDomain;

  private final CalendarStorageService calendarStorageService;
  private final IdentityRegistry identityRegistry;
  private final OrganizationService organizationService;

  public ExchangeListenerService(OrganizationService organizationService, CalendarStorageService calendarStorageService, IdentityRegistry identityRegistry, InitParams params) {
    this.calendarStorageService = calendarStorageService;
    this.identityRegistry = identityRegistry;
    this.organizationService = organizationService;

    if (!params.containsKey(EXCHANGE_SERVER_URL_PARAM_NAME)) {
      throw new IllegalStateException("Please add 'exchange.ews.url' parameter in configuration.properties.");
    } else {
      exchangeServerURL = params.getValueParam(EXCHANGE_SERVER_URL_PARAM_NAME).getValue();
    }
    if (params.containsKey(EXCHANGE_DIFF_TIME_NAME)) {
      String diffTimeZoneString = params.getValueParam(EXCHANGE_DIFF_TIME_NAME).getValue();
      diffTimeZone = Integer.valueOf(diffTimeZoneString);
    }
    if (!params.containsKey(EXCHANGE_DOMAIN_PARAM_NAME)) {
      throw new IllegalStateException("Please add 'exchange.domain' parameter in configuration.properties.");
    } else {
      exchangeDomain = params.getValueParam(EXCHANGE_DOMAIN_PARAM_NAME).getValue();
    }
    LOG.info("Successfully started.");
  }

  @Override
  public void start() {
  }

  @Override
  public void stop() {
    scheduledExecutor.shutdownNow();
    threadExecutor.shutdownNow();
  }

  public void userLoggedIn(final String username, final String password) {
    try {
      Identity identity = identityRegistry.getIdentity(username);
      if (identity == null) {
        throw new IllegalStateException("Identity of user '" + username + "' not found.");
      }

      // Execute the synchronization once logged in
      Thread command = new ExchangeIntegrationTask(username, password, true);
      try {
        threadExecutor.execute(command);
      } catch (Exception e) {
        // e.printStackTrace();
      }

      // Scheduled task: listen the changes made on MS Exchange Calendar
      command = new ExchangeIntegrationTask(username, password, false);
      ScheduledFuture<?> future = scheduledExecutor.scheduleWithFixedDelay(command, 30, 20, TimeUnit.SECONDS);

      // Add future task to the map to destroy thread when the user logout
      futures.put(username, future);

      LOG.info("User '" + username + "' logged in, exchange synchronization task started.");
    } catch (Exception e) {
      LOG.error("Can't synchronize with exchange data.", e);
    }
  }

  public void userLoggedOut(String username) {
    ScheduledFuture<?> future = futures.get(username);
    if (future == null) {
      LOG.warn("User '" + username + "' logged out, exchange synchronization task wasn't found.");
    } else {
      future.cancel(true);
      LOG.info("User '" + username + "' logged out, exchange synchronization task stopped.");
    }
  }

  public Date getUserLastCheckDate(String username) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
    long time = userProfile.getAttribute(USER_EXCHANGE_HANDLED_ATTRIBUTE) == null ? 0 : Long.valueOf(userProfile.getAttribute(USER_EXCHANGE_HANDLED_ATTRIBUTE));
    Date lastSyncDate = null;
    if (time > 0) {
      lastSyncDate = new Date(time);
    }
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
    }
    return lastSyncDate;
  }

  public void setUserLastCheckDate(String username, long time) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
    userProfile.setAttribute(USER_EXCHANGE_HANDLED_ATTRIBUTE, "" + time);
    organizationService.getUserProfileHandler().saveUserProfile(userProfile, false);
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
    }
  }

  protected class ExchangeIntegrationTask extends Thread {
    private final ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
    private final ExchangeIntegrationService integrationrService;
    private final List<FolderId> calendarFolderIds;
    private PullSubscription subscription = null;
    private String username;
    private boolean lazy;

    public ExchangeIntegrationTask(String username, String password, boolean lazy) throws Exception {
      super("ExchangeIntegrationTask-" + (threadIndex++));
      this.username = username;
      this.lazy = lazy;

      ExchangeCredentials credentials = new WebCredentials(username + "@" + exchangeDomain, password);
      service.setCredentials(credentials);
      service.setUrl(new URI(exchangeServerURL));

      integrationrService = new ExchangeIntegrationService(calendarStorageService, service, username);

      try {
        calendarFolderIds = integrationrService.getExchangeCalendars();
      } catch (Exception e) {
        if (lazy) {
          throw new RuntimeException("Error while authenticating user '" + username + "' to exchange, please make sure you are connected to the correct URL with correct credentials.", e);
        } else {
          throw e;
        }
      }
      if (!lazy) {
        newSubscription();
      }
    }

    @Override
    public void run() {
      try {
        // This is used once, when user login
        if (lazy) {
          LOG.info("run first synchronization for user: " + username);
          Date lastSyncDate = getUserLastCheckDate(username);
          for (FolderId folderId : calendarFolderIds) {
            Calendar calendar = calendarStorageService.findUserCalendar(username, folderId.getUniqueId());
            if (calendar == null || lastSyncDate == null) {
              integrationrService.synchronizeFullCalendar(folderId);
            } else {
              integrationrService.synchronizeModificationsOfCalendar(folderId, lastSyncDate, diffTimeZone);
            }
          }
        } else {
          LOG.info("run scheduled synchronization for user: " + username);
          // This is used in a scheduled task when the user session still alive
          GetEventsResults events;
          try {
            events = subscription.getEvents();
          } catch (Exception e) {
            LOG.warn("subscription seems ended: " + e.getMessage() + ", retry...");

            newSubscription();
            events = subscription.getEvents();
          }

          Iterable<ItemEvent> itemEvents = events.getItemEvents();
          // If new Calendars was added
          boolean folderCreated = false;
          if (events.getFolderEvents() != null && events.getFolderEvents().iterator().hasNext()) {
            FolderEvent folderEvent = events.getFolderEvents().iterator().next();
            if (folderEvent.getEventType().equals(EventType.Created) || folderEvent.getEventType().equals(EventType.Modified)) {
              if (!integrationrService.isCalendarPresent(username, folderEvent.getFolderId().getUniqueId())) {
                integrationrService.synchronizeFullCalendar(folderEvent.getFolderId());
                if (!calendarFolderIds.contains(folderEvent.getFolderId())) {
                  calendarFolderIds.add(folderEvent.getFolderId());
                }
                folderCreated = true;
              }
            }
          }
          // loop through Appointment events
          for (ItemEvent itemEvent : itemEvents) {
            integrationrService.createOrUpdateOrDelete(itemEvent);
          }

          // Update listened folders
          if (folderCreated) {
            newSubscription();
          }
        }

        // Update date of last check in a user profile attribute
        setUserLastCheckDate(username, java.util.Calendar.getInstance().getTimeInMillis());
      } catch (Exception e) {
        LOG.error("Error while synchronizing calndar entries.", e);
      }
    }

    @Override
    public void interrupt() {
      if (subscription != null) {
        try {
          LOG.info("Thread interruption: unsubscribe user service:" + username);
          subscription.unsubscribe();
        } catch (Exception e) {
          LOG.error("Thread interruption: Error while unsubscribe to thread of user:" + username);
        }
      }
      super.interrupt();
    }

    @Override
    protected void finalize() throws Throwable {
      if (subscription != null) {
        try {
          LOG.info("Object finalization: unsubscribe user service:" + username);
          subscription.unsubscribe();
        } catch (Exception e) {
          LOG.error("Object finalization: Error while unsubscribe to thread of user:" + username);
        }
      }
      super.finalize();
    }

    private void newSubscription() throws Exception {
      subscription = service.subscribeToPullNotifications(calendarFolderIds, 5, null, EventType.Modified, EventType.Moved, EventType.FreeBusyChanged, EventType.Created, EventType.Deleted);
    }

  }
}
